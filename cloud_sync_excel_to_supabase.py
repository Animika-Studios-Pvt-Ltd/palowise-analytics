"""
Cloud Sync Worker: OneDrive Excel to Supabase Cloud PostgreSQL
==============================================================
Palowise Cloud Automation Worker

Workflow:
  1. Retrieves latest Excel workbook from OneDrive via anonymous sharing link
     (or fallback to local mirror for workstation testing).
  2. In Dry-Run mode (Default):
     - Validates workbook and business rules.
     - Connects to Supabase in a strictly read-only session.
     - Simulates exact row diffs (unchanged, updated, new, preserved).
     - Verifies downstream view public.demo_timing.
     - Performs ZERO database writes.
  3. In Live mode (--live):
     - Evaluates SHA-256 fingerprint against Supabase sync state table.
       Skips processing if workbook is unchanged (idempotent, compute-saving).
     - Applies rigorous Palowise business validations:
       * Excludes dummy test rows ('Test').
       * Excludes blank rows / missing lead names.
       * Deduplicates composite keys: lower(trim(name)) + lower(trim(industry)).
       * Validates and parses ordinal date formats (e.g., '17th May').
       * Standardizes flag columns ('YES' / 'NO').
     - Executes atomic upsert into Supabase 'public.marketing_analysis_clean' via
       temporary staging table with 'IS DISTINCT FROM' diffing.
     - Verifies downstream analytical view 'public.demo_timing'.
     - Updates persistent sync state in Supabase.
     - Commits atomically or rolls back on any failure.

Modes:
  --dry-run       (Default) Validate workbook and simulate database merge without writing.
  --live          Execute live idempotent synchronization into Supabase.
  --force         Bypass SHA-256 state check and force validation/sync.
  --local-file    Specify a local file path instead of downloading via OneDrive sharing link.
  --share-url     Override OneDrive sharing URL.
"""

import argparse
import hashlib
import io
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
import openpyxl
import pandas as pd
import psycopg2
from psycopg2 import extras

# Reuse the proven, tested download engine and sanitizers from test_onedrive_download.py
from test_onedrive_download import (
    download_file_bytes,
    validate_excel_signature,
    sanitize_text,
)

# ==============================================================================
# Business Rules & Schema Mappings (100% Identical to Production Local Sync)
# ==============================================================================

COLUMN_MAPPING = {
    "NAME": "name",
    "DESIGNATION": "designation",
    "INDUSTRY": "industry",
    "STARTING DATE": "starting_date",
    "Demo Date": "demo_date",
    "Pilot Conducted": "pilot_conducted",
    "Proposals Submitted": "proposals_submitted",
    "Follow Up": "follow_up",
    "STATUS": "status",
    "Remarks": "remarks",
    "Sentiment": "sentiment",
    "REASON": "reason",
    "Feedback Theme": "feedback_theme",
    "priority": "priority",
    "Pipeline Stage": "pipeline_stage",
    "NEXT ACTION": "next_action",
    "Compony": "company",
    "Level": "level",
    "Des_role": "des_role",
    "Source": "source",
    "Demo Booked": "demo_booked",
    "Demo Conducted": "demo_conducted",
    "Specification Form Sent": "specification_form_sent",
    "Proposal Sent": "proposal_sent",
    "Pilot Conducted2": "pilot_conducted2",
    "Not Interested": "not_interested",
    "Reach Out Later": "reach_out_later",
    "status2": "status2",
}

FLAG_COLUMNS = {
    "pilot_conducted",
    "proposals_submitted",
    "demo_booked",
    "demo_conducted",
    "specification_form_sent",
    "proposal_sent",
    "pilot_conducted2",
    "not_interested",
    "reach_out_later",
}

NON_KEY_COLUMNS = [
    "designation",
    "industry",
    "starting_date",
    "demo_date",
    "pilot_conducted",
    "proposals_submitted",
    "follow_up",
    "status",
    "remarks",
    "sentiment",
    "reason",
    "feedback_theme",
    "priority",
    "pipeline_stage",
    "next_action",
    "company",
    "level",
    "des_role",
    "source",
    "demo_booked",
    "demo_conducted",
    "specification_form_sent",
    "proposal_sent",
    "pilot_conducted2",
    "not_interested",
    "reach_out_later",
    "status2",
]

ALL_COLUMNS = list(COLUMN_MAPPING.values())
VALID_MONTHS = {"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"}

# ==============================================================================
# Helper Functions: Formatting, Validation, Security
# ==============================================================================

def sanitize_message(msg: str, secret_list: list[str] | None = None) -> str:
    """Sanitizes sensitive tokens, passwords, and URI patterns from messages."""
    if not msg:
        return ""
    clean = str(msg)
    if secret_list:
        for sec in secret_list:
            if sec and len(sec) > 3:
                clean = clean.replace(sec, "[REDACTED]")
    clean = re.sub(r":([^:@\s]+)@", r":[REDACTED]@", clean)
    clean = sanitize_text(clean)
    return clean


def compute_bytes_sha256(data: bytes) -> str:
    """Computes SHA-256 hash of raw bytes."""
    h = hashlib.sha256()
    h.update(data)
    return h.hexdigest()


def clean_cell_value(val) -> str | None:
    if pd.isna(val) or val is None:
        return None
    val_str = str(val).strip()
    return val_str if val_str else None


def standardize_flag(val: str | None) -> str | None:
    if val is None:
        return None
    cleaned = str(val).strip()
    if not cleaned:
        return None
    lower_val = cleaned.lower()
    if lower_val in ("yes", "y", "true"):
        return "YES"
    elif lower_val in ("no", "n", "false"):
        return "NO"
    return cleaned


def validate_date_format(val: str | None) -> tuple[bool, str]:
    if val is None or val == "":
        return True, "NULL/Blank"
    m = re.match(r"^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,})$", str(val).strip(), re.IGNORECASE)
    if not m:
        return False, f"Invalid date format '{val}' (expected e.g. '17th May')"
    day_str, month_str = m.groups()
    if month_str[:3].lower() not in VALID_MONTHS:
        return False, f"Unrecognized month '{month_str}' in '{val}'"
    day_int = int(day_str)
    if not (1 <= day_int <= 31):
        return False, f"Invalid day '{day_int}' in '{val}'"
    return True, "Valid"


# ==============================================================================
# Configuration Loading
# ==============================================================================

def load_cloud_config() -> dict:
    """
    Loads configuration from environment variables (GitHub Secrets or .env.supabase).
    """
    env_file = Path(__file__).parent / ".env.supabase"
    file_cfg = {}
    if env_file.is_file():
        with open(env_file, "r", encoding="utf-8") as f:
            for line in f:
                stripped = line.strip()
                if stripped and not stripped.startswith("#") and "=" in stripped:
                    k, v = stripped.split("=", 1)
                    file_cfg[k.strip()] = v.strip().strip("'\"")

    return {
        # Supabase Settings
        "SUPABASE_HOST": os.environ.get("SUPABASE_HOST") or file_cfg.get("SUPABASE_HOST", ""),
        "SUPABASE_PORT": os.environ.get("SUPABASE_PORT") or file_cfg.get("SUPABASE_PORT", "5432"),
        "SUPABASE_DATABASE": (
            os.environ.get("SUPABASE_DATABASE")
            or os.environ.get("SUPABASE_DB")
            or file_cfg.get("SUPABASE_DATABASE")
            or file_cfg.get("SUPABASE_DB", "postgres")
        ),
        "SUPABASE_USER": os.environ.get("SUPABASE_USER") or file_cfg.get("SUPABASE_USER", ""),
        "SUPABASE_PASSWORD": os.environ.get("SUPABASE_PASSWORD") or file_cfg.get("SUPABASE_PASSWORD", ""),
        "SUPABASE_SSLMODE": os.environ.get("SUPABASE_SSLMODE") or file_cfg.get("SUPABASE_SSLMODE", "require"),
        # OneDrive Sharing Link
        "ONEDRIVE_SHARE_URL": os.environ.get("ONEDRIVE_SHARE_URL") or file_cfg.get("ONEDRIVE_SHARE_URL", ""),
        # Optional Microsoft Graph Settings (fallback)
        "MS_GRAPH_TENANT_ID": os.environ.get("MS_GRAPH_TENANT_ID", ""),
        "MS_GRAPH_CLIENT_ID": os.environ.get("MS_GRAPH_CLIENT_ID", ""),
        "MS_GRAPH_CLIENT_SECRET": os.environ.get("MS_GRAPH_CLIENT_SECRET", ""),
        "ONEDRIVE_USER_PRINCIPAL_NAME": os.environ.get("ONEDRIVE_USER_PRINCIPAL_NAME", ""),
        "ONEDRIVE_FILE_PATH": os.environ.get("ONEDRIVE_FILE_PATH", "lumos10.xlsx"),
    }


# ==============================================================================
# OneDrive Download Engine
# ==============================================================================

def get_graph_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Requests OAuth 2.0 token from Microsoft Entra ID (Graph API fallback)."""
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    data = urllib.parse.urlencode(payload).encode("utf-8")
    req = urllib.request.Request(
        token_url,
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        body = json.loads(resp.read().decode("utf-8"))
        return body["access_token"]


def download_onedrive_workbook_graph(cfg: dict) -> tuple[bytes, str]:
    """Downloads workbook via Microsoft Graph API (if credentials present)."""
    tenant_id = cfg.get("MS_GRAPH_TENANT_ID")
    client_id = cfg.get("MS_GRAPH_CLIENT_ID")
    client_secret = cfg.get("MS_GRAPH_CLIENT_SECRET")
    upn = cfg.get("ONEDRIVE_USER_PRINCIPAL_NAME")
    file_path = cfg.get("ONEDRIVE_FILE_PATH", "lumos10.xlsx").strip("/")

    access_token = get_graph_access_token(tenant_id, client_id, client_secret)
    encoded_path = urllib.parse.quote(file_path)
    url = f"https://graph.microsoft.com/v1.0/users/{upn}/drive/root:/{encoded_path}:/content"

    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        file_bytes = resp.read()
        return file_bytes, f"Microsoft Graph ({file_path})"


def resolve_excel_bytes(cfg: dict, local_file: Path | None = None, share_url: str | None = None) -> tuple[bytes, str]:
    """
    Retrieves Excel workbook bytes:
      1. Explicit local file (--local-file)
      2. OneDrive sharing link (ONEDRIVE_SHARE_URL) via robust download engine
      3. Microsoft Graph API fallback (if credentials configured)
      4. Workstation local mirror fallback (for offline developer testing)
    """
    if local_file:
        if not local_file.is_file():
            raise FileNotFoundError(f"Local file not found at: {local_file}")
        with open(local_file, "rb") as f:
            data = f.read()
        return data, f"Local File: {local_file.name}"

    effective_share_url = share_url or cfg.get("ONEDRIVE_SHARE_URL")
    if effective_share_url and effective_share_url.strip():
        print("Downloading latest Excel workbook from OneDrive sharing link...")
        data = download_file_bytes(effective_share_url.strip())
        return data, "OneDrive Sharing Link (lumos10.xlsx)"

    # Fallback: Microsoft Graph
    has_graph = all([
        cfg.get("MS_GRAPH_TENANT_ID"),
        cfg.get("MS_GRAPH_CLIENT_ID"),
        cfg.get("MS_GRAPH_CLIENT_SECRET"),
        cfg.get("ONEDRIVE_USER_PRINCIPAL_NAME"),
    ])
    if has_graph:
        return download_onedrive_workbook_graph(cfg)

    # Fallback: Workstation mirror
    local_mirror = Path(r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\lumos10.xlsx")
    if local_mirror.is_file():
        print(f"Notice: ONEDRIVE_SHARE_URL not configured. Using workstation mirror: {local_mirror.name}")
        with open(local_mirror, "rb") as f:
            data = f.read()
        return data, f"Workstation Mirror ({local_mirror.name})"

    raise ValueError(
        "No Excel source found. Please configure ONEDRIVE_SHARE_URL in GitHub Secrets / environment, "
        "or pass --local-file <path>."
    )


# ==============================================================================
# Workbook Processing & Business Rule Engine
# ==============================================================================

def process_workbook(file_source: bytes | Path) -> tuple[int, list, list, list]:
    """
    Parses workbook, applies validation, filters dummy rows, validates date formats,
    and resolves composite key duplicates (latest row supersedes prior row).
    Returns: (total_raw_rows, eligible_records, excluded_records, validation_errors)
    """
    if isinstance(file_source, bytes):
        stream = io.BytesIO(file_source)
        df = pd.read_excel(stream, sheet_name="in", engine="openpyxl")
    else:
        df = pd.read_excel(file_source, sheet_name="in", engine="openpyxl")

    total_raw_rows = len(df)
    eligible_records = []
    excluded_records = []
    validation_errors = []
    records_by_key = {}

    for idx, row in df.iterrows():
        row_number = idx + 2
        raw_name = clean_cell_value(row.get("NAME"))

        # Rule 1: Exclude dummy test row (NAME == "Test")
        if raw_name and raw_name.lower() == "test":
            excluded_records.append({
                "row_number": row_number,
                "name": raw_name,
                "reason": "Excluded dummy test record (NAME == 'Test')",
            })
            continue

        # Rule 2: Exclude blank row or missing lead NAME
        if not raw_name:
            if row.isna().all():
                excluded_records.append({
                    "row_number": row_number,
                    "name": None,
                    "reason": "Excluded completely blank row",
                })
            else:
                excluded_records.append({
                    "row_number": row_number,
                    "name": None,
                    "reason": "Excluded row with missing lead NAME",
                })
            continue

        record = {}
        row_has_error = False

        for excel_col, pg_col in COLUMN_MAPPING.items():
            cell_val = clean_cell_value(row.get(excel_col))

            if pg_col in FLAG_COLUMNS and cell_val is not None:
                cell_val = standardize_flag(cell_val)

            if pg_col in ("starting_date", "demo_date") and cell_val is not None:
                is_valid, msg = validate_date_format(cell_val)
                if not is_valid:
                    validation_errors.append({
                        "row_number": row_number,
                        "lead_name": raw_name,
                        "field": pg_col,
                        "error": msg,
                    })
                    row_has_error = True

            record[pg_col] = cell_val

        if row_has_error:
            continue

        # Rule 3: Deduplicate composite key in Source Excel: lower(trim(name)) + lower(trim(industry))
        norm_key = (
            record["name"].lower().strip(),
            (record["industry"] or "").lower().strip(),
        )
        if norm_key in records_by_key:
            prior_row, prior_rec = records_by_key[norm_key]
            excluded_records.append({
                "row_number": prior_row,
                "name": prior_rec["name"],
                "reason": f"Superseded by later updated Excel entry on row #{row_number}",
            })
        records_by_key[norm_key] = (row_number, record)

    eligible_records = list(records_by_key.values())
    return total_raw_rows, eligible_records, excluded_records, validation_errors


# ==============================================================================
# Persistent Cloud State Engine (Supabase-backed)
# ==============================================================================

def init_cloud_state_table(conn):
    """Ensures the cloud sync state table exists in Supabase."""
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS public._palowise_sync_state (
                id int PRIMARY KEY DEFAULT 1,
                sha256 text NOT NULL,
                file_size bigint,
                last_synced_at timestamptz DEFAULT now(),
                rows_in_db int,
                inserted_count int,
                updated_count int,
                preserved_count int,
                CONSTRAINT single_row CHECK (id = 1)
            );
        """)


def get_cloud_state(conn) -> dict | None:
    """Retrieves the last sync metadata from Supabase."""
    init_cloud_state_table(conn)
    with conn.cursor() as cur:
        cur.execute("""
            SELECT sha256, file_size, last_synced_at, rows_in_db 
            FROM public._palowise_sync_state 
            WHERE id = 1;
        """)
        row = cur.fetchone()
        if row:
            return {
                "sha256": row[0],
                "file_size": row[1],
                "last_synced_at": row[2],
                "rows_in_db": row[3],
            }
        return None


def save_cloud_state(conn, sha256: str, file_size: int, rows_in_db: int, inserted: int, updated: int, preserved: int):
    """Upserts the sync state in Supabase."""
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO public._palowise_sync_state (id, sha256, file_size, last_synced_at, rows_in_db, inserted_count, updated_count, preserved_count)
            VALUES (1, %s, %s, now(), %s, %s, %s, %s)
            ON CONFLICT (id) DO UPDATE SET
                sha256 = EXCLUDED.sha256,
                file_size = EXCLUDED.file_size,
                last_synced_at = now(),
                rows_in_db = EXCLUDED.rows_in_db,
                inserted_count = EXCLUDED.inserted_count,
                updated_count = EXCLUDED.updated_count,
                preserved_count = EXCLUDED.preserved_count;
        """, (sha256, file_size, rows_in_db, inserted, updated, preserved))


# ==============================================================================
# Database Connection & Execution Modes
# ==============================================================================

def connect_supabase(cfg: dict, readonly: bool = False):
    """Connects to Supabase PostgreSQL enforcing SSL."""
    conn = psycopg2.connect(
        host=cfg["SUPABASE_HOST"],
        port=int(cfg.get("SUPABASE_PORT", 5432)),
        dbname=cfg.get("SUPABASE_DATABASE", "postgres"),
        user=cfg["SUPABASE_USER"],
        password=cfg.get("SUPABASE_PASSWORD", ""),
        sslmode=cfg.get("SUPABASE_SSLMODE", "require"),
        connect_timeout=20,
    )
    if readonly:
        conn.set_session(readonly=True)
    return conn


def execute_dry_run(file_bytes: bytes, source_name: str, cfg: dict) -> bool:
    """
    Executes a strictly read-only simulation against Supabase:
      - Validates workbook and business rules
      - Queries current Supabase database state in a read-only transaction
      - Calculates exact counts for unchanged, updated, new, and preserved records
      - Verifies downstream view public.demo_timing
      - Executes ZERO database modifications (no INSERT, UPDATE, DELETE, DDL)
    """
    file_size = len(file_bytes)
    current_hash = compute_bytes_sha256(file_bytes)

    print("=" * 70)
    print("PALOWISE CLOUD SYNC: DRY-RUN SIMULATION (READ-ONLY)")
    print("=" * 70)
    print(f"Source:                    {source_name} ({file_size:,} bytes)")
    print(f"SHA-256:                   {current_hash}")
    print(f"Target Database:           Supabase ({cfg['SUPABASE_HOST']}:{cfg['SUPABASE_PORT']})")
    print(f"Database Name:             {cfg['SUPABASE_DATABASE']}")
    print(f"SSL Mode:                  {cfg.get('SUPABASE_SSLMODE', 'require')} (Enforced)")
    print(f"Mode:                      DRY-RUN (0 database writes; read-only session)")
    print("-" * 70)

    # 1. Workbook Processing & Validation
    print("[1] Workbook Validation & Business Rules:")
    total_raw, eligible, excluded, errors = process_workbook(file_bytes)
    print(f"    Total Raw Rows in Sheet:       {total_raw}")
    print(f"    Excluded Rows:                 {len(excluded)}")
    print(f"    Validation & Format Errors:    {len(errors)}")
    print(f"    Eligible Leads to Merge:       {len(eligible)}")

    if errors:
        print("\n[!] CRITICAL: Validation errors encountered in source data (Sync Blocked):")
        for err in errors[:10]:
            print(f"    - Row #{err['row_number']}: Field '{err['field']}' -> {err['error']}")
        print("=" * 70)
        return False

    # 2. Database Comparison Simulation (Read-Only)
    print("\n[2] Database Comparison Simulation (Read-Only against Supabase):")
    conn = None
    passwords = [cfg.get("SUPABASE_PASSWORD", "")]
    try:
        conn = connect_supabase(cfg, readonly=True)
        with conn.cursor() as cur:
            # Check for composite key duplicates in target table
            cur.execute("""
                SELECT lower(trim(name)), lower(trim(coalesce(industry, ''))), count(*)
                FROM public.marketing_analysis_clean
                GROUP BY 1, 2
                HAVING count(*) > 1;
            """)
            ambiguous_target = cur.fetchall()
            if ambiguous_target:
                print("    [!] WARNING: Target table contains duplicate composite keys:")
                for name_val, ind_val, cnt in ambiguous_target:
                    print(f"        - Duplicate key in Supabase has {cnt} rows.")
                print("=" * 70)
                return False

            # Fetch all existing database records
            cols_sql = ", ".join(ALL_COLUMNS)
            cur.execute(f"SELECT {cols_sql} FROM public.marketing_analysis_clean;")
            existing_db_rows = cur.fetchall()

            # Map DB records by composite key
            db_records_by_key = {}
            for row in existing_db_rows:
                row_dict = dict(zip(ALL_COLUMNS, row))
                key = (
                    (row_dict["name"] or "").lower().strip(),
                    (row_dict["industry"] or "").lower().strip(),
                )
                db_records_by_key[key] = row_dict

            # Compare Excel eligible rows against DB records
            excel_keys = set()
            new_records_to_insert = 0
            records_to_update = 0
            records_unchanged = 0

            for _, rec in eligible:
                rec_key = (
                    (rec["name"] or "").lower().strip(),
                    (rec["industry"] or "").lower().strip(),
                )
                excel_keys.add(rec_key)

                if rec_key not in db_records_by_key:
                    new_records_to_insert += 1
                else:
                    db_rec = db_records_by_key[rec_key]
                    # Check if any non-key field is distinct
                    has_diff = False
                    for col in NON_KEY_COLUMNS:
                        val_excel = rec.get(col)
                        val_db = db_rec.get(col)
                        v1 = str(val_excel).strip() if val_excel is not None else None
                        v2 = str(val_db).strip() if val_db is not None else None
                        if v1 != v2:
                            has_diff = True
                            break
                    if has_diff:
                        records_to_update += 1
                    else:
                        records_unchanged += 1

            # Preserved database-only records
            preserved_keys = set(db_records_by_key.keys()) - excel_keys
            preserved_count = len(preserved_keys)
            current_db_count = len(existing_db_rows)
            projected_db_count = current_db_count + new_records_to_insert

            print(f"    Current Records in Supabase:       {current_db_count}")
            print(f"    Unchanged Records:                 {records_unchanged}")
            print(f"    Records Requiring Update:          {records_to_update}")
            print(f"    New Records to Insert:             {new_records_to_insert}")
            print(f"    Database-Only Records Preserved:   {preserved_count}")
            print(f"    Projected Total Database Rows:     {projected_db_count}")

            # 3. Check Downstream View
            print("\n[3] Downstream View Verification (public.demo_timing):")
            cur.execute("SELECT count(*), round(avg(days_to_demo), 2) FROM public.demo_timing;")
            view_count, avg_days = cur.fetchone()
            print(f"    Current Demo Timing View Rows:     {view_count}")
            print(f"    Current Average Days to Demo:      {avg_days}")

            print("\n" + "=" * 70)
            print("DRY-RUN RESULT: PASSED (Simulation Successful - Ready for Controlled Live Sync)")
            print("=" * 70)
            return True

    except Exception as e:
        safe_err = sanitize_message(str(e), passwords)
        print(f"\n[-] DRY-RUN FAILED:\n    Error: {safe_err}")
        print("=" * 70)
        return False
    finally:
        if conn and not conn.closed:
            conn.close()


def execute_live_sync(file_bytes: bytes, source_name: str, cfg: dict, force: bool = False) -> bool:
    """
    Executes idempotent live merge into Supabase:
      - Checks SHA-256 against cloud state table; skips if unchanged (unless --force)
      - Opens single atomic transaction:
          * Pre-sync row count
          * CREATE TEMP TABLE staging_marketing_analysis (LIKE public.marketing_analysis_clean) ON COMMIT DROP;
          * Inserts eligible records into staging table
          * Updates matching records where non-key fields differ (IS DISTINCT FROM)
          * Inserts new records not in target table
          * Counts preserved database-only records
          * Verifies post-sync row count and demo_timing view
          * Upserts persistent cloud state
          * Commits transaction
      - On any error, rolls back transaction and logs sanitized error
    """
    file_size = len(file_bytes)
    current_hash = compute_bytes_sha256(file_bytes)

    print("=" * 70)
    print("PALOWISE CLOUD SYNC: LIVE ATOMIC MERGE")
    print("=" * 70)
    print(f"Source:                    {source_name} ({file_size:,} bytes)")
    print(f"SHA-256:                   {current_hash}")
    print(f"Target Database:           Supabase ({cfg['SUPABASE_HOST']}:{cfg['SUPABASE_PORT']})")
    print(f"SSL Mode:                  {cfg.get('SUPABASE_SSLMODE', 'require')} (Enforced)")
    print(f"Mode:                      LIVE ATOMIC MERGE (Single Transaction)")
    print("-" * 70)

    # 1. Workbook Processing & Validation
    print("[1/4] Validating Workbook & Business Rules...")
    total_raw, eligible, excluded, errors = process_workbook(file_bytes)
    print(f"    Total Raw Rows in Sheet:       {total_raw}")
    print(f"    Excluded Rows:                 {len(excluded)}")
    print(f"    Validation & Format Errors:    {len(errors)}")
    print(f"    Eligible Leads to Merge:       {len(eligible)}")

    if errors:
        print("\n[!] CRITICAL: Validation errors encountered in source data (Sync Aborted):")
        for err in errors[:10]:
            print(f"    - Row #{err['row_number']}: Field '{err['field']}' -> {err['error']}")
        return False

    passwords = [cfg.get("SUPABASE_PASSWORD", "")]
    conn = None

    try:
        conn = connect_supabase(cfg, readonly=False)
        conn.autocommit = False

        # 2. Check SHA-256 against Supabase cloud state
        print("\n[2/4] Evaluating Cloud State...")
        last_state = get_cloud_state(conn)

        if last_state and not force:
            if last_state.get("sha256") == current_hash:
                print(f"    RECURRING SYNC NOTICE: Workbook unchanged since last sync (SHA-256 match).")
                print(f"    Last Synced: {last_state.get('last_synced_at')} (Rows in DB: {last_state.get('rows_in_db')}).")
                print("    Skipping database write. Zero compute consumed.")
                print("=" * 70)
                return True

        # 3. Execute Atomic Sync via Staging Table inside transaction
        print("\n[3/4] Executing Atomic Merge into Supabase...")
        with conn.cursor() as cur:
            # Check target table composite key duplicates
            cur.execute("""
                SELECT lower(trim(name)), lower(trim(coalesce(industry, ''))), count(*)
                FROM public.marketing_analysis_clean
                GROUP BY 1, 2
                HAVING count(*) > 1;
            """)
            ambiguous_target = cur.fetchall()
            if ambiguous_target:
                print("CRITICAL INTEGRITY HALT: Target table contains duplicate composite keys. Aborting.")
                conn.rollback()
                return False

            # Pre-sync count
            cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
            pre_count = cur.fetchone()[0]

            # Create temporary staging table (auto-dropped on commit)
            cur.execute("""
                CREATE TEMP TABLE staging_marketing_analysis (
                    LIKE public.marketing_analysis_clean
                ) ON COMMIT DROP;
            """)

            # Populate staging table
            cols_sql = ", ".join(ALL_COLUMNS)
            staging_rows = [[rec[c] for c in ALL_COLUMNS] for _, rec in eligible]
            insert_staging = f"INSERT INTO staging_marketing_analysis ({cols_sql}) VALUES %s;"
            extras.execute_values(cur, insert_staging, staging_rows, page_size=500)

            # UPDATE existing records where any non-key field IS DISTINCT FROM
            set_clauses = [f"{c} = s.{c}" for c in NON_KEY_COLUMNS]
            update_set_sql = ", ".join(set_clauses)
            distinct_conditions = [f"m.{c} IS DISTINCT FROM s.{c}" for c in NON_KEY_COLUMNS]
            distinct_sql = " OR ".join(distinct_conditions)

            update_sql = f"""
                UPDATE public.marketing_analysis_clean m
                SET {update_set_sql}
                FROM staging_marketing_analysis s
                WHERE lower(trim(m.name)) = lower(trim(s.name))
                  AND lower(trim(coalesce(m.industry, ''))) = lower(trim(coalesce(s.industry, '')))
                  AND ({distinct_sql});
            """
            cur.execute(update_sql)
            updated_count = cur.rowcount

            # INSERT new records
            insert_new_sql = f"""
                INSERT INTO public.marketing_analysis_clean ({cols_sql})
                SELECT s.{', s.'.join(ALL_COLUMNS)}
                FROM staging_marketing_analysis s
                LEFT JOIN public.marketing_analysis_clean m
                  ON lower(trim(s.name)) = lower(trim(m.name))
                 AND lower(trim(coalesce(s.industry, ''))) = lower(trim(coalesce(m.industry, '')))
                WHERE m.name IS NULL;
            """
            cur.execute(insert_new_sql)
            inserted_count = cur.rowcount

            # Preserved count (in DB but not in Excel)
            cur.execute("""
                SELECT count(*)
                FROM public.marketing_analysis_clean m
                LEFT JOIN staging_marketing_analysis s
                  ON lower(trim(m.name)) = lower(trim(s.name))
                 AND lower(trim(coalesce(m.industry, ''))) = lower(trim(coalesce(m.industry, '')))
                WHERE s.name IS NULL;
            """)
            preserved_count = cur.fetchone()[0]

            # Post-sync count
            cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
            post_count = cur.fetchone()[0]

            # 4. Verify downstream view
            cur.execute("SELECT count(*), round(avg(days_to_demo), 2) FROM public.demo_timing;")
            view_count, avg_days = cur.fetchone()

            # Save updated cloud state
            save_cloud_state(conn, current_hash, file_size, post_count, inserted_count, updated_count, preserved_count)

            # Commit transaction atomically
            conn.commit()

        print(f"\n[4/4] Commit Succeeded!")
        print(f"    Pre-sync DB Rows:          {pre_count}")
        print(f"    New Leads Inserted:        {inserted_count}")
        print(f"    Existing Leads Updated:    {updated_count}")
        print(f"    Leads Preserved (DB-only): {preserved_count}")
        print(f"    Total Database Rows:       {post_count}")
        print(f"    Demo Timing View Rows:     {view_count} (Avg Days to Demo: {avg_days})")
        print("=" * 70)
        return True

    except Exception as e:
        if conn and not conn.closed:
            conn.rollback()
        safe_err = sanitize_message(str(e), passwords)
        print(f"\n[-] CLOUD SYNC FAILED (Transaction Rolled Back):\n    Error: {safe_err}")
        return False
    finally:
        if conn and not conn.closed:
            conn.close()


# ==============================================================================
# CLI Entrypoint
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(description="Palowise Cloud Sync Worker for Supabase")
    parser.add_argument("--dry-run", action="store_true", help="Validate and simulate sync without writing to Supabase (Default)")
    parser.add_argument("--live", action="store_true", help="Execute live idempotent synchronization into Supabase")
    parser.add_argument("--force", action="store_true", help="Force sync execution ignoring SHA-256 state")
    parser.add_argument("--local-file", type=Path, help="Use a local Excel file instead of downloading from OneDrive")
    parser.add_argument("--share-url", type=str, help="Override OneDrive sharing URL")

    args = parser.parse_args()
    cfg = load_cloud_config()

    try:
        file_bytes, source_name = resolve_excel_bytes(cfg, local_file=args.local_file, share_url=args.share_url)
    except Exception as e:
        safe_msg = sanitize_message(str(e))
        print("STATUS: FAILED")
        print(f"ERROR: {safe_msg}")
        sys.exit(1)

    # Determine execution mode: safety default is dry-run unless --live is explicitly set
    if args.live:
        success = execute_live_sync(file_bytes, source_name, cfg, force=args.force)
    else:
        success = execute_dry_run(file_bytes, source_name, cfg)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
