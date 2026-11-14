"""
Cloud Sync Worker: OneDrive Excel to Supabase Cloud PostgreSQL
==============================================================
Lumos Cloud Automation Worker

Workflow:
  1. Retrieves latest Excel workbook from OneDrive via Microsoft Graph API
     (or fallback to local file if configured for testing).
  2. Evaluates SHA-256 fingerprint against Supabase sync state table.
     Skips processing if workbook is unchanged (idempotent, compute-saving).
  3. Applies rigorous Lumos business validations:
     - Excludes dummy test rows ('Test').
     - Excludes blank rows / missing lead names.
     - Deduplicates composite keys: lower(trim(name)) + lower(trim(industry)).
     - Validates and parses ordinal date formats (e.g., '17th May').
     - Standardizes flag columns ('YES' / 'NO').
  4. Executes atomic upsert into Supabase 'public.marketing_analysis_clean' via
     temporary staging table with 'IS DISTINCT FROM' diffing.
  5. Verifies downstream analytical view 'public.demo_timing'.
  6. Updates persistent sync state in Supabase.

Modes:
  --dry-run       Validate workbook and simulate database merge without writing.
  --force         Bypass SHA-256 state check and force validation/sync.
  --local-file    Specify a local file path instead of downloading via Microsoft Graph API.
  --test-graph    Test Microsoft Graph API authentication and OneDrive file discovery.
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
    # Optional fallback to local .env.supabase if present
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
        # Microsoft Graph API Settings
        "MS_GRAPH_TENANT_ID": os.environ.get("MS_GRAPH_TENANT_ID", ""),
        "MS_GRAPH_CLIENT_ID": os.environ.get("MS_GRAPH_CLIENT_ID", ""),
        "MS_GRAPH_CLIENT_SECRET": os.environ.get("MS_GRAPH_CLIENT_SECRET", ""),
        "ONEDRIVE_USER_PRINCIPAL_NAME": os.environ.get("ONEDRIVE_USER_PRINCIPAL_NAME", ""),
        "ONEDRIVE_FILE_PATH": os.environ.get("ONEDRIVE_FILE_PATH", "lumos10.xlsx"),
    }


# ==============================================================================
# Microsoft Graph API: OneDrive Downloader
# ==============================================================================

def get_graph_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """
    Requests an application-level OAuth 2.0 token from Microsoft Entra ID.
    Uses Python standard library (urllib.request) to avoid extra dependencies.
    """
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

    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            return body["access_token"]
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8")
        raise RuntimeError(f"Microsoft Graph OAuth token request failed (HTTP {e.code}): {err_msg}")
    except Exception as e:
        raise RuntimeError(f"Failed to obtain Microsoft Graph token: {e}")


def download_onedrive_workbook(cfg: dict) -> tuple[bytes, str]:
    """
    Downloads Excel workbook binary content from OneDrive using Microsoft Graph API.
    Endpoint: /users/{user-principal-name}/drive/root:/{path}:/content
    """
    tenant_id = cfg.get("MS_GRAPH_TENANT_ID")
    client_id = cfg.get("MS_GRAPH_CLIENT_ID")
    client_secret = cfg.get("MS_GRAPH_CLIENT_SECRET")
    upn = cfg.get("ONEDRIVE_USER_PRINCIPAL_NAME")
    file_path = cfg.get("ONEDRIVE_FILE_PATH", "lumos10.xlsx").strip("/")

    if not all([tenant_id, client_id, client_secret, upn]):
        raise ValueError(
            "Missing Microsoft Graph credentials. Required: MS_GRAPH_TENANT_ID, "
            "MS_GRAPH_CLIENT_ID, MS_GRAPH_CLIENT_SECRET, ONEDRIVE_USER_PRINCIPAL_NAME."
        )

    print("Authenticating with Microsoft Graph API via OAuth 2.0 client credentials...")
    access_token = get_graph_access_token(tenant_id, client_id, client_secret)

    # Encode path for URL
    encoded_path = urllib.parse.quote(file_path)
    url = f"https://graph.microsoft.com/v1.0/users/{upn}/drive/root:/{encoded_path}:/content"

    print(f"Downloading workbook from OneDrive: '{file_path}' (Account: {upn})...")
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        method="GET",
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            file_bytes = resp.read()
            size_kb = len(file_bytes) / 1024
            print(f"SUCCESS: Downloaded '{file_path}' ({size_kb:.1f} KB, {len(file_bytes):,} bytes).")
            return file_bytes, file_path
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8")
        raise RuntimeError(f"Microsoft Graph file download failed (HTTP {e.code}): {err_msg}")
    except Exception as e:
        raise RuntimeError(f"Failed to download OneDrive file: {e}")


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
# Database Connection & Idempotent Sync Execution
# ==============================================================================

def connect_supabase(cfg: dict):
    return psycopg2.connect(
        host=cfg["SUPABASE_HOST"],
        port=int(cfg["SUPABASE_PORT"]),
        dbname=cfg["SUPABASE_DATABASE"],
        user=cfg["SUPABASE_USER"],
        password=cfg.get("SUPABASE_PASSWORD", ""),
        sslmode=cfg.get("SUPABASE_SSLMODE", "require"),
        connect_timeout=20,
    )


def execute_sync(file_bytes: bytes, source_name: str, cfg: dict, dry_run: bool = False, force: bool = False) -> bool:
    """
    Executes idempotent merge into Supabase.
    """
    file_size = len(file_bytes)
    current_hash = compute_bytes_sha256(file_bytes)

    print("=" * 70)
    print("LUMOS CLOUD AUTOMATION SYNC WORKER")
    print("=" * 70)
    print(f"Source:           {source_name} ({file_size:,} bytes)")
    print(f"SHA-256:          {current_hash}")
    print(f"Target Database:  Supabase ({cfg['SUPABASE_HOST']}:{cfg['SUPABASE_PORT']})")
    print(f"Mode:             {'DRY-RUN (No Writes)' if dry_run else 'LIVE ATOMIC MERGE'}")
    print("-" * 70)

    # 1. Process and validate workbook
    print("[1/4] Validating Workbook & Business Rules...")
    total_raw, eligible, excluded, errors = process_workbook(file_bytes)
    print(f"    Total Raw Rows:          {total_raw}")
    print(f"    Excluded Rows:           {len(excluded)}")
    print(f"    Eligible Leads to Merge: {len(eligible)}")
    print(f"    Validation Errors:       {len(errors)}")

    if errors:
        print("\n[!] CRITICAL: Validation errors encountered in source data (Sync Aborted):")
        for err in errors[:10]:
            print(f"    - Row #{err['row_number']} (Lead '{err['lead_name']}'): {err['field']} -> {err['error']}")
        return False

    passwords = [cfg.get("SUPABASE_PASSWORD", "")]
    conn = None

    try:
        conn = connect_supabase(cfg)

        # 2. Check SHA-256 against Supabase cloud state
        print("\n[2/4] Evaluating Cloud State...")
        last_state = get_cloud_state(conn)

        if last_state and not force and not dry_run:
            if last_state.get("sha256") == current_hash:
                print(f"    RECURRING SYNC NOTICE: Workbook unchanged since last sync (SHA-256 match).")
                print(f"    Last Synced: {last_state.get('last_synced_at')} (Rows in DB: {last_state.get('rows_in_db')}).")
                print("    Skipping database write. Zero compute consumed.")
                print("=" * 70)
                return True

        if dry_run:
            print("    [Dry-run]: Skipping database writes.")
            print("=" * 70)
            return True

        # 3. Execute Atomic Sync via Staging Table
        print("\n[3/4] Executing Atomic Merge into Supabase...")
        with conn:
            with conn.cursor() as cur:
                # Pre-sync count
                cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
                pre_count = cur.fetchone()[0]

                # Create temporary staging table (auto-dropped on commit)
                cols_ddl = ", ".join(f"{col} text" for col in ALL_COLUMNS)
                cur.execute(f"CREATE TEMP TABLE staging_marketing_analysis ({cols_ddl}) ON COMMIT DROP;")

                # Populate staging table
                staging_rows = [[rec[c] for c in ALL_COLUMNS] for _, rec in eligible]
                cols_sql = ", ".join(ALL_COLUMNS)
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
                     AND lower(trim(coalesce(m.industry, ''))) = lower(trim(coalesce(s.industry, '')))
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

        print(f"\n[4/4] Commit Succeeded!")
        print(f"    Pre-sync DB Rows:        {pre_count}")
        print(f"    New Leads Inserted:      {inserted_count}")
        print(f"    Existing Leads Updated:  {updated_count}")
        print(f"    Leads Preserved:         {preserved_count}")
        print(f"    Total Database Rows:     {post_count}")
        print(f"    Demo Timing View Rows:   {view_count} (Avg Days to Demo: {avg_days})")
        print("=" * 70)
        return True

    except Exception as e:
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
    parser = argparse.ArgumentParser(description="Lumos Cloud Python Sync Worker for Supabase")
    parser.add_argument("--dry-run", action="store_true", help="Validate and simulate sync without writing to Supabase")
    parser.add_argument("--force", action="store_true", help="Force sync execution ignoring SHA-256 state")
    parser.add_argument("--local-file", type=Path, help="Use a local Excel file instead of downloading from OneDrive")
    parser.add_argument("--test-graph", action="store_true", help="Test Microsoft Graph API connection and download")

    args = parser.parse_args()
    cfg = load_cloud_config()

    # Mode 1: Test Graph API download only
    if args.test_graph:
        print("Testing Microsoft Graph API integration...")
        file_bytes, file_name = download_onedrive_workbook(cfg)
        print(f"Graph API download test successful: {file_name} ({len(file_bytes):,} bytes).")
        sys.exit(0)

    # Mode 2: Resolve Excel source
    if args.local_file:
        if not args.local_file.is_file():
            print(f"ERROR: Local file not found at: {args.local_file}")
            sys.exit(1)
        with open(args.local_file, "rb") as f:
            file_bytes = f.read()
        source_name = f"Local File: {args.local_file.name}"
    else:
        # Check if Graph credentials exist, else fallback to standard local OneDrive path if running on developer workstation
        has_graph = all([
            cfg.get("MS_GRAPH_TENANT_ID"),
            cfg.get("MS_GRAPH_CLIENT_ID"),
            cfg.get("MS_GRAPH_CLIENT_SECRET"),
            cfg.get("ONEDRIVE_USER_PRINCIPAL_NAME"),
        ])
        if has_graph:
            file_bytes, source_name = download_onedrive_workbook(cfg)
        else:
            local_fallback = Path(r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\lumos10.xlsx")
            if local_fallback.is_file():
                print(f"Notice: Microsoft Graph credentials not set. Falling back to local mirror: {local_fallback.name}")
                with open(local_fallback, "rb") as f:
                    file_bytes = f.read()
                source_name = f"Local Mirror: {local_fallback.name}"
            else:
                print("ERROR: Microsoft Graph API credentials are not set and local file does not exist.")
                print("Required environment variables:")
                print("  MS_GRAPH_TENANT_ID, MS_GRAPH_CLIENT_ID, MS_GRAPH_CLIENT_SECRET, ONEDRIVE_USER_PRINCIPAL_NAME")
                sys.exit(1)

    success = execute_sync(file_bytes, source_name, cfg, dry_run=args.dry_run, force=args.force)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()

