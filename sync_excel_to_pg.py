"""
Palowise Excel-to-PostgreSQL Synchronization Script
Loads verified OneDrive Excel workbook 'lumos10.xlsx' into PostgreSQL table 'public.marketing_analysis_clean'.
Targets strictly the development database 'palowise_integration_dev'.

Modes:
  --dry-run    (Default) Validate workbook, check readiness, and preview changes without writing.
  --live       Execute single live idempotent synchronization.
  --recurring  Automation mode: verifies OneDrive readiness, evaluates SHA-256 state,
               skips redundant runs, and performs safe idempotent synchronization if changes exist.
  --force      Bypass SHA-256 state check and force validation/sync.
"""

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
import time
import pandas as pd
import openpyxl

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

VALID_MONTHS = {"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"}

STATE_FILE = Path(__file__).parent / ".sync_state.json"

def load_env(env_path: Path) -> dict:
    if not env_path.is_file():
        raise FileNotFoundError(f".env file not found at: {env_path}")
    
    cfg = {}
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "=" in stripped:
                k, v = stripped.split("=", 1)
                cfg[k.strip()] = v.strip().strip("'\"")
    return cfg

def compute_sha256(filepath: Path) -> str:
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()

def load_state() -> dict:
    if STATE_FILE.is_file():
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_state(state_data: dict):
    tmp_path = STATE_FILE.with_suffix(".tmp")
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(state_data, f, indent=2)
    tmp_path.replace(STATE_FILE)

def check_onedrive_readiness(excel_path: Path) -> tuple[bool, str]:
    """
    Checks whether the OneDrive source file is ready to read and not in the middle of syncing.
    """
    if not excel_path.is_file():
        return False, f"Source workbook not found at: {excel_path}"

    # 1. Check for temporary lock files created by Excel or OneDrive (e.g. ~$lumos10.xlsx)
    parent_dir = excel_path.parent
    lock_file = parent_dir / f"~${excel_path.name}"
    if lock_file.is_file():
        return False, f"Temporary lock file detected ('{lock_file.name}'). Excel or OneDrive is currently modifying the file."

    # 2. Check file size
    stat_info = excel_path.stat()
    if stat_info.st_size == 0:
        return False, "Workbook file size is 0 bytes (incomplete cloud sync in-flight)."

    # 3. Settling window debounce (10 seconds)
    mtime_delta = time.time() - stat_info.st_mtime
    if mtime_delta < 10:
        return False, f"File was modified {mtime_delta:.1f} seconds ago. Deferring to allow OneDrive sync to settle."

    # 4. Probe shared read access
    try:
        with open(excel_path, "rb") as f:
            f.read(4096)
    except PermissionError:
        return False, "File is locked by an active process (OneDrive chunk sync or exclusive Excel handle)."
    except Exception as e:
        return False, f"File read probe error: {e}"

    return True, "Workbook is accessible and ready."

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

def clean_cell_value(val) -> str | None:
    if pd.isna(val) or val is None:
        return None
    val_str = str(val).strip()
    return val_str if val_str else None

def process_workbook(excel_path: Path):
    """
    Reads workbook, applies cleaning, filters dummy rows, validates formats,
    and detects any source ambiguity.
    Returns: (total_raw_rows, eligible_records, excluded_records, validation_errors)
    """
    df = pd.read_excel(excel_path, sheet_name="in", engine="openpyxl")
    total_raw_rows = len(df)

    eligible_records = []
    excluded_records = []
    validation_errors = []
    records_by_key = {}

    for idx, row in df.iterrows():
        row_number = idx + 2
        raw_name = clean_cell_value(row.get("NAME"))

        # Rule 1: Exclude dummy test row (NAME is "Test", case-insensitively)
        if raw_name and raw_name.lower() == "test":
            excluded_records.append({
                "row_number": row_number,
                "name": raw_name,
                "reason": "Excluded dummy test record (NAME == 'Test')",
                "raw_data": {k: clean_cell_value(v) for k, v in row.items()}
            })
            continue

        # Rule 2: Check for blank row or missing lead name
        if not raw_name:
            if row.isna().all():
                excluded_records.append({
                    "row_number": row_number,
                    "name": None,
                    "reason": "Excluded completely blank row",
                    "raw_data": {}
                })
            else:
                excluded_records.append({
                    "row_number": row_number,
                    "name": None,
                    "reason": "Excluded row with missing lead NAME (trailing metadata or orphan row)",
                    "raw_data": {k: clean_cell_value(v) for k, v in row.dropna().items()}
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
                        "error": msg
                    })
                    row_has_error = True

            record[pg_col] = cell_val

        if row_has_error:
            continue

        # Rule 3: Handle duplicate composite keys in Source Excel (name + industry)
        norm_key = (
            record["name"].lower().strip(),
            (record["industry"] or "").lower().strip()
        )
        if norm_key in records_by_key:
            prior_row, prior_rec = records_by_key[norm_key]
            excluded_records.append({
                "row_number": prior_row,
                "name": prior_rec["name"],
                "reason": f"Superseded by later updated Excel entry on row #{row_number}",
                "raw_data": prior_rec
            })
        records_by_key[norm_key] = (row_number, record)

    eligible_records = list(records_by_key.values())
    return total_raw_rows, eligible_records, excluded_records, validation_errors

def run_dry_run(excel_path: Path, cfg: dict):
    print("=" * 70)
    print("PALOWISE EXCEL-TO-POSTGRESQL SYNC: DRY-RUN REPORT")
    print("=" * 70)
    print(f"Excel Source Path: {excel_path}")
    target_db = cfg.get("PG_DATABASE", "palowise_integration_dev")
    print(f"Target Database:   {target_db} (public.marketing_analysis_clean)")
    print(f"Mode:              DRY-RUN (Safe validation only; 0 database writes)")
    print("-" * 70)

    # 1. Readiness check
    ready, ready_msg = check_onedrive_readiness(excel_path)
    print(f"[1] OneDrive Source Readiness:\n    Status: {'READY' if ready else 'NOT READY'} -> {ready_msg}")

    # 2. Process workbook
    total_raw, eligible, excluded, errors = process_workbook(excel_path)

    print(f"\n[2] Record Processing Summary:")
    print(f"    Total Rows in Excel Sheet:    {total_raw}")
    print(f"    Excluded Rows:                {len(excluded)}")
    print(f"    Validation & Ambiguity Errors:{len(errors)}")
    print(f"    Eligible Rows for Merge:      {len(eligible)}")

    if excluded:
        print(f"\n[3] Excluded Rows Details:")
        for exc in excluded:
            print(f"    - Excel Row #{exc['row_number']}: Lead '{exc['name']}' -> Reason: {exc['reason']}")

    if errors:
        print(f"\n[4] Validation & Ambiguity Errors (Blocking):")
        for err in errors:
            print(f"    - Excel Row #{err['row_number']} (Lead '{err.get('lead_name')}'): Field '{err['field']}' -> {err['error']}")
    else:
        print(f"\n[4] Validation Status: PASSED (0 format, validation, or source ambiguity errors).")

    # 3. Simulate Staging Diff against Current PostgreSQL State (Read-Only)
    import psycopg2
    target_host = cfg.get("PG_HOST", "localhost")
    target_port = cfg.get("PG_PORT", "5432")
    target_db = cfg.get("PG_DATABASE", "palowise_integration_dev")
    user = cfg.get("PG_USER", "")
    password = cfg.get("PG_PASSWORD", "")

    print(f"\n[5] Database Comparison Simulation (Read-Only against {target_db}):")
    conn = None
    try:
        conn = psycopg2.connect(
            host=target_host,
            port=target_port,
            dbname=target_db,
            user=user,
            password=password,
            connect_timeout=5,
        )
        conn.set_session(readonly=True, autocommit=True)
        with conn.cursor() as cur:
            # Query existing rows
            cur.execute("SELECT name, industry FROM public.marketing_analysis_clean;")
            existing_rows = cur.fetchall()
            existing_keys = {
                (r[0].lower().strip(), (r[1] or "").lower().strip())
                for r in existing_rows
            }
            excel_keys = {
                (rec["name"].lower().strip(), (rec["industry"] or "").lower().strip())
                for _, rec in eligible
            }

            to_insert_keys = excel_keys - existing_keys
            matching_keys = excel_keys & existing_keys
            preserved_keys = existing_keys - excel_keys

            print(f"    Current Records in PostgreSQL:     {len(existing_rows)}")
            print(f"    Matching Leads to Update / Keep:   {len(matching_keys)}")
            print(f"    New Leads to Insert:               {len(to_insert_keys)}")
            print(f"    Existing Leads Preserved (Not in Excel): {len(preserved_keys)}")

            if to_insert_keys:
                print(f"    Sample New Leads to Insert: {list(to_insert_keys)[:3]}")
            if preserved_keys:
                print(f"    Sample Preserved Leads:     {list(preserved_keys)[:3]}")

    except Exception as e:
        print(f"    Notice: Could not query existing database rows for simulation: {e}")
    finally:
        if conn and not conn.closed:
            conn.close()

    print("\n" + "=" * 70)
    if len(errors) == 0 and len(eligible) > 0 and ready:
        print("DRY-RUN RESULT: PASSED - Cleaned data is ready for idempotent synchronization.")
    else:
        print("DRY-RUN RESULT: BLOCKED - Fix errors or wait for file readiness before proceeding.")
    print("=" * 70)

def execute_idempotent_sync(excel_path: Path, cfg: dict, is_recurring: bool = False, force: bool = False):
    """
    Executes idempotent merge into public.marketing_analysis_clean using a temporary staging table.
    Performs all operations inside a single atomic transaction.
    """
    target_host = cfg.get("PG_HOST", "localhost")
    target_port = cfg.get("PG_PORT", "5432")
    target_db = cfg.get("PG_DATABASE", "")
    user = cfg.get("PG_USER", "")
    password = cfg.get("PG_PASSWORD", "")

    # Safety assertions
    ALLOWED_DATABASES = {"palowise_integration_dev", "test_database"}
    if target_db not in ALLOWED_DATABASES:
        print(f"CRITICAL SAFETY HALT: Target database '{target_db}' is not authorized. Must be one of {ALLOWED_DATABASES}.")
        sys.exit(1)
    if target_host not in ("localhost", "127.0.0.1"):
        print(f"CRITICAL SAFETY HALT: Target host must be localhost. Got '{target_host}'.")
        sys.exit(1)

    # 1. OneDrive Readiness Check
    ready, msg = check_onedrive_readiness(excel_path)
    if not ready:
        print(f"ONEDRIVE SYNC NOTICE: {msg}")
        print("Execution deferred safely. No database operations attempted.")
        return False

    # 2. SHA-256 State Check for Recurring Mode
    current_hash = compute_sha256(excel_path)
    stat_info = excel_path.stat()
    last_state = load_state()

    if is_recurring and not force:
        if last_state.get("sha256") == current_hash:
            print(f"RECURRING SYNC NOTICE: Source workbook unchanged since last sync (SHA-256 match).")
            print(f"Skipping database connection. [Last sync: {last_state.get('last_synced_at', 'unknown')}]")
            return True

    # 3. Clean and validate workbook data
    total_raw, eligible, excluded, errors = process_workbook(excel_path)
    if errors:
        print(f"ABORTED: Found {len(errors)} validation/ambiguity errors. Aborting sync to protect data integrity:")
        for err in errors:
            print(f"  - Row #{err['row_number']}: {err['field']} -> {err['error']}")
        return False

    if not eligible:
        print("ABORTED: No eligible records found in workbook.")
        return False

    # 4. Connect to PostgreSQL and begin transaction
    import psycopg2
    from psycopg2.extras import execute_values

    conn = None
    try:
        conn = psycopg2.connect(
            host=target_host,
            port=target_port,
            dbname=target_db,
            user=user,
            password=password,
            connect_timeout=5,
        )
        conn.autocommit = False

        with conn.cursor() as cur:
            # Check 4a: Target table ambiguity check
            cur.execute("""
                SELECT lower(trim(name)), lower(trim(coalesce(industry, ''))), count(*)
                FROM public.marketing_analysis_clean
                GROUP BY 1, 2
                HAVING count(*) > 1;
            """)
            ambiguous_target = cur.fetchall()
            if ambiguous_target:
                print("CRITICAL INTEGRITY HALT: Target table contains duplicate composite keys:")
                for name, ind, cnt in ambiguous_target:
                    print(f"  - Lead '{name}' in industry '{ind}' has {cnt} rows in PostgreSQL.")
                print("Aborting to prevent non-deterministic updates. Manual deduplication required.")
                conn.rollback()
                return False

            # Check 4b: Record pre-sync state
            cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
            pre_count = cur.fetchone()[0]

            # 5. Create temporary staging table with identical schema (auto-dropped on commit)
            cur.execute("""
                CREATE TEMP TABLE staging_marketing_analysis (
                    LIKE public.marketing_analysis_clean
                ) ON COMMIT DROP;
            """)

            # 6. Insert eligible Excel records into staging table
            cols = list(COLUMN_MAPPING.values())
            cols_sql = ", ".join(cols)
            insert_staging_sql = f"INSERT INTO staging_marketing_analysis ({cols_sql}) VALUES %s"
            staging_data = [tuple(rec[col] for col in cols) for _, rec in eligible]
            execute_values(cur, insert_staging_sql, staging_data)

            # 7. Update existing matching records when attributes have changed
            update_clauses = [f"{col} = s.{col}" for col in NON_KEY_COLUMNS]
            update_set_sql = ", ".join(update_clauses)
            distinct_clauses = [f"m.{col} IS DISTINCT FROM s.{col}" for col in NON_KEY_COLUMNS]
            distinct_condition_sql = " OR ".join(distinct_clauses)

            update_sql = f"""
                UPDATE public.marketing_analysis_clean m
                SET {update_set_sql}
                FROM staging_marketing_analysis s
                WHERE lower(trim(m.name)) = lower(trim(s.name))
                  AND lower(trim(coalesce(m.industry, ''))) = lower(trim(coalesce(s.industry, '')))
                  AND ({distinct_condition_sql});
            """
            cur.execute(update_sql)
            updated_count = cur.rowcount

            # 8. Insert new leads not currently in PostgreSQL
            insert_new_sql = f"""
                INSERT INTO public.marketing_analysis_clean ({cols_sql})
                SELECT s.{', s.'.join(cols)}
                FROM staging_marketing_analysis s
                LEFT JOIN public.marketing_analysis_clean m
                  ON lower(trim(s.name)) = lower(trim(m.name))
                 AND lower(trim(coalesce(s.industry, ''))) = lower(trim(coalesce(m.industry, '')))
                WHERE m.name IS NULL;
            """
            cur.execute(insert_new_sql)
            inserted_count = cur.rowcount

            # 9. Count preserved rows (in PostgreSQL but absent from Excel)
            cur.execute("""
                SELECT count(*)
                FROM public.marketing_analysis_clean m
                LEFT JOIN staging_marketing_analysis s
                  ON lower(trim(m.name)) = lower(trim(s.name))
                 AND lower(trim(coalesce(m.industry, ''))) = lower(trim(coalesce(s.industry, '')))
                WHERE s.name IS NULL;
            """)
            preserved_count = cur.fetchone()[0]

            # 10. Check post-sync count
            cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
            post_count = cur.fetchone()[0]

            # 11. Downstream verification: verify demo_timing view executes cleanly
            cur.execute("SELECT count(*), round(avg(days_to_demo), 1) FROM public.demo_timing;")
            view_count, avg_days = cur.fetchone()

            # 12. Commit transaction
            conn.commit()

        # 13. Update SHA-256 state file ONLY after successful commit
        save_state({
            "file_size": stat_info.st_size,
            "last_modified": stat_info.st_mtime,
            "sha256": current_hash,
            "last_synced_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "rows_in_db": post_count,
            "inserted": inserted_count,
            "updated": updated_count,
            "preserved": preserved_count,
        })

        print("=" * 70)
        print("PALOWISE RECURRING SYNC: SUCCESSFUL COMMIT")
        print("=" * 70)
        print(f"Pre-sync Table Count:        {pre_count}")
        print(f"New Leads Inserted:          {inserted_count}")
        print(f"Existing Leads Updated:      {updated_count}")
        print(f"Leads Preserved (DB-only):   {preserved_count}")
        print(f"Total Current Rows:          {post_count}")
        print(f"Downstream View Rows:        {view_count} (Avg Days to Demo: {avg_days})")
        print(f"State File Updated:          {STATE_FILE.name} (SHA-256: {current_hash[:12]}...)")
        print("=" * 70)
        return True

    except Exception as e:
        if conn:
            conn.rollback()
            print("CRITICAL: Transaction rolled back due to error. Database state preserved.")
        err_msg = str(e).strip()
        if password and password in err_msg:
            err_msg = err_msg.replace(password, "[REDACTED]")
        print(f"DATABASE SYNC ERROR:\n{err_msg}")
        return False
    finally:
        if conn and not conn.closed:
            conn.close()

def main():
    parser = argparse.ArgumentParser(description="Palowise Excel-to-PostgreSQL Recurring Sync")
    parser.add_argument("--dry-run", action="store_true", help="Perform dry-run validation without writing to database")
    parser.add_argument("--live", action="store_true", help="Execute single live idempotent synchronization")
    parser.add_argument("--recurring", action="store_true", help="Execute recurring sync (state check + safe merge)")
    parser.add_argument("--force", action="store_true", help="Force execution ignoring state file hash")
    args = parser.parse_args()

    env_path = Path(__file__).parent / ".env"
    cfg = load_env(env_path)

    excel_path = Path(r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\lumos10.xlsx")

    # If --live or --recurring is passed, check which action to run
    if args.live:
        execute_idempotent_sync(excel_path, cfg, is_recurring=False, force=True)
    elif args.recurring:
        execute_idempotent_sync(excel_path, cfg, is_recurring=True, force=args.force)
    else:
        # Default is safe dry-run
        run_dry_run(excel_path, cfg)

if __name__ == "__main__":
    main()
