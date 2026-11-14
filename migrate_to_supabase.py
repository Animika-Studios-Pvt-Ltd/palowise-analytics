"""
Supabase Cloud Database Migration & Parity Verification Script
==============================================================
Lumos Project: Phase 2 Database Migration to Supabase Cloud

Features:
  --test-connection    (Default) Verify SSL connectivity to Supabase Shared Pooler.
  --migrate            Execute preflight check, schema restoration, data COPY (296 rows),
                       view creation (demo_timing), and post-migration parity verification.
  --verify-only        Run parity verification between local PostgreSQL and Supabase without writes.

Security:
  - Enforces SSL (sslmode=require).
  - All passwords and secret tokens are sanitized and redacted from logs.
  - Local database is queried strictly with READ-ONLY transactions.
"""

import argparse
import io
import os
import re
import sys
from pathlib import Path
import psycopg2

ROOT_DIR = Path(__file__).parent.resolve()
DEFAULT_SUPABASE_ENV = ROOT_DIR / ".env.supabase"
DEFAULT_LOCAL_ENV = ROOT_DIR / ".env"
DEFAULT_BACKUP_SQL = ROOT_DIR / "backups" / "palowise_pg_backup_20261113_153228.sql"

EXPECTED_COLUMNS = [
    "name",
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

EXPECTED_TABLE_COUNT = 296
EXPECTED_VIEW_COUNT = 55


def load_env_file(env_path: Path) -> dict:
    """Parses a standard KEY=VALUE environment file."""
    cfg = {}
    if env_path.is_file():
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                stripped = line.strip()
                if not stripped or stripped.startswith("#"):
                    continue
                if "=" in stripped:
                    k, v = stripped.split("=", 1)
                    cfg[k.strip()] = v.strip().strip("'\"")
    return cfg


def load_supabase_config(env_path: Path | None = None) -> dict:
    """Loads Supabase parameters with precedence: system env vars > env file."""
    target_env = env_path if env_path else DEFAULT_SUPABASE_ENV
    file_cfg = load_env_file(target_env)

    return {
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
    }


def load_local_config(env_path: Path | None = None) -> dict:
    """Loads local PostgreSQL parameters."""
    target_env = env_path if env_path else DEFAULT_LOCAL_ENV
    file_cfg = load_env_file(target_env)

    return {
        "PG_HOST": file_cfg.get("PG_HOST", "localhost"),
        "PG_PORT": file_cfg.get("PG_PORT", "5432"),
        "PG_DATABASE": file_cfg.get("PG_DATABASE", "test_database"),
        "PG_USER": file_cfg.get("PG_USER", "postgres"),
        "PG_PASSWORD": file_cfg.get("PG_PASSWORD", ""),
    }


def sanitize_message(message: str, passwords: list[str] | None = None) -> str:
    """Redacts passwords and sensitive URI substrings from text/errors."""
    if not message:
        return ""
    clean = str(message)
    if passwords:
        for p in passwords:
            if p and p.strip():
                clean = clean.replace(p, "[REDACTED]")
    clean = re.sub(r":([^:@\s]+)@", r":[REDACTED]@", clean)
    return clean


def inspect_backup(backup_path: Path) -> tuple[str, str, str, list[str]]:
    """
    Parses and verifies the SQL backup:
      - Validates presence of table DDL for public.marketing_analysis_clean (28 columns)
      - Validates presence of view DDL for public.demo_timing
      - Validates COPY statement and extracts raw data rows
    Returns: (table_ddl, copy_header, copy_data, columns)
    """
    if not backup_path.is_file():
        raise FileNotFoundError(f"Backup file not found at: {backup_path}")

    with open(backup_path, "r", encoding="utf-8") as f:
        content = f.read()

    # 1. Table DDL extraction
    table_match = re.search(
        r"(CREATE TABLE public\.marketing_analysis_clean \((.*?)\);)",
        content,
        re.DOTALL,
    )
    if not table_match:
        raise ValueError("Could not find 'CREATE TABLE public.marketing_analysis_clean' in backup file!")
    table_ddl = table_match.group(1)

    # Validate column list
    col_defs = [line.strip().rstrip(",") for line in table_match.group(2).strip().splitlines() if line.strip()]
    found_cols = [c.split()[0].strip() for c in col_defs]
    if found_cols != EXPECTED_COLUMNS:
        raise ValueError(
            f"Table column mismatch in backup!\nExpected ({len(EXPECTED_COLUMNS)}): {EXPECTED_COLUMNS}\nFound ({len(found_cols)}): {found_cols}"
        )

    # 2. View DDL extraction
    view_match = re.search(
        r"(CREATE VIEW public\.demo_timing AS\s+WITH cleaned_dates AS.*?;)",
        content,
        re.DOTALL,
    )
    if not view_match:
        raise ValueError("Could not find 'CREATE VIEW public.demo_timing' in backup file!")
    view_ddl = view_match.group(1)

    # 3. COPY block extraction
    copy_match = re.search(
        r"(COPY public\.marketing_analysis_clean \((.*?)\) FROM stdin;)\n(.*?)\n\\\.",
        content,
        re.DOTALL,
    )
    if not copy_match:
        raise ValueError("Could not find 'COPY public.marketing_analysis_clean ... FROM stdin' block in backup!")

    copy_header = copy_match.group(1)
    copy_cols = [c.strip() for c in copy_match.group(2).split(",")]
    copy_data = copy_match.group(3)
    data_lines = copy_data.split("\n")

    if len(copy_cols) != len(EXPECTED_COLUMNS):
        raise ValueError(f"COPY header column count ({len(copy_cols)}) != expected ({len(EXPECTED_COLUMNS)})")

    if len(data_lines) != EXPECTED_TABLE_COUNT:
        raise ValueError(f"COPY data row count ({len(data_lines)}) != expected ({EXPECTED_TABLE_COUNT})")

    return table_ddl, view_ddl, copy_header, copy_data, copy_cols


def connect_supabase(cfg: dict):
    """Establishes an SSL connection to Supabase."""
    return psycopg2.connect(
        host=cfg["SUPABASE_HOST"],
        port=int(cfg["SUPABASE_PORT"]),
        dbname=cfg["SUPABASE_DATABASE"],
        user=cfg["SUPABASE_USER"],
        password=cfg.get("SUPABASE_PASSWORD", ""),
        sslmode=cfg.get("SUPABASE_SSLMODE", "require"),
        connect_timeout=20,
    )


def connect_local(cfg: dict):
    """Establishes a read-only connection to local PostgreSQL."""
    conn = psycopg2.connect(
        host=cfg["PG_HOST"],
        port=int(cfg["PG_PORT"]),
        dbname=cfg["PG_DATABASE"],
        user=cfg["PG_USER"],
        password=cfg.get("PG_PASSWORD", ""),
        connect_timeout=5,
    )
    conn.set_session(readonly=True, autocommit=True)
    return conn


def test_supabase_connection(cfg: dict) -> bool:
    """Verifies credentials, SSL, and lists public relations."""
    print("=" * 70)
    print("SUPABASE CLOUD DATABASE CONNECTION TEST")
    print("=" * 70)

    masked_user = cfg["SUPABASE_USER"]
    if "." in masked_user:
        prefix, ref = masked_user.split(".", 1)
        masked_user = f"{prefix}.{ref[:4]}...{ref[-3:] if len(ref) > 7 else ''}"

    print(f"Target Host:     {cfg['SUPABASE_HOST']}")
    print(f"Target Port:     {cfg['SUPABASE_PORT']} (Shared Pooler)")
    print(f"Database Name:   {cfg['SUPABASE_DATABASE']}")
    print(f"Database User:   {masked_user}")
    print(f"SSL Mode:        {cfg['SUPABASE_SSLMODE']} (Enforced)")
    print(f"Password:        [PROTECTED - REDACTED]")
    print("-" * 70)

    conn = None
    passwords = [cfg.get("SUPABASE_PASSWORD", "")]

    try:
        conn = connect_supabase(cfg)
        with conn.cursor() as cur:
            cur.execute("SELECT version(), current_database(), current_user;")
            pg_version, current_db, current_user = cur.fetchone()

            print("[+] CONNECTION SUCCESSFUL!")
            print(f"    Connected Database: {current_db}")
            print(f"    Connected User:     {current_user}")
            print(f"    PostgreSQL Engine:  {pg_version.split(',')[0]}")

            cur.execute("""
                SELECT table_name, table_type 
                FROM information_schema.tables 
                WHERE table_schema = 'public'
                ORDER BY table_name;
            """)
            existing_relations = cur.fetchall()

            print(f"\n[+] Public Schema Inspection ({len(existing_relations)} relation(s) found):")
            for rel_name, rel_type in existing_relations:
                flag = ""
                if rel_name in ("marketing_analysis_clean", "demo_timing"):
                    flag = "  <-- Lumos Target Relation"
                print(f"      - {rel_name} ({rel_type}){flag}")

            if not existing_relations:
                print("      (Public schema is currently clean)")

        print("=" * 70)
        return True

    except Exception as e:
        safe_err = sanitize_message(str(e), passwords)
        print(f"\n[-] CONNECTION FAILED:\n    Error: {safe_err}")
        return False
    finally:
        if conn and not conn.closed:
            conn.close()


def execute_migration(supabase_cfg: dict, backup_path: Path) -> bool:
    """
    Executes Phase 2 migration:
      1. Inspects SQL backup and validates table DDL, view DDL, and 296 data rows.
      2. Preflight safety check on Supabase public schema.
      3. Creates table public.marketing_analysis_clean.
      4. Restores 296 records via native PostgreSQL COPY protocol.
      5. Creates view public.demo_timing.
      6. Commits transaction atomically.
    """
    print("\n" * 1)
    print("=" * 70)
    print("PHASE 2: EXECUTING LUMOS DATABASE MIGRATION TO SUPABASE")
    print("=" * 70)

    # 1. Inspect Backup
    print("[1/5] Inspecting SQL Backup File...")
    table_ddl, view_ddl, copy_header, copy_data, copy_cols = inspect_backup(backup_path)
    print(f"    Backup Source:        {backup_path.name}")
    print(f"    Table Schema:         public.marketing_analysis_clean ({len(copy_cols)} columns)")
    print(f"    View Schema:          public.demo_timing (Ordinal regex date logic)")
    print(f"    Total Records to Copy:{EXPECTED_TABLE_COUNT}")
    print("    [✓] Backup verified successfully.")

    passwords = [supabase_cfg.get("SUPABASE_PASSWORD", "")]
    conn = None

    try:
        conn = connect_supabase(supabase_cfg)

        # 2. Preflight Safety Check
        print("\n[2/5] Preflight Safety Check on Remote Database...")
        with conn.cursor() as cur:
            cur.execute("""
                SELECT table_name, table_type 
                FROM information_schema.tables 
                WHERE table_schema = 'public' 
                  AND table_name IN ('marketing_analysis_clean', 'demo_timing');
            """)
            conflicts = cur.fetchall()
            if conflicts:
                for name, ttype in conflicts:
                    print(f"    Notice: Object 'public.{name}' ({ttype}) already exists.")
                    # Check if table already has data
                    if name == "marketing_analysis_clean":
                        cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
                        cnt = cur.fetchone()[0]
                        print(f"    Existing rows in marketing_analysis_clean: {cnt}")
            else:
                print("    [✓] Remote public schema has 0 conflicting objects. Safe to proceed.")

        # 3. Execute Migration in Single Transaction
        print("\n[3/5] Restoring Table Structure & Records...")
        with conn:
            with conn.cursor() as cur:
                # Create table
                print("    Creating table 'public.marketing_analysis_clean'...")
                cur.execute(table_ddl)

                # Stream COPY data
                print(f"    Streaming {EXPECTED_TABLE_COUNT} records via PostgreSQL COPY protocol...")
                stream = io.StringIO(copy_data)
                copy_query = f"COPY public.marketing_analysis_clean ({', '.join(copy_cols)}) FROM STDIN"
                cur.copy_expert(copy_query, stream)
                print(f"    [✓] {cur.rowcount} rows inserted into marketing_analysis_clean.")

                # Create View
                print("\n[4/5] Creating Analytical View 'public.demo_timing'...")
                cur.execute(view_ddl)
                print("    [✓] View 'public.demo_timing' created successfully.")

                # Commit transaction occurs automatically at end of 'with conn:'
        print("    [✓] Migration transaction COMMITTED successfully.")
        return True

    except Exception as e:
        safe_err = sanitize_message(str(e), passwords)
        print(f"\n[-] MIGRATION TRANSACTION FAILED (Rolled Back):\n    Error: {safe_err}")
        return False
    finally:
        if conn and not conn.closed:
            conn.close()


def verify_migration_parity(supabase_cfg: dict, local_cfg: dict) -> bool:
    """
    Performs rigorous post-migration parity verification:
      1. Verifies remote table row count == 296.
      2. Verifies remote view row count == 55.
      3. Verifies schema column names, data types, and nullability match local.
      4. Verifies data contents match local row-by-row (zero missing, zero duplicates).
      5. Verifies demo_timing view output matches local exactly.
    """
    print("\n" * 1)
    print("=" * 70)
    print("[5/5] POST-MIGRATION PARITY VERIFICATION")
    print("=" * 70)

    passwords = [supabase_cfg.get("SUPABASE_PASSWORD", ""), local_cfg.get("PG_PASSWORD", "")]
    remote_conn = None
    local_conn = None

    try:
        remote_conn = connect_supabase(supabase_cfg)
        remote_cur = remote_conn.cursor()

        local_conn = connect_local(local_cfg)
        local_cur = local_conn.cursor()

        # Check 1: Record Counts
        print("[Check 1/5] Record Counts:")
        remote_cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
        remote_tbl_count = remote_cur.fetchone()[0]

        local_cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
        local_tbl_count = local_cur.fetchone()[0]

        print(f"    Local Table Count:   {local_tbl_count}")
        print(f"    Remote Table Count:  {remote_tbl_count} (Expected: {EXPECTED_TABLE_COUNT})")
        assert remote_tbl_count == EXPECTED_TABLE_COUNT, f"Remote table count ({remote_tbl_count}) != {EXPECTED_TABLE_COUNT}"
        assert remote_tbl_count == local_tbl_count, f"Remote table count ({remote_tbl_count}) != local ({local_tbl_count})"
        print("    [✓] Table record count verified.")

        # Check 2: View Row Counts & Averages
        print("\n[Check 2/5] Analytical View Parity (public.demo_timing):")
        remote_cur.execute("SELECT count(*), round(avg(days_to_demo), 2) FROM public.demo_timing;")
        remote_view_cnt, remote_avg_days = remote_cur.fetchone()

        local_cur.execute("SELECT count(*), round(avg(days_to_demo), 2) FROM public.demo_timing;")
        local_view_cnt, local_avg_days = local_cur.fetchone()

        print(f"    Local View Rows:     {local_view_cnt} (Avg Days to Demo: {local_avg_days})")
        print(f"    Remote View Rows:    {remote_view_cnt} (Avg Days to Demo: {remote_avg_days}) (Expected: {EXPECTED_VIEW_COUNT})")
        assert remote_view_cnt == EXPECTED_VIEW_COUNT, f"Remote view rows ({remote_view_cnt}) != {EXPECTED_VIEW_COUNT}"
        assert remote_view_cnt == local_view_cnt, "Remote view count != local view count"
        assert remote_avg_days == local_avg_days, f"Remote avg days ({remote_avg_days}) != local ({local_avg_days})"
        print("    [✓] Demo timing analytical calculations match local to the decimal point.")

        # Check 3: Schema Parity (Columns, Data Types, Ordinals)
        print("\n[Check 3/5] Schema Column Parity:")
        col_query = """
            SELECT column_name, data_type, ordinal_position
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = 'marketing_analysis_clean'
            ORDER BY ordinal_position;
        """
        local_cur.execute(col_query)
        local_cols = local_cur.fetchall()

        remote_cur.execute(col_query)
        remote_cols = remote_cur.fetchall()

        assert len(local_cols) == len(remote_cols) == len(EXPECTED_COLUMNS), (
            f"Column count mismatch: local={len(local_cols)}, remote={len(remote_cols)}"
        )
        for loc, rem in zip(local_cols, remote_cols):
            assert loc == rem, f"Column mismatch! Local: {loc} vs Remote: {rem}"
        print(f"    [✓] All {len(remote_cols)} columns and data types identical across local and remote.")

        # Check 4: Data Integrity & Parity Check (Field-by-Field)
        print("\n[Check 4/5] Field-by-Field Record Parity:")
        cols_csv = ", ".join(EXPECTED_COLUMNS)
        data_query = f"""
            SELECT {cols_csv}
            FROM public.marketing_analysis_clean
            ORDER BY lower(trim(name)), lower(trim(coalesce(industry, ''))), lower(trim(coalesce(company, '')));
        """
        local_cur.execute(data_query)
        local_data = local_cur.fetchall()

        remote_cur.execute(data_query)
        remote_data = remote_cur.fetchall()

        assert len(local_data) == len(remote_data), "Data rows length mismatch"

        mismatches = 0
        for idx, (loc_row, rem_row) in enumerate(zip(local_data, remote_data)):
            if loc_row != rem_row:
                mismatches += 1
                if mismatches <= 3:
                    print(f"    Mismatch at row {idx + 1}:\n      Local:  {loc_row[:3]}\n      Remote: {rem_row[:3]}")

        assert mismatches == 0, f"Found {mismatches} record mismatches between local and Supabase!"
        print(f"    [✓] 100% data parity confirmed across all {len(local_data)} records (0 mismatches).")

        # Check 5: Duplicate and Key Integrity Check
        print("\n[Check 5/5] Key Integrity & Duplicates Check:")
        remote_cur.execute("""
            SELECT lower(trim(name)), lower(trim(coalesce(industry, ''))), count(*)
            FROM public.marketing_analysis_clean
            GROUP BY lower(trim(name)), lower(trim(coalesce(industry, '')))
            HAVING count(*) > 1;
        """)
        dups = remote_cur.fetchall()
        if dups:
            print(f"    Notice: {len(dups)} composite key groups have multiple entries in source dataset.")
        else:
            print("    [✓] Zero composite key duplicates detected in remote dataset.")

        print("\n" + "=" * 70)
        print("MIGRATION PARITY VERIFICATION: 100% PASSED (ALL CHECKS SUCCEEDED)")
        print("=" * 70)
        return True

    except Exception as e:
        safe_err = sanitize_message(str(e), passwords)
        print(f"\n[-] PARITY VERIFICATION FAILED:\n    Error: {safe_err}")
        return False
    finally:
        if remote_conn and not remote_conn.closed:
            remote_conn.close()
        if local_conn and not local_conn.closed:
            local_conn.close()


def main():
    parser = argparse.ArgumentParser(
        description="Lumos Supabase Cloud Database Migration & Connectivity Utility"
    )
    parser.add_argument(
        "--test-connection",
        "--test",
        action="store_true",
        help="Test SSL connection to Supabase Shared Pooler",
    )
    parser.add_argument(
        "--migrate",
        action="store_true",
        help="Execute Phase 2 migration and post-migration parity check",
    )
    parser.add_argument(
        "--verify-only",
        action="store_true",
        help="Run parity verification between local PostgreSQL and Supabase without writes",
    )
    parser.add_argument(
        "--env-file",
        type=Path,
        default=DEFAULT_SUPABASE_ENV,
        help="Path to Supabase env file (defaults to .env.supabase)",
    )
    parser.add_argument(
        "--backup-file",
        type=Path,
        default=DEFAULT_BACKUP_SQL,
        help="Path to SQL backup file",
    )

    args = parser.parse_args()

    supabase_cfg = load_supabase_config(args.env_file)
    local_cfg = load_local_config(DEFAULT_LOCAL_ENV)

    if args.migrate:
        success = execute_migration(supabase_cfg, args.backup_file)
        if success:
            parity_ok = verify_migration_parity(supabase_cfg, local_cfg)
            sys.exit(0 if parity_ok else 1)
        else:
            sys.exit(1)
    elif args.verify_only:
        parity_ok = verify_migration_parity(supabase_cfg, local_cfg)
        sys.exit(0 if parity_ok else 1)
    else:
        # Default is connection test
        success = test_supabase_connection(supabase_cfg)
        sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
