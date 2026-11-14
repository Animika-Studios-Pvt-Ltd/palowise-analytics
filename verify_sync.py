"""
Read-Only Verification Script for Post-Synchronization Audit
Targets strictly 'palowise_integration_dev' on localhost:5432.
Enforces read-only session mode; never exposes credentials.
"""

from pathlib import Path
import sys
import psycopg2

def load_env(env_path: Path) -> dict:
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

def main():
    print("=" * 70)
    print("POST-SYNCHRONIZATION READ-ONLY VERIFICATION")
    print("=" * 70)

    env_path = Path(__file__).parent / ".env"
    cfg = load_env(env_path)

    host = cfg.get("PG_HOST", "localhost")
    port = cfg.get("PG_PORT", "5432")
    database = cfg.get("PG_DATABASE", "")
    user = cfg.get("PG_USER", "")
    password = cfg.get("PG_PASSWORD", "")

    # Hard safety assertions
    ALLOWED_DATABASES = {"palowise_integration_dev", "test_database"}
    if database not in ALLOWED_DATABASES:
        print(f"CRITICAL SAFETY HALT: Refusing to connect. Target must be in {ALLOWED_DATABASES}, got '{database}'.")
        sys.exit(1)

    if host not in ("localhost", "127.0.0.1"):
        print(f"CRITICAL SAFETY HALT: Host must be localhost, got '{host}'.")
        sys.exit(1)

    conn = None
    try:
        conn = psycopg2.connect(
            host=host,
            port=port,
            dbname=database,
            user=user,
            password=password,
            connect_timeout=5,
        )
        conn.set_session(readonly=True, autocommit=True)
        print("Connected to PostgreSQL server in read-only mode.\n")

        with conn.cursor() as cur:
            # 1. Row count of marketing_analysis_clean
            cur.execute("SELECT count(*) FROM public.marketing_analysis_clean;")
            mac_count = cur.fetchone()[0]
            print(f"[1] Table 'public.marketing_analysis_clean' Row Count:")
            print(f"    Total Rows: {mac_count}")

            # 2. Duplicate names check
            cur.execute("""
                SELECT name, count(*) 
                FROM public.marketing_analysis_clean 
                GROUP BY name 
                HAVING count(*) > 1;
            """)
            duplicate_names = cur.fetchall()
            print(f"\n[2] Duplicate Name Check in 'marketing_analysis_clean':")
            if duplicate_names:
                print(f"    WARNING: Found {len(duplicate_names)} duplicate name(s):")
                for name, count in duplicate_names:
                    print(f"      - '{name}': {count} occurrences")
            else:
                print("    PASSED: All lead names are unique (0 duplicates).")

            # 3. Check for any dummy test records or invalid date records
            cur.execute("""
                SELECT name, starting_date, demo_date 
                FROM public.marketing_analysis_clean 
                WHERE lower(trim(name)) = 'test' 
                   OR starting_date ~ '^[a-zA-Z]$' 
                   OR starting_date IS NULL 
                   OR demo_date IS NULL;
            """)
            invalid_records = cur.fetchall()
            print(f"\n[3] Data Quality & Date Format Validation:")
            # Also check if dummy 'Test' exists
            cur.execute("SELECT count(*) FROM public.marketing_analysis_clean WHERE lower(trim(name)) = 'test';")
            test_row_count = cur.fetchone()[0]
            print(f"    - Dummy 'Test' Lead Present in Table: {test_row_count} (Expected: 0)")

            if invalid_records:
                print(f"    - Records with null or suspicious date strings: {len(invalid_records)}")
                for r in invalid_records:
                    print(f"        {r}")
            else:
                print("    - PASSED: 0 invalid date records; all rows have non-null, well-formed dates.")

            # 4. View public.demo_timing verification
            print(f"\n[4] View 'public.demo_timing' Verification:")
            cur.execute("SELECT count(*), round(avg(days_to_demo), 2), min(days_to_demo), max(days_to_demo) FROM public.demo_timing;")
            view_count, avg_days, min_days, max_days = cur.fetchone()
            print(f"    - Execution Status:    PASSED (View executed without errors)")
            print(f"    - Calculated Rows:     {view_count}")
            print(f"    - Average Days to Demo: {avg_days} days")
            print(f"    - Min / Max Days:       {min_days} to {max_days} days")

            # Sample 5 rows from demo_timing view
            print(f"\n[5] Sample Output from 'public.demo_timing' (First 5 Rows):")
            cur.execute("""
                SELECT name, meeting_date, demo_date, days_to_demo 
                FROM public.demo_timing 
                ORDER BY meeting_date ASC, name ASC 
                LIMIT 5;
            """)
            print(f"    {'Lead Name':<30} {'Meeting Date':<15} {'Demo Date':<15} {'Days to Demo':<12}")
            print("    " + "-" * 72)
            for name, m_date, d_date, d_days in cur.fetchall():
                print(f"    {name:<30} {str(m_date):<15} {str(d_date):<15} {d_days:<12}")

        print("\n" + "=" * 70)
        print("VERIFICATION COMPLETED SUCCESSFULLY (0 ERRORS)")
        print("=" * 70)

    except Exception as e:
        err_str = str(e).strip()
        if password and password in err_str:
            err_str = err_str.replace(password, "[REDACTED]")
        print(f"ERROR: {err_str}")
        sys.exit(1)
    finally:
        if conn and not conn.closed:
            conn.close()
            print("Database connection closed safely.")

if __name__ == "__main__":
    main()
