"""
Safe Read-Only PostgreSQL Connection Test for Palowise Project
Connects strictly to development database 'palowise_integration_dev' on localhost:5432.
Enforces read-only session mode and never exposes credentials.
"""

from pathlib import Path
import sys

def load_env_file(filepath: Path) -> dict:
    if not filepath.is_file():
        raise FileNotFoundError(f"Environment file not found at: {filepath}")
    
    env_vars = {}
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "=" in stripped:
                key, val = stripped.split("=", 1)
                env_vars[key.strip()] = val.strip().strip("'\"")
    return env_vars

def main():
    print("=" * 60)
    print("READ-ONLY POSTGRESQL CONNECTION TEST")
    print("=" * 60)

    env_path = Path(__file__).parent / ".env"
    try:
        cfg = load_env_file(env_path)
    except Exception as e:
        print(f"FAILED: Unable to load .env file: {e}")
        sys.exit(1)

    # 1. Validation & Safety Checks
    host = cfg.get("PG_HOST", "localhost")
    port = cfg.get("PG_PORT", "5432")
    database = cfg.get("PG_DATABASE", "")
    user = cfg.get("PG_USER", "")
    password = cfg.get("PG_PASSWORD", "")

    print(f"Target Host:     {host}:{port}")
    print(f"Target Database: {database}")
    print(f"Target User:     {user}")
    print(f"Password Status: {'[CONFIGURED]' if password else '[MISSING]'}")
    print("-" * 60)

    # Hard safety check: allowed databases
    ALLOWED_DATABASES = {"palowise_integration_dev", "test_database"}
    if database not in ALLOWED_DATABASES:
        print(f"CRITICAL SAFETY HALT: Target database must be in {ALLOWED_DATABASES}.")
        print(f"Encountered value: '{database}'. Execution aborted to prevent unauthorized connection.")
        sys.exit(1)

    if host not in ("localhost", "127.0.0.1"):
        print(f"CRITICAL SAFETY HALT: Target host must be localhost. Encountered '{host}'. Execution aborted.")
        sys.exit(1)

    # 2. Connect via psycopg2
    import psycopg2

    conn = None
    try:
        try:
            conn = psycopg2.connect(
                host=host,
                port=port,
                dbname=database,
                user=user,
                password=password,
                connect_timeout=5,
            )
        except psycopg2.OperationalError as e:
            if host == "localhost":
                # Attempt IPv4 127.0.0.1 in case pg_hba has distinct rules
                conn = psycopg2.connect(
                    host="127.0.0.1",
                    port=port,
                    dbname=database,
                    user=user,
                    password=password,
                    connect_timeout=5,
                )
            else:
                raise
        # Enforce read-only transaction mode
        conn.set_session(readonly=True, autocommit=True)
        print("SUCCESS: Connected to PostgreSQL server in read-only mode.\n")

        with conn.cursor() as cur:
            # Check 1: current database
            cur.execute("SELECT current_database();")
            current_db = cur.fetchone()[0]
            print(f"[1] Verified Active Database: {current_db}")

            # Check 2: current user
            cur.execute("SELECT current_user;")
            current_usr = cur.fetchone()[0]
            print(f"[2] Verified Current User:    {current_usr}")

            # Check 3: list tables in public schema
            cur.execute("""
                SELECT table_name 
                FROM information_schema.tables 
                WHERE table_schema = 'public' 
                ORDER BY table_name;
            """)
            tables = [row[0] for row in cur.fetchall()]
            print(f"\n[3] Tables in 'public' schema ({len(tables)} found):")
            if tables:
                for t in tables:
                    print(f"    - {t}")
            else:
                print("    - (None - schema is currently empty)")

            # Check 4: Row count of public.marketing_analysis_clean if exists
            target_table = "marketing_analysis_clean"
            print(f"\n[4] Target Table Inspection ('{target_table}'):")
            if target_table in tables:
                cur.execute(f"SELECT count(*) FROM public.{target_table};")
                row_count = cur.fetchone()[0]
                print(f"    - Table '{target_table}' exists with {row_count:,} rows.")
            else:
                print(f"    - Table '{target_table}' does not exist in the public schema.")

        print("\n" + "=" * 60)
        print("READ-ONLY TEST COMPLETED SUCCESSFULLY")
        print("=" * 60)

    except psycopg2.OperationalError as e:
        # Sanitize error message to prevent accidental credential leakage
        err_msg = str(e).strip()
        if password and password in err_msg:
            err_msg = err_msg.replace(password, "[REDACTED]")
        print(f"CONNECTION ERROR (OperationalError):\n{err_msg}")
        sys.exit(1)
    except Exception as e:
        err_msg = str(e).strip()
        if password and password in err_msg:
            err_msg = err_msg.replace(password, "[REDACTED]")
        print(f"QUERY ERROR:\n{err_msg}")
        sys.exit(1)
    finally:
        if conn and not conn.closed:
            conn.close()
            print("Database connection closed safely.")

if __name__ == "__main__":
    main()
