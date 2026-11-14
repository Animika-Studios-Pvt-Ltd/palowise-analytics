"""
Step 3: Schema Compatibility Validation Script
Reads Excel workbook and PostgreSQL table schema in read-only mode.
Compares columns, data types, value formats, and duplicate lead names.
Strictly read-only; no data or schema modifications.
"""

from pathlib import Path
import sys
import pandas as pd
import openpyxl
import psycopg2

def load_env(env_path: Path) -> dict:
    env_vars = {}
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                k, v = line.split("=", 1)
                env_vars[k.strip()] = v.strip().strip("'\"")
    return env_vars

def main():
    print("=" * 70)
    print("STEP 3: SCHEMA COMPATIBILITY VALIDATION")
    print("=" * 70)

    # 1. Load Excel Workbook
    excel_path = Path(r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\lumos10.xlsx")
    if not excel_path.is_file():
        print(f"ERROR: Excel file not found at: {excel_path}")
        sys.exit(1)

    print(f"\n[1] Excel Source Inspection:")
    print(f"    Path: {excel_path}")
    print(f"    Size: {excel_path.stat().st_size:,} bytes")

    # Read Excel sheet
    df = pd.read_excel(excel_path, sheet_name="in", engine="openpyxl")
    raw_rows = len(df)
    
    # Check for completely blank rows
    blank_rows = df.isna().all(axis=1).sum()
    
    # Clean df without completely empty rows for analysis
    active_df = df.dropna(how="all").copy()
    active_rows = len(active_df)

    print(f"    Total Rows in Sheet: {raw_rows}")
    print(f"    Completely Blank Rows: {blank_rows}")
    print(f"    Active Data Rows: {active_rows}")
    print(f"    Excel Column Count: {len(df.columns)}")
    print(f"    Excel Columns:")
    for idx, col in enumerate(df.columns, start=1):
        non_null_cnt = df[col].notna().sum()
        dtype_str = str(df[col].dtype)
        print(f"      {idx:2d}. '{col}' (dtype: {dtype_str}, non-null: {non_null_cnt}/{raw_rows})")

    # Duplicate lead names check
    name_col = None
    for c in df.columns:
        if c.strip().lower() == "name":
            name_col = c
            break

    print(f"\n[2] Lead Name Analysis (Column: '{name_col}'):")
    if name_col:
        names = active_df[name_col].dropna().astype(str).str.strip()
        dupes = names[names.duplicated(keep=False)].value_counts()
        if len(dupes) > 0:
            print(f"    WARNING: Duplicate lead names found ({len(dupes)} duplicated names):")
            for d_name, cnt in dupes.items():
                print(f"      - '{d_name}' appears {cnt} times")
        else:
            print("    PASSED: All lead names are unique (0 duplicates).")
    else:
        print("    WARNING: 'NAME' column not detected.")

    # Inspect specific formatting: Dates and Flags
    print(f"\n[3] Excel Data Formatting Sample:")
    sample_cols = ["STARTING DATE", "Demo Date", "Pilot Conducted", "Proposals Submitted", "Follow Up"]
    for sc in sample_cols:
        if sc in df.columns:
            vals = active_df[sc].dropna().unique()
            print(f"    - '{sc}' unique values (first 10): {list(vals[:10])}")

    # 2. Connect to PostgreSQL (Read-Only)
    env_path = Path(__file__).parent / ".env"
    cfg = load_env(env_path)

    host = cfg.get("PG_HOST", "localhost")
    port = cfg.get("PG_PORT", "5432")
    database = cfg.get("PG_DATABASE", "")
    user = cfg.get("PG_USER", "")
    password = cfg.get("PG_PASSWORD", "")

    # Safety assertion
    ALLOWED_DATABASES = {"palowise_integration_dev", "test_database"}
    if database not in ALLOWED_DATABASES:
        print(f"CRITICAL ERROR: Refusing to connect to '{database}'. Must be in {ALLOWED_DATABASES}.")
        sys.exit(1)

    print(f"\n[4] PostgreSQL Schema Inspection:")
    print(f"    Target Server: {host}:{port}")
    print(f"    Database:      {database}")

    conn = psycopg2.connect(
        host=host,
        port=port,
        dbname=database,
        user=user,
        password=password,
        connect_timeout=5,
    )
    conn.set_session(readonly=True, autocommit=True)

    pg_columns = []
    constraints = []
    view_columns = []

    try:
        with conn.cursor() as cur:
            # Query columns of marketing_analysis_clean
            cur.execute("""
                SELECT 
                    ordinal_position,
                    column_name, 
                    data_type, 
                    character_maximum_length,
                    is_nullable, 
                    column_default
                FROM information_schema.columns 
                WHERE table_schema = 'public' 
                  AND table_name = 'marketing_analysis_clean'
                ORDER BY ordinal_position;
            """)
            pg_columns = cur.fetchall()

            # Query constraints on marketing_analysis_clean
            cur.execute("""
                SELECT conname, contype, pg_get_constraintdef(c.oid)
                FROM pg_constraint c
                JOIN pg_namespace n ON n.oid = c.connamespace
                JOIN pg_class cl ON cl.oid = c.conrelid
                WHERE n.nspname = 'public' AND cl.relname = 'marketing_analysis_clean';
            """)
            constraints = cur.fetchall()

            # Query demo_timing view
            cur.execute("""
                SELECT table_name, table_type
                FROM information_schema.tables
                WHERE table_schema = 'public' AND table_name = 'demo_timing';
            """)
            demo_timing_meta = cur.fetchall()

            cur.execute("""
                SELECT column_name, data_type
                FROM information_schema.columns
                WHERE table_schema = 'public' AND table_name = 'demo_timing'
                ORDER BY ordinal_position;
            """)
            view_columns = cur.fetchall()

            cur.execute("SELECT pg_get_viewdef('public.demo_timing'::regclass, true);")
            view_def = cur.fetchone()[0]

    finally:
        conn.close()

    print(f"    Table 'public.marketing_analysis_clean' ({len(pg_columns)} columns):")
    pg_col_dict = {}
    for col in pg_columns:
        pos, cname, dtype, char_len, nullable, default_val = col
        type_str = f"{dtype}({char_len})" if char_len else dtype
        pg_col_dict[cname.lower()] = {
            "name": cname,
            "type": type_str,
            "nullable": nullable,
            "default": default_val,
        }
        print(f"      {pos:2d}. {cname:<25} {type_str:<20} Nullable: {nullable:<5} Default: {default_val}")

    if constraints:
        print(f"\n    Constraints on 'marketing_analysis_clean':")
        for cname, ctype, cdef in constraints:
            type_label = {"p": "PRIMARY KEY", "u": "UNIQUE", "c": "CHECK", "f": "FOREIGN KEY"}.get(ctype, ctype)
            print(f"      - [{type_label}] {cname}: {cdef}")
    else:
        print(f"\n    Constraints on 'marketing_analysis_clean': None (no primary keys or unique constraints).")

    if demo_timing_meta:
        print(f"\n    Object 'public.demo_timing' ({demo_timing_meta[0][1]}):")
        for cname, dtype in view_columns:
            print(f"      - {cname:<25} ({dtype})")
        print(f"    View SQL Definition:\n{view_def.strip()}")

    # 3. Schema Comparison & Mapping Analysis
    print(f"\n[5] Column Mapping & Compatibility Comparison:")
    excel_cols = list(df.columns)
    pg_col_names = [col[1] for col in pg_columns]
    pg_col_names_lower = {k: v["name"] for k, v in pg_col_dict.items()}

    # Exact matching and normalized matching (lowercase + spaces to underscores)
    matched = []
    unmatched_excel = []
    
    for ec in excel_cols:
        norm_ec = ec.strip().lower().replace(" ", "_").replace("-", "_")
        exact_match = ec in pg_col_names
        norm_match = norm_ec in pg_col_names_lower
        
        if exact_match:
            matched.append((ec, ec, "Exact name match", pg_col_dict[ec.lower()]["type"]))
        elif norm_match:
            pg_target = pg_col_names_lower[norm_ec]
            matched.append((ec, pg_target, "Normalized match (spaces -> underscores / lowercase)", pg_col_dict[norm_ec]["type"]))
        else:
            unmatched_excel.append(ec)

    print(f"    Matched Columns ({len(matched)} / {len(excel_cols)}):")
    for orig, target, reason, target_type in matched:
        print(f"      - Excel: '{orig}' --> PG: '{target}' ({target_type}) [{reason}]")

    if unmatched_excel:
        print(f"\n    Unmatched Excel Columns ({len(unmatched_excel)}):")
        for uc in unmatched_excel:
            print(f"      - '{uc}'")
    else:
        print(f"\n    Unmatched Excel Columns: None (All Excel columns map to PostgreSQL!)")

    # Check for extra columns in PG not in Excel
    matched_pg_names = [m[1] for m in matched]
    extra_pg_cols = [c for c in pg_col_names if c not in matched_pg_names]
    print(f"\n    PostgreSQL Columns Not in Excel ({len(extra_pg_cols)}):")
    if extra_pg_cols:
        for ep in extra_pg_cols:
            print(f"      - '{ep}' ({pg_col_dict[ep.lower()]['type']}, Nullable: {pg_col_dict[ep.lower()]['nullable']}, Default: {pg_col_dict[ep.lower()]['default']})")
    else:
        print("      - None")

    print("\n" + "=" * 70)
    print("VALIDATION SUMMARY FINISHED")
    print("=" * 70)

if __name__ == "__main__":
    main()
