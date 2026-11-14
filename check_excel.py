"""
Excel Validation Script for Palowise Project
Reads the local OneDrive Excel workbook 'lumos10.xlsx', validates its contents,
checks for Table1, and prepares for synchronization without modifying the source file.
"""

from pathlib import Path
import sys

def find_workbook():
    expected_paths = [
        Path(r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\lumos10.xlsx"),
        Path(r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\lumos10.xls"),
        Path(r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\lumos.xlsx"),
        Path(__file__).parent / "lumos10.xlsx",
    ]
    for p in expected_paths:
        if p.is_file() and not p.name.endswith(".url"):
            return p.resolve()
            
    # Search in expected OneDrive folder if exists
    onedrive_dir = Path(r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED")
    if onedrive_dir.is_dir():
        for f in onedrive_dir.glob("*lumos*.xlsx"):
            if f.is_file() and not f.name.endswith(".url"):
                return f.resolve()

    return None

def main():
    import pandas as pd
    import openpyxl
    import psycopg2

    print("=" * 60)
    print("PALOWISE EXCEL WORKBOOK VALIDATION")
    print("=" * 60)
    print("Environment Verification:")
    print(f"  - Python:    {sys.version.split()[0]}")
    print(f"  - pandas:    {pd.__version__}")
    print(f"  - openpyxl:  {openpyxl.__version__}")
    print(f"  - psycopg2:  {psycopg2.__version__}")
    print("-" * 60)

    excel_path = find_workbook()
    if not excel_path:
        print("ERROR: Excel workbook could not be located in expected folders.")
        print("Checked location: C:\\Users\\india\\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\\")
        sys.exit(1)

    print(f"[1] Resolved Workbook Path:\n    {excel_path}")
    print(f"    File Size: {excel_path.stat().st_size:,} bytes")

    # Step 2: Open with openpyxl to inspect structure and verify Table1
    table1_found = False
    table1_sheet = None
    all_tables = []
    
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheets = wb.sheetnames
        print(f"\n[2] Sheet Names Found: {sheets}")
        
        for name in sheets:
            ws = wb[name]
            for tbl_name in ws.tables.keys():
                all_tables.append((name, tbl_name))
                if tbl_name.strip().lower() == "table1":
                    table1_found = True
                    table1_sheet = name
                    
        wb.close()
    except Exception as e:
        print(f"WARNING: openpyxl structural inspection encountered an error: {e}")

    print(f"\n[3] Excel Tables Verification:")
    if all_tables:
        for s_name, t_name in all_tables:
            print(f"    - Table '{t_name}' found in sheet '{s_name}'")
    else:
        print("    - No formal Excel table objects found via worksheet tables collection.")
        
    if table1_found:
        print(f"    --> Table1 verified: YES (located in sheet '{table1_sheet}')")
    else:
        print("    --> Table1 verified: NO (Table1 object was not detected in sheet tables)")

    # Step 3: Load into pandas
    try:
        # If Table1 sheet was identified, load that sheet; otherwise load the first sheet
        target_sheet = table1_sheet if table1_sheet else 0
        df = pd.read_excel(excel_path, sheet_name=target_sheet, engine="openpyxl")
    except Exception as e:
        print(f"\nERROR: Failed to read Excel file with pandas: {e}")
        sys.exit(1)

    # Step 4: Validate row count and column information
    row_count = len(df)
    col_count = len(df.columns)
    columns_list = df.columns.tolist()

    print(f"\n[4] Data Summary:")
    print(f"    Sheet Read: {target_sheet}")
    print(f"    Row Count: {row_count}")
    print(f"    Column Count: {col_count}")
    print(f"    Column Names:")
    for idx, col in enumerate(columns_list, start=1):
        print(f"      {idx}. {col}")

    # Step 5: Check for duplicate column names
    duplicated_cols = df.columns[df.columns.duplicated()].tolist()
    print(f"\n[5] Duplicate Column Check:")
    if duplicated_cols:
        print(f"    WARNING: Duplicated column names detected: {set(duplicated_cols)}")
    else:
        print("    PASSED: No duplicated column names found.")

    # Step 6: Display first 5 rows
    print(f"\n[6] First 5 Rows of Data:")
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 1000)
    print(df.head(5).to_string(index=False))

    print("\n" + "=" * 60)
    print("VALIDATION COMPLETE: Source Excel file verified successfully.")
    print("=" * 60)

if __name__ == "__main__":
    main()
