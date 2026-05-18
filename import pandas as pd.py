import pandas as pd
from pathlib import Path

def calculate_rewards_by_margin(margin_pct, slabs, reward_type):
    """Finds the correct reward level based on the computed profit margin"""
    if pd.isna(margin_pct):
        return 0
    for slab in slabs:
        if slab['min_margin'] <= margin_pct <= slab['max_margin']:
            return slab[reward_type]
    return 0

def process_excel_dynamic_master(source_folder_path, slabs, target_column):
    source_folder = Path(source_folder_path)
    if not source_folder.exists():
        print(f"❌ Error: The folder path '{source_folder_path}' does not exist.")
        return
        
    # Gather all .xlsx files from the source directory
    all_files = list(source_folder.glob("*.xlsx"))
    
    # Ignore temporary files and previously generated outputs
    excel_files = []
    for f in all_files:
        name = f.name
        if name.startswith("~$") or "Unified_Calculated_Rewards" in name or "Combined_Calculated_Rewards" in name or "Final_Rewards" in name:
            continue
        excel_files.append(f)
        
    if not excel_files:
        print(f"❌ Error: No clean raw Excel source files (.xlsx) found inside {source_folder_path}")
        return
        
    excel_file = excel_files[0]
    print(f"📂 Scanning Source Folder: {source_folder}")
    print(f"📄 Found True Source File: {excel_file.name}")

    try:
        xl = pd.ExcelFile(excel_file)
    except Exception as e:
        print(f"❌ Engine Error opening {excel_file.name}: {e}")
        return
        
    available_sheets = xl.sheet_names
    
    # Find the worksheet containing the most data rows
    target_sheet = None
    max_rows = -1
    df = None

    print("🔍 Analyzing workbook tabs to locate your master data...")
    for sheet in available_sheets:
        try:
            test_df = pd.read_excel(excel_file, sheet_name=sheet)
            if test_df.shape[0] > max_rows:
                max_rows = test_df.shape[0]
                target_sheet = sheet
                df = test_df
        except Exception:
            continue

    if df is None:
        print("❌ Error: Could not read data rows from any tab in this spreadsheet.")
        return

    print(f"📊 Auto-targeted Master Sheet: '{target_sheet}' (Contains {df.shape[0]} rows)")

    # Sanitize column headers and strip raw single quotes
    cleaned_headers = []
    for col in df.columns:
        col_str = str(col).strip().upper()
        col_str = col_str.strip("'").strip('"')  # Fixes headers like PRODUCT NAME' or B2B'
        if col_str.endswith('.0') and col_str[:-2].isdigit():
            col_str = col_str[:-2]
        cleaned_headers.append(col_str)
        
    df.columns = cleaned_headers

    # Auto-adjust column spelling typos
    if 'RP RPICING' in df.columns:
        print("🔧 Auto-correcting 'RP RPICING' header typo to 'RP PRICING'...")
        df = df.rename(columns={'RP RPICING': 'RP PRICING'})

    target_col_clean = str(target_column).strip().upper().strip("'").strip('"')

    # Verify key structural columns exist
    for col in ['B2B', 'RP PRICING', target_col_clean]:
        if col not in df.columns:
            print(f"❌ Error: Missing required column '{col}' from this sheet.")
            print(f"Available headers: {list(df.columns)}")
            return

    print("\n⚙️ Calculating margin metrics across ALL master rows...")
    
    # Cast variables safely to numeric types
    df['B2B'] = pd.to_numeric(df['B2B'], errors='coerce')
    df['RP PRICING'] = pd.to_numeric(df['RP PRICING'], errors='coerce')

    # Execute math operations across every single row
    df['Profit Value'] = (df['RP PRICING'] - df['B2B']).round(2)
    df['Profit Margin %'] = ((df['Profit Value'] / df['RP PRICING']) * 100).round(2)
    df['Profit Margin %'] = df['Profit Margin %'].fillna(0)

    print("🛡️ Mapping margin-safe rewards percentages...")
    df['Earn %'] = df['Profit Margin %'].apply(lambda x: calculate_rewards_by_margin(x, slabs, 'earn'))
    df['Redeem %'] = df['Profit Margin %'].apply(lambda x: calculate_rewards_by_margin(x, slabs, 'redeem'))

    df['Earn Value'] = (df['RP PRICING'] * (df['Earn %'] / 100)).round(2)
    df['Redeem Value'] = (df['RP PRICING'] * (df['Redeem %'] / 100)).round(2)

    # Save rows into a unified file inside the source folder
    output_file = source_folder / f"Rewards_{excel_file.name}"
    print(f"🗂️ Compiling all {len(df)} entries into a single spreadsheet...")

    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="All_Products_Rewards", index=False)
        print(f"\n✅ Success! Consolidated calculations sheet created at:\n{output_file}")
    except Exception as e:
        print(f"❌ Error writing output file: {e}")

if __name__ == "__main__":
    MARGIN_SLABS = [
        {"min_margin": -float('inf'), "max_margin": 10.0, "earn": 2,  "redeem": 2},   
        {"min_margin": 10.01,         "max_margin": 20.0, "earn": 5,  "redeem": 7},  
        {"min_margin": 20.01,         "max_margin": float('inf'), "earn": 10, "redeem": 12} 
    ]
    
    # POINTING DIRECTLY TO YOUR SOURCE DATA LOCATION
    RAW_DATA_FOLDER = r"C:\Users\User\Desktop\SHREE"
    COLUMN_TO_TARGET = "BRAND"
    
    process_excel_dynamic_master(RAW_DATA_FOLDER, MARGIN_SLABS, COLUMN_TO_TARGET)