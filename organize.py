import pandas as pd
from pathlib import Path

def process_excel_with_fixed_rewards(folder_path, target_column):
    project_folder = Path(folder_path)
    
    if not project_folder.exists():
        print(f"❌ Error: The project folder path '{folder_path}' does not exist.")
        return
        
    all_files = list(project_folder.glob("*.xlsx"))
    
    # Grab any real Excel file, ignoring only temporary Windows files (~$)
    excel_files = [f for f in all_files if not f.name.startswith("~$")]
        
    if not excel_files:
        print(f"❌ Error: No Excel files (.xlsx) found inside {folder_path}")
        return
        
    # Pick the first available data file
    excel_file = excel_files[0]
    print(f"📂 Project Directory: {project_folder}")
    print(f"📄 Target File Found: {excel_file.name}")

    try:
        xl = pd.ExcelFile(excel_file)
    except Exception as e:
        print(f"❌ Engine Error opening {excel_file.name}: {e}")
        return
        
    # Auto-target the sheet with data
    available_sheets = xl.sheet_names
    target_sheet = available_sheets[0]
    df = pd.read_excel(excel_file, sheet_name=target_sheet)

    print(f"📊 Sheet Targeted: '{target_sheet}' (Contains {df.shape[0]} rows)")

    # Sanitize column headers completely
    df.columns = [str(col).strip().upper().strip("'").strip('"') for col in df.columns]

    # Auto-adjust column spelling typos if present
    if 'RP RPICING' in df.columns:
        print("🔧 Auto-correcting 'RP RPICING' header typo to 'RP PRICING'...")
        df = df.rename(columns={'RP RPICING': 'RP PRICING'})

    target_col_clean = str(target_column).strip().upper().strip("'").strip('"')

    # Verify foundational columns exist
    required_cols = ['B2B', 'RP PRICING', 'MRP', target_col_clean]
    for col in required_cols:
        if col not in df.columns:
            print(f"❌ Error: Missing required column '{col}' from this sheet.")
            print(f"Available headers in your file: {list(df.columns)}")
            return

    print("\n⚙️ Running business math calculations...")
    
    # Clean and convert data columns to floating numbers safely
    df['B2B'] = pd.to_numeric(df['B2B'], errors='coerce').fillna(0)
    df['RP PRICING'] = pd.to_numeric(df['RP PRICING'], errors='coerce').fillna(0)
    df['MRP'] = pd.to_numeric(df['MRP'], errors='coerce').fillna(0)

    # --- BUSINESS MATH ENGINE ---
    # 1. Base Target RP calculated with a strict 20% profit markup on B2B cost
    calculated_target = df['B2B'] * 1.20
    
    # 2. Safety Check: Cap Target RP at MRP if the markup runs too high
    df['Target RP (20% B2B Markup)'] = calculated_target.where(calculated_target < df['MRP'], df['MRP']).round(2)
    
    # 3. Setup flat reward tracking percentages
    df['Earn %'] = 10
    df['Redeem %'] = 12
    
    # 4. Compute true total monetary values directly from the Target RP column
    df['Earn Value'] = (df['Target RP (20% B2B Markup)'] * 0.10).round(2)
    df['Redeem Value'] = (df['Target RP (20% B2B Markup)'] * 0.12).round(2)

    # 5. Profit context metrics
    df['Current Profit Value'] = (df['RP PRICING'] - df['B2B']).round(2)
    df['Current Profit Margin %'] = ((df['Current Profit Value'] / df['RP PRICING'].replace(0, 1)) * 100).round(2)
    df['Current Profit Margin %'] = df['Current Profit Margin %'].fillna(0)

    # Reorder columns logically
    front_cols = ['BRAND', 'PRODUCT NAME', 'B2B', 'Target RP (20% B2B Markup)', 'Earn %', 'Earn Value', 'Redeem %', 'Redeem Value', 'RP PRICING', 'MRP']
    existing_front = [c for c in front_cols if c in df.columns]
    remaining_cols = [c for c in df.columns if c not in existing_front]
    df = df[existing_front + remaining_cols]

    # Save output uniquely
    output_file = project_folder / "Final_Calculated_Rewards_Sheet.xlsx"
    print(f"🗂️ Compiling all {len(df)} entries into a single output sheet...")

    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="All_Products_Rewards", index=False)
        print(f"\n✅ Success! Total reward values computed. Spreadsheet created at:\n{output_file}")
    except Exception as e:
        print(f"❌ Error writing output file: {e}")
        print("💡 Tip: Close 'Final_Calculated_Rewards_Sheet.xlsx' if it is open in Excel and run again!")

if __name__ == "__main__":
    TARGET_WORKSPACE = r"C:\Users\User\Desktop\SHREE"
    COLUMN_TO_TARGET = "PRODUCT NAME"
    
    process_excel_with_fixed_rewards(TARGET_WORKSPACE, COLUMN_TO_TARGET)