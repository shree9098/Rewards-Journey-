import pandas as pd
import abc
from abc import ABC, abstractmethod
from pathlib import Path


class generate_yearly_sales_analysis(ABC):
    def _get_excel_files(self, source_folder: Path):
        """Return a list of valid Excel files to analyze."""
        source_folder = Path(source_folder)
        if not source_folder.exists() or not source_folder.is_dir():
            return []

        all_files = sorted(source_folder.glob("*.xlsx"), key=lambda path: path.name.lower())
        return [
            f for f in all_files
            if not f.name.startswith("~$") and "Unified_Calculated_Rewards" not in f.name
        ]

    def _load_transaction_data(self, excel_file: Path):
        """Read all valid sheets from the chosen workbook and return a clean DataFrame."""
        excel_path = Path(excel_file)
        if not excel_path.exists():
            return None

        try:
            excel = pd.ExcelFile(excel_path)
        except Exception:
            return None

        frames = []
        for sheet_name in excel.sheet_names:
            try:
                sheet_df = pd.read_excel(excel_path, sheet_name=sheet_name)
                if sheet_df is None or sheet_df.empty:
                    continue

                sheet_df = self._normalize_columns(sheet_df)
                if all(col in sheet_df.columns for col in ['B2B', 'RP PRICING', 'MRP']):
                    frames.append(sheet_df)
            except Exception:
                continue

        if not frames:
            return None

        df = pd.concat(frames, ignore_index=True)
        for column in ['B2B', 'RP PRICING', 'MRP']:
            if column in df.columns:
                df[column] = pd.to_numeric(df[column], errors='coerce')

        df = self._prepare_year_column(df)
        return df

    def _compute_summary(self, df: pd.DataFrame):
        """Create grouped performance metrics for the yearly sales summary."""
        if df is None or df.empty:
            return pd.DataFrame()

        if 'YEAR' not in df.columns:
            df = self._prepare_year_column(df)

        df = df.copy()
        target_rp = df['B2B'] * 1.20
        df['TARGET_RP'] = target_rp.where(target_rp < df['MRP'], df['MRP']).round(2)
        df['CURRENT_PROFIT'] = (df['RP PRICING'] - df['B2B']).round(2)

        summary_by_year = df.groupby('YEAR', dropna=False).agg(
            Total_Products=('PRODUCT NAME', 'count'),
            Avg_Cost_B2B=('B2B', 'mean'),
            Avg_Current_Retail_RP=('RP PRICING', 'mean'),
            Avg_Max_Retail_MRP=('MRP', 'mean'),
            Projected_Target_RP_Avg=('TARGET_RP', 'mean'),
            Total_Current_Profit_Pool=('CURRENT_PROFIT', 'sum')
        ).round(2).reset_index()
        return summary_by_year

    def _save_analysis(self, output_folder: Path, summary_by_year: pd.DataFrame, df: pd.DataFrame):
        """Write the output workbook and return the file path."""
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        output_file = output_folder / "YEARLY_SALES_ANALYSIS.xlsx"

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            summary_by_year.to_excel(writer, sheet_name='Yearly_Sales_Comparison', index=False)
            df.to_excel(writer, sheet_name='Analyzed_Raw_Data', index=False)

        return output_file

    def run(self, source_folder_path, output_folder_path) -> None:
        source_folder = Path(source_folder_path)
        output_folder = Path(output_folder_path)

        if not source_folder.exists():
            print(f"❌ Error: The source folder path '{source_folder_path}' does not exist.")
            return

        excel_files = self._get_excel_files(source_folder)
        if not excel_files:
            print(f"❌ Error: No clean raw Excel source files (.xlsx) found inside {source_folder_path}")
            return

        excel_file = excel_files[0]
        print(f"📂 Analyzing Data for Sales Trends: {excel_file.name}")

        df = self._load_transaction_data(excel_file)
        if df is None or df.empty:
            print("❌ Error: No valid transaction data found with B2B, RP PRICING, and MRP layouts.")
            return

        summary_by_year = self._compute_summary(df)
        output_file = self._save_analysis(output_folder, summary_by_year, df)
        print(f"\n✅ Success! Yearly sales comparison generated at:\n{output_file}")

    def _find_valid_excel_files(self, source_folder: Path):
        all_files = list(source_folder.glob("*.xlsx"))
        return [f for f in all_files if not f.name.startswith("~$") and "Unified_Calculated_Rewards" not in f.name]

    def _normalize_columns(self, df: pd.DataFrame):
        df = df.copy()
        df.columns = [str(c).strip().upper().strip("'").strip('"') for c in df.columns]
        if 'RP RPICING' in df.columns:
            df = df.rename(columns={'RP RPICING': 'RP PRICING'})
        return df

    def _prepare_year_column(self, df: pd.DataFrame):
        df = df.copy()
        if 'YEAR' not in df.columns and 'DATE' not in df.columns:
            print("💡 No intrinsic date matrix found. Evenly assigning records across 2025 and 2026 for trend tracking...")
            df['YEAR'] = [2025 if i % 2 == 0 else 2026 for i in range(len(df))]
        elif 'YEAR' not in df.columns and 'DATE' in df.columns:
            df['YEAR'] = pd.to_datetime(df['DATE'], errors='coerce').dt.year
            df['YEAR'] = df['YEAR'].fillna(2026).astype(int)
        return df


class SalesYearlyAnalysis(generate_yearly_sales_analysis):
    def _get_excel_files(self, source_folder: Path):
        return self._find_valid_excel_files(source_folder)

    def _load_transaction_data(self, excel_file: Path):
        xl = pd.ExcelFile(excel_file)
        master_df = []

        for sheet in xl.sheet_names:
            try:
                sheet_df = pd.read_excel(excel_file, sheet_name=sheet)
                sheet_df = self._normalize_columns(sheet_df)
                if all(col in sheet_df.columns for col in ['B2B', 'RP PRICING', 'MRP']):
                    master_df.append(sheet_df)
            except Exception as e:
                print(f"⚠️ Skipping tab '{sheet}' due to formatting issues: {e}")

        if not master_df:
            return None

        df = pd.concat(master_df, ignore_index=True)
        df['B2B'] = pd.to_numeric(df['B2B'], errors='coerce')
        df['RP PRICING'] = pd.to_numeric(df['RP PRICING'], errors='coerce')
        df['MRP'] = pd.to_numeric(df['MRP'], errors='coerce')
        df = self._prepare_year_column(df)
        return df

    def _compute_summary(self, df: pd.DataFrame):
        calculated_target = df['B2B'] * 1.20
        df['TARGET_RP'] = calculated_target.where(calculated_target < df['MRP'], df['MRP']).round(2)
        df['CURRENT_PROFIT'] = (df['RP PRICING'] - df['B2B']).round(2)

        summary_by_year = df.groupby('YEAR').agg(
            Total_Products=('PRODUCT NAME', 'count'),
            Avg_Cost_B2B=('B2B', 'mean'),
            Avg_Current_Retail_RP=('RP PRICING', 'mean'),
            Avg_Max_Retail_MRP=('MRP', 'mean'),
            Projected_Target_RP_Avg=('TARGET_RP', 'mean'),
            Total_Current_Profit_Pool=('CURRENT_PROFIT', 'sum')
        ).round(2).reset_index()
        return summary_by_year

    def _save_analysis(self, output_folder: Path, summary_by_year: pd.DataFrame, df: pd.DataFrame):
        output_folder.mkdir(parents=True, exist_ok=True)
        output_file = output_folder / "YUONE_SALES_ANALYSIS_2025_2026.xlsx"

        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                summary_by_year.to_excel(writer, sheet_name="Yearly_Sales_Comparison", index=False)
                df.to_excel(writer, sheet_name="Analyzed_Raw_Data", index=False)
        except Exception as e:
            print(f"❌ Error writing output document: {e}")
            return output_file

        return output_file


if __name__ == "__main__":
    RAW_DATA_FOLDER = r"C:\Users\User\Desktop\SHREE"
    SCRIPT_OUTPUT_FOLDER = r"C:\Users\User\Desktop\SHREE"

    SalesYearlyAnalysis().run(RAW_DATA_FOLDER, SCRIPT_OUTPUT_FOLDER)