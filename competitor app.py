import os
import pandas as pd

# Define target company keyword mappings (flexible matching)
COMPANY_TARGETS = {
    "白马": "白马公司",
    "SR": "SR公司",
    "sr": "SR公司"
}

def find_header_row(df_raw):
    """Dynamically search for the row that looks like a header."""
    for idx, row in df_raw.head(10).iterrows():
        # Check if row contains non-null values and valid text
        non_nulls = row.dropna().tolist()
        if len(non_nulls) >= 2:  # Adjust threshold based on your columns
            return idx
    return 0

def process_excel_file(file_path, summary_data):
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        print(f"Reading file: {os.path.basename(file_path)} | Sheets found: {sheet_names}")
        
        for sheet in sheet_names:
            clean_sheet_name = sheet.strip()
            matched_company = None
            
            # Check against target keywords
            for keyword, mapped_name in COMPANY_TARGETS.items():
                if keyword.lower() in clean_sheet_name.lower():
                    matched_company = mapped_name
                    break
            
            if not matched_company:
                print(f"  [Skipped] Sheet '{sheet}' does not match target keywords.")
                continue
            
            # Read sheet without header first to detect format
            df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
            if df_raw.empty:
                print(f"  [Warning] Sheet '{sheet}' is completely empty!")
                continue
                
            # Locate actual header row
            header_idx = find_header_row(df_raw)
            df = pd.read_excel(file_path, sheet_name=sheet, header=header_idx)
            
            # Forward fill for merged cell issues and clean column names
            df.columns = [str(col).strip() for col in df.columns]
            df = df.dropna(how="all")  # Remove completely blank rows
            
            if df.empty:
                print(f"  [Warning] Sheet '{sheet}' has no valid data after header row {header_idx}.")
                continue
                
            df["归属公司"] = matched_company
            df["来源Sheet"] = sheet
            summary_data.append(df)
            print(f"  [Success] Successfully processed '{sheet}' as '{matched_company}' ({len(df)} rows).")
            
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")

# Main execution loop
def build_summary(file_paths, output_path):
    summary_list = []
    
    for path in file_paths:
        process_excel_file(path, summary_list)
        
    if summary_list:
        final_df = pd.concat(summary_list, ignore_index=True)
        final_df.to_excel(output_path, index=False)
        print(f"\nSummary successfully created at: {output_path} | Total rows: {len(final_df)}")
    else:
        print("\nNo matching data was found across all specified files/sheets.")

# Example usage:
# files_to_process = ["data1.xlsx", "data2.xlsx"]
# build_summary(files_to_process, "Summary_Output.xlsx")