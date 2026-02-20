
```python
"""
Advanced Smart Excel Cleaner
Professional Python Automation Tool for Excel Processing
Author: Radfan
"""

import pandas as pd
import os
import argparse
from datetime import datetime
import glob

class ExcelCleaner:
    def __init__(self):
        self.report = {}
        
    def clean_file(self, file_path, sort_column=None, output_folder="output"):
        """
        Advanced cleaning with statistics and multiple options
        """
        try:
            # Create output folder if not exists
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            # Read file
            print(f"\n📂 Processing: {os.path.basename(file_path)}")
            df = pd.read_excel(file_path)
            
            # Store original info
            self.report['filename'] = os.path.basename(file_path)
            self.report['original_rows'] = df.shape[0]
            self.report['original_columns'] = df.shape[1]
            self.report['timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Remove completely empty rows
            empty_rows_before = df.shape[0]
            df.dropna(how='all', inplace=True)
            self.report['empty_rows_removed'] = empty_rows_before - df.shape[0]

            # Remove completely empty columns
            empty_cols_before = df.shape[1]
            df.dropna(axis=1, how='all', inplace=True)
            self.report['empty_columns_removed'] = empty_cols_before - df.shape[1]

            # Remove duplicates
            before_duplicates = df.shape[0]
            df.drop_duplicates(inplace=True)
            self.report['duplicates_removed'] = before_duplicates - df.shape[0]

            # Handle missing values
            missing_values_before = df.isna().sum().sum()
            df.fillna("N/A", inplace=True)
            self.report['missing_values_filled'] = missing_values_before

            # Sort if column specified
            if sort_column and sort_column in df.columns:
                df.sort_values(by=sort_column, inplace=True)
                self.report['sorted_by'] = sort_column
            else:
                self.report['sorted_by'] = "Not applied"

            # Final stats
            self.report['final_rows'] = df.shape[0]
            self.report['final_columns'] = df.shape[1]
            self.report['rows_removed'] = self.report['original_rows'] - df.shape[0]

            # Generate statistics
            stats = df.describe(include='all')

            # Save files
            output_file = os.path.join(output_folder, f"cleaned_{os.path.basename(file_path)}")
            
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Cleaned Data")
                stats.to_excel(writer, sheet_name="Statistics")
                
                # Create summary dataframe
                summary_df = pd.DataFrame([self.report])
                summary_df.to_excel(writer, index=False, sheet_name="Summary Report")

            print(f"✅ Completed: {os.path.basename(file_path)}")
            print(f"   Rows: {self.report['original_rows']} → {self.report['final_rows']}")
            print(f"   Removed: {self.report['rows_removed']} rows")
            
            return output_file, self.report

        except Exception as e:
            print(f"❌ Error processing {file_path}: {e}")
            return None, None

    def process_folder(self, folder_path, sort_column=None, output_folder="output"):
        """
        Process all Excel files in a folder
        """
        excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
        excel_files.extend(glob.glob(os.path.join(folder_path, "*.xls")))
        
        if not excel_files:
            print("❌ No Excel files found in folder")
            return
        
        print(f"📁 Found {len(excel_files)} Excel files")
        
        all_reports = []
        for file in excel_files:
            _, report = self.clean_file(file, sort_column, output_folder)
            if report:
                all_reports.append(report)
        
        # Save summary report for all files
        if all_reports:
            summary_df = pd.DataFrame(all_reports)
            summary_file = os.path.join(output_folder, "batch_processing_summary.xlsx")
            summary_df.to_excel(summary_file, index=False)
            print(f"\n📊 Batch summary saved to: {summary_file}")

def main():
    parser = argparse.ArgumentParser(description="Advanced Excel Cleaning Tool")
    parser.add_argument("path", help="Path to Excel file or folder")
    parser.add_argument("--sort", help="Column name to sort by", required=False)
    parser.add_argument("--output", help="Output folder name", default="output")
    
    args = parser.parse_args()
    
    cleaner = ExcelCleaner()
    
    if os.path.isfile(args.path):
        # Process single file
        cleaner.clean_file(args.path, args.sort, args.output)
    elif os.path.isdir(args.path):
        # Process folder
        cleaner.process_folder(args.path, args.sort, args.output)
    else:
        print("❌ Invalid path")

if __name__ == "__main__":
    main()
