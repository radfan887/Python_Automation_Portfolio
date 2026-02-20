
"""
Excel File Merger
Combine multiple Excel files into one
Author: Radfan
"""

import pandas as pd
import os
import sys
import glob
from datetime import datetime

def merge_excel_files(folder_path, output_name="merged_output.xlsx"):
    """
    Merge all Excel files in a folder into one file
    """
    try:
        # Find all Excel files
        excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
        excel_files.extend(glob.glob(os.path.join(folder_path, "*.xls")))
        
        if not excel_files:
            print("❌ No Excel files found in the folder")
            return None
        
        print(f"📁 Found {len(excel_files)} Excel files")
        
        # Create a list to store all dataframes
        all_dfs = []
        file_names = []
        
        # Read each file and add to list
        for file in excel_files:
            try:
                print(f"📂 Reading: {os.path.basename(file)}")
                df = pd.read_excel(file)
                
                # Add filename column to track source
                df['Source_File'] = os.path.basename(file)
                
                all_dfs.append(df)
                file_names.append(os.path.basename(file))
                print(f"   ✅ {df.shape[0]} rows, {df.shape[1]} columns")
                
            except Exception as e:
                print(f"   ❌ Error reading {file}: {e}")
        
        if not all_dfs:
            print("❌ No valid Excel files to merge")
            return None
        
        # Combine all dataframes
        print("\n🔄 Merging files...")
        combined_df = pd.concat(all_dfs, ignore_index=True, sort=False)
        
        # Create summary
        summary = pd.DataFrame({
            'File Name': file_names,
            'Rows': [len(df) for df in all_dfs],
            'Columns': [len(df.columns)-1 for df in all_dfs]  # -1 for Source_File column
        })
        
        # Save merged file
        output_path = os.path.join(folder_path, output_name)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            combined_df.to_excel(writer, sheet_name='Merged Data', index=False)
            summary.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"\n✅ Merged {len(all_dfs)} files successfully!")
        print(f"📊 Total rows: {combined_df.shape[0]}")
        print(f"📊 Total columns: {combined_df.shape[1] - 1}")  # -1 for Source_File
        print(f"💾 Saved to: {output_path}")
        
        return output_path
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("📌 Usage: python merge.py folder_path")
        print("📌 Example: python merge.py ./excel_files")
        print("📌 Example: python merge.py C:\\Users\\Radfan\\Documents\\excels")
    else:
        folder = sys.argv[1]
        if os.path.exists(folder):
            merge_excel_files(folder)
        else:
            print("❌ Folder does not exist")
