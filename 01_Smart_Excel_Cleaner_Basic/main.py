"""
Smart Excel Cleaner - Basic Version
Python Automation Tool for Cleaning Excel Files
Author: Radfan
"""

import pandas as pd
import sys
import os

def clean_excel(file_path):
    """
    Cleans messy Excel file by:
    - Removing empty rows and columns
    - Removing duplicates
    - Filling missing values
    - Sorting data
    - Generating summary
    """
    try:
        # قراءة الملف
        print(f"\n📂 Reading file: {file_path}")
        df = pd.read_excel(file_path)
        
        print(f"📊 Original Data: {df.shape[0]} rows, {df.shape[1]} columns")

        # حذف الصفوف الفارغة بالكامل
        df.dropna(how='all', inplace=True)
        
        # حذف الأعمدة الفارغة بالكامل
        df.dropna(axis=1, how='all', inplace=True)

        # إزالة التكرار
        before_duplicates = df.shape[0]
        df.drop_duplicates(inplace=True)
        duplicates_removed = before_duplicates - df.shape[0]

        # ملء القيم الفارغة بكلمة N/A
        df.fillna("N/A", inplace=True)

        # ترتيب البيانات حسب أول عمود
        if len(df.columns) > 0:
            first_column = df.columns[0]
            df.sort_values(by=first_column, inplace=True)
            print(f"🔤 Sorted by column: {first_column}")

        # إنشاء تقرير ملخص
        summary = {
            "Original Rows": [before_duplicates],
            "Final Rows": [df.shape[0]],
            "Duplicates Removed": [duplicates_removed],
            "Columns": [df.shape[1]]
        }
        summary_df = pd.DataFrame(summary)

        # حفظ الملف الجديد
        output_file = "cleaned_" + os.path.basename(file_path)
        
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Cleaned Data")
            summary_df.to_excel(writer, index=False, sheet_name="Summary")

        print(f"✅ Cleaning completed successfully!")
        print(f"💾 Saved as: {output_file}")
        print(f"📈 Summary: {df.shape[0]} rows, {df.shape[1]} columns")
        
        return output_file

    except Exception as e:
        print(f"❌ Error: {e}")
        return None

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("📌 Usage: python main.py your_file.xlsx")
        print("📌 Example: python main.py data.xlsx")
    else:
        clean_excel(sys.argv[1])
