
"""
PDF to Excel Converter
Extract tables from PDF and convert to Excel
Author: Radfan
"""

import pandas as pd
import os
import sys
import tabula
import warnings
warnings.filterwarnings('ignore')

def pdf_to_excel(pdf_path, output_excel=None):
    """
    Extract tables from PDF and save to Excel
    """
    try:
        print(f"📂 Reading PDF: {os.path.basename(pdf_path)}")
        
        # Extract tables from PDF
        print("🔍 Extracting tables...")
        tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
        
        if not tables:
            print("❌ No tables found in PDF")
            return None
        
        print(f"✅ Found {len(tables)} table(s)")
        
        # Generate output filename if not provided
        if not output_excel:
            output_excel = pdf_path.replace('.pdf', '_converted.xlsx')
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                sheet_name = f"Table_{i+1}"
                table.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"   📊 Table {i+1}: {table.shape[0]} rows, {table.shape[1]} columns")
        
        # Create summary sheet
        summary_data = {
            'Table Number': list(range(1, len(tables)+1)),
            'Rows': [t.shape[0] for t in tables],
            'Columns': [t.shape[1] for t in tables]
        }
        summary_df = pd.DataFrame(summary_data)
        
        # Add summary to Excel
        with pd.ExcelWriter(output_excel, engine='openpyxl', mode='a') as writer:
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"\n✅ Conversion complete!")
        print(f"💾 Saved to: {output_excel}")
        
        return output_excel
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("📌 Usage: python pdf_to_excel.py your_file.pdf [output_excel.xlsx]")
        print("📌 Example: python pdf_to_excel.py invoice.pdf")
        print("📌 Example: python pdf_to_excel.py report.pdf report.xlsx")
    else:
        pdf_file = sys.argv[1]
        if os.path.exists(pdf_file):
            output = sys.argv[2] if len(sys.argv) > 2 else None
            pdf_to_excel(pdf_file, output)
        else:
            print("❌ PDF file does not exist")
