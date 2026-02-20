
#  Advanced Smart Excel Cleaner

Professional Python automation tool for batch processing Excel files with advanced statistics.

##  Features

- **Batch Processing** - Clean multiple files at once
-  **Detailed Statistics** - Get insights about your data
-  **Flexible Sorting** - Sort by any column
-  **Empty Data Removal** - Remove empty rows and columns
- **Duplicate Detection** - Find and remove duplicates
-  **Missing Value Handling** - Fill missing values
-  **Summary Reports** - Generate cleaning reports
-  **Statistical Analysis** - Get data statistics automatically

##  Installation

```bash
pip install pandas openpyxl

Usage
Process single file:
bash
python main.py data.xlsx --sort Name
Process all Excel files in folder:
bash
python main.py ./my_excel_files --sort Date
Specify output folder:
bash
python main.py data.xlsx --sort Name --output cleaned_data
 Output Structure
text
output/
├── cleaned_data.xlsx
├── cleaned_another_file.xlsx
└── batch_processing_summary.xlsx
Each cleaned file contains:

Cleaned Data: Processed data

Statistics: Statistical summary
Summary Report: Cleaning operation details

 Business Value
This tool saves hours of manual Excel cleaning by:

Automating repetitive tasks

Ensuring consistency

Providing quality reports

Processing hundreds of files automatically

 Author
Radfan - Python Automation Specialist
