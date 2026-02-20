# 🔗 Excel File Merger

Merge multiple Excel files into one master file with source tracking.

## ✨ Features

- 📂 Merge all Excel files in a folder
- 🔍 Track source of each row
- 📊 Generate merge summary
- 🔄 Handle different column structures
- 📈 Preserve all data from original files

## 🔧 Installation

```bash
pip install pandas openpyxl
🚀 Usage
bash
python merge.py folder_path
Example:
bash
python merge.py ./sales_reports
python merge.py C:\\Users\\Radfan\\Documents\\excels
📋 Output
The script creates merged_output.xlsx with two sheets:
1. Merged Data
All data from all files combined

Extra column 'Source_File' showing original file name

2. Summary
List of all merged files

Number of rows from each file

Column count from each file

💼 Business Use Cases
📊 Sales Reports: Merge monthly reports into yearly report

👥 Employee Data: Combine department lists
📦 Inventory: Merge multiple warehouse inventories

📈 Analytics: Combine multiple data sources for analysis

🎯 Target Clients
Accountants merging financial reports

HR managers combining employee data

Data analysts preparing datasets

Business owners consolidating reports

👨‍💻 Author
Radfan - Python Automation Specialist
