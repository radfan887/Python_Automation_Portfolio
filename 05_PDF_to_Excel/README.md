
# 📄 PDF to Excel Converter

Extract tables from PDF files and convert them to Excel spreadsheets automatically.

## ✨ Features

- 📑 Extract tables from any PDF
- 🔢 Handle multiple tables per PDF
- 📊 Save each table as separate Excel sheet
- 📈 Generate summary of all tables
- 🔧 Simple command-line interface

## 🔧 Installation

```bash
# Install required libraries
pip install pandas openpyxl tabula-py

# Note: tabula-py requires Java to be installed
# Download Java from: https://www.java.com/download/
🚀 Usage
Basic usage:
bash
python pdf_to_excel.py your_file.pdf
Specify output file:
bash
python pdf_to_excel.py invoice.pdf my_invoice_data.xlsx
📋 Output Structure
The Excel file contains:

Table_1, Table_2, ... - Each extracted table

Summary - Overview of all tables found

💼 Business Use Cases
💰 Invoices: Convert PDF invoices to Excel for accounting

📊 Reports: Extract data from PDF reports

📋 Forms: Convert PDF forms to editable Excel

📈 Financial Statements: Extract tables from financial PDFs

🏦 Bank Statements: Convert bank statements to Excel

🎯 Target Clients
🎯 Target Clients
Accountants and bookkeepers

Financial analysts

Business owners

Data entry professionals

Anyone working with PDF data

⚠️ Requirements
Python 3.6+

Java Runtime Environment (JRE)

Internet connection for first run (downloads Java PDF library)

👨‍💻 Author
Radfan - Python Automation Specialist
