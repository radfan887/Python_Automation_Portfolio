
"""
Excel Cleaner GUI Version
Professional Desktop Application for Excel Cleaning
Author: Radfan
"""

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from datetime import datetime
import threading

class ExcelCleanerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Excel Cleaner Pro - Radfan")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        
        # Variables
        self.file_path = tk.StringVar()
        self.sort_column = tk.StringVar()
        self.output_folder = tk.StringVar(value="output")
        
        # Setup GUI
        self.setup_ui()
        
    def setup_ui(self):
        # Title
        title_label = tk.Label(
            self.root, 
            text="📊 Smart Excel Cleaner Professional", 
            font=("Arial", 18, "bold"),
            fg="blue"
        )
        title_label.pack(pady=10)
        
        # Subtitle
        subtitle_label = tk.Label(
            self.root,
            text="Clean, Organize, and Analyze Excel Files Automatically",
            font=("Arial", 10)
        )
        subtitle_label.pack()
        
        # Main Frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File Selection
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_file).grid(row=0, column=2)
        
        # Options Frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(options_frame, text="Sort by Column (optional):").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(options_frame, textvariable=self.sort_column, width=30).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(options_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(options_frame, textvariable=self.output_folder, width=30).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        self.clean_button = ttk.Button(
            button_frame, 
            text="🧹 Clean Excel File", 
            command=self.start_cleaning,
            width=20
        )
        self.clean_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="Clear All",
            command=self.clear_all,
            width=20
        ).pack(side=tk.LEFT, padx=5)
        
        # Progress Bar
        self.progress = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            length=400
        )
        self.progress.pack(pady=10)
        
        # Log Area
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=10,
            width=70,
            wrap=tk.WORD,
            font=("Courier", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Status Bar
        self.status_label = tk.Label(
            self.root,
            text="Ready",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            self.log(f"📂 Selected file: {os.path.basename(filename)}")
            
    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def clear_all(self):
        self.file_path.set("")
        self.sort_column.set("")
        self.log_text.delete(1.0, tk.END)
        self.status_label.config(text="Ready")
        self.log("🧹 Cleared all fields")
        
    def start_cleaning(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Please select an Excel file")
            return
            
        # Disable button during processing
        self.clean_button.config(state='disabled')
        self.progress.start()
        self.status_label.config(text="Processing...")
        
        # Run cleaning in separate thread
        thread = threading.Thread(target=self.clean_excel)
        thread.start()
        
    def clean_excel(self):
        try:
            file_path = self.file_path.get()
            sort_col = self.sort_column.get() if self.sort_column.get() else None
            output_folder = self.output_folder.get()
            
            self.log("\n" + "="*50)
            self.log("🚀 Starting Excel Cleaning Process")
            self.log("="*50)
            
            # Create output folder
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                self.log(f"📁 Created output folder: {output_folder}")
            
            # Read file
            self.log(f"📂 Reading file: {os.path.basename(file_path)}")
            df = pd.read_excel(file_path)
            self.log(f"📊 Original shape: {df.shape[0]} rows, {df.shape[1]} columns")
            
            # Remove empty rows
            empty_rows_before = df.shape[0]
            df.dropna(how='all', inplace=True)
            empty_rows_removed = empty_rows_before - df.shape[0]
            self.log(f"🗑️ Removed {empty_rows_removed} completely empty rows")
            
            # Remove empty columns
            empty_cols_before = df.shape[1]
            df.dropna(axis=1, how='all', inplace=True)
            empty_cols_removed = empty_cols_before - df.shape[1]
            self.log(f"🗑️ Removed {empty_cols_removed} completely empty columns")
            
            # Remove duplicates
            before_duplicates = df.shape[0]
            df.drop_duplicates(inplace=True)
            duplicates_removed = before_duplicates - df.shape[0]
            self.log(f"🔄 Removed {duplicates_removed} duplicate rows")
            
            # Fill missing values
            missing_before = df.isna().sum().sum()
            df.fillna("N/A", inplace=True)
            self.log(f"📝 Filled {missing_before} missing values with 'N/A'")
            
            # Sort if specified
            if sort_col:
                if sort_col in df.columns:
                    df.sort_values(by=sort_col, inplace=True)
                    self.log(f"🔤 Sorted data by column: {sort_col}")
                else:
                    self.log(f"⚠️ Column '{sort_col}' not found, skipping sort")
            
            # Final stats
            self.log(f"✅ Final shape: {df.shape[0]} rows, {df.shape[1]} columns")
            self.log(f"📉 Removed {df.shape[0] - before_duplicates} rows total")
            
            # Save file
            output_file = os.path.join(output_folder, f"cleaned_{os.path.basename(file_path)}")
            df.to_excel(output_file, index=False)
            self.log(f"💾 Saved cleaned file: {output_file}")
            
            # Create summary report
            summary = {
                "Filename": os.path.basename(file_path),
                "Original Rows": before_duplicates,
                "Final Rows": df.shape[0],
                "Rows Removed": before_duplicates - df.shape[0],
                "Original Columns": empty_cols_before,
                "Final Columns": df.shape[1],
                "Duplicates Removed": duplicates_removed,
                "Missing Values Filled": missing_before,
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            summary_file = os.path.join(output_folder, f"summary_{os.path.basename(file_path).replace('.xlsx', '.txt')}")
            with open(summary_file, 'w') as f:
                for key, value in summary.items():
                    f.write(f"{key}: {value}\n")
                    self.log(f"📊 {key}: {value}")
            
            self.log("="*50)
            self.log("✅ Cleaning completed successfully!")
            self.log("="*50 + "\n")
            
            # Show success message
            self.root.after(0, lambda: self.show_success(output_file))
            
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            
        finally:
            # Re-enable button
            self.root.after(0, lambda: self.clean_button.config(state='normal'))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.status_label.config(text="Ready"))
    
    def show_success(self, output_file):
        messagebox.showinfo(
            "Success", 
            f"Excel file cleaned successfully!\n\nSaved to: {output_file}"
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCleanerGUI(root)
    root.mainloop()
