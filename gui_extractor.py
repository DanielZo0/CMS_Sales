#!/usr/bin/env python3
"""
CMZ Data Extractor with GUI
GUI-based data extractor for CMZ sales and purchases data
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Import the existing extractor
from combined_data_extractor import CombinedDataExtractor


class CMZExtractorGUI:
    """GUI interface for CMZ Data Extractor"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CMZ Data Extractor")
        self.root.geometry("800x750")
        self.root.resizable(True, True)
        
        # Initialize variables
        self.input_dir = tk.StringVar()  # Single input directory for both PDFs and Excel files
        self.sales_template_path = tk.StringVar()
        self.purchases_template_path = tk.StringVar()
        self.upload_template_path = tk.StringVar()
        
        # Set default values
        self.sales_template_path.set("templates/JSON_Template_Sales.json")
        self.purchases_template_path.set("templates/JSON_Template_Purchases.json")
        self.upload_template_path.set("templates/JSON_Template_Upload.json")
        
        self.create_widgets()
        self.center_window()
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def create_widgets(self):
        """Create and layout GUI widgets"""
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="CMZ Data Extractor", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        row = 1
        
        # Input Directory Section
        ttk.Label(main_frame, text="Input Directory:", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=3, 
                                                 sticky=tk.W, pady=(10, 5))
        row += 1
        
        # Single Input Directory (contains both PDFs and Excel files)
        ttk.Label(main_frame, text="Data Files Directory:").grid(row=row, column=0, 
                                                                 sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.input_dir, width=50).grid(
            row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        ttk.Button(main_frame, text="Browse...", 
                  command=lambda: self.browse_directory(self.input_dir)).grid(
            row=row, column=2, pady=2)
        row += 1
        
        # Info about what should be in the directory
        info_text = "Directory should contain: Sales PDFs and Purchases Excel files"
        ttk.Label(main_frame, text=info_text, font=('Arial', 9), 
                 foreground='gray').grid(row=row, column=0, columnspan=3, 
                                       sticky=tk.W, pady=(2, 10))
        row += 1
        
        # Template Files Section
        ttk.Label(main_frame, text="Template Files:", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=3, 
                                                 sticky=tk.W, pady=(20, 5))
        row += 1
        
        # Sales Template
        ttk.Label(main_frame, text="Sales Template:").grid(row=row, column=0, 
                                                           sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.sales_template_path, width=50).grid(
            row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        ttk.Button(main_frame, text="Browse...", 
                  command=lambda: self.browse_file(self.sales_template_path, 
                                                 "JSON files", "*.json")).grid(
            row=row, column=2, pady=2)
        row += 1
        
        # Purchases Template
        ttk.Label(main_frame, text="Purchases Template:").grid(row=row, column=0, 
                                                               sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.purchases_template_path, width=50).grid(
            row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        ttk.Button(main_frame, text="Browse...", 
                  command=lambda: self.browse_file(self.purchases_template_path, 
                                                 "JSON files", "*.json")).grid(
            row=row, column=2, pady=2)
        row += 1
        
        # Upload Template
        ttk.Label(main_frame, text="Upload Template:").grid(row=row, column=0, 
                                                            sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.upload_template_path, width=50).grid(
            row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        ttk.Button(main_frame, text="Browse...", 
                  command=lambda: self.browse_file(self.upload_template_path, 
                                                 "JSON files", "*.json")).grid(
            row=row, column=2, pady=2)
        row += 1
        
        # Processing Options
        ttk.Label(main_frame, text="Processing Options:", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=3, 
                                                 sticky=tk.W, pady=(20, 5))
        row += 1
        
        # Checkboxes for processing options
        self.rename_pdfs = tk.BooleanVar(value=True)
        self.generate_individual = tk.BooleanVar(value=True)
        self.generate_combined = tk.BooleanVar(value=True)
        self.generate_upload = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(main_frame, text="Rename PDF files to standard format", 
                       variable=self.rename_pdfs).grid(row=row, column=0, columnspan=3, 
                                                      sticky=tk.W, pady=2)
        row += 1
        
        ttk.Checkbutton(main_frame, text="Generate individual data files (JSON/CSV)", 
                       variable=self.generate_individual).grid(row=row, column=0, columnspan=3, 
                                                              sticky=tk.W, pady=2)
        row += 1
        
        ttk.Checkbutton(main_frame, text="Generate combined sales/purchases CSV", 
                       variable=self.generate_combined).grid(row=row, column=0, columnspan=3, 
                                                            sticky=tk.W, pady=2)
        row += 1
        
        ttk.Checkbutton(main_frame, text="Generate upload CSV format (based on template)", 
                       variable=self.generate_upload).grid(row=row, column=0, columnspan=3, 
                                                          sticky=tk.W, pady=2)
        row += 1
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                           maximum=100)
        self.progress_bar.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), 
                              pady=(20, 5))
        row += 1
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to process")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="Start", 
                  command=self.process_data, style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear All", 
                  command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", 
                  command=self.root.quit).pack(side=tk.LEFT, padx=5)
    
    def browse_directory(self, var: tk.StringVar):
        """Browse for directory"""
        directory = filedialog.askdirectory(
            title="Select Directory",
            initialdir=var.get() if var.get() else os.getcwd()
        )
        if directory:
            var.set(directory)
    
    def browse_file(self, var: tk.StringVar, file_type: str, pattern: str):
        """Browse for file"""
        file_path = filedialog.askopenfilename(
            title=f"Select {file_type}",
            filetypes=[(file_type, pattern), ("All files", "*.*")],
            initialdir=os.path.dirname(var.get()) if var.get() else os.getcwd()
        )
        if file_path:
            var.set(file_path)
    
    def clear_all(self):
        """Clear all input fields"""
        self.input_dir.set("")
        self.progress_var.set(0)
        self.status_var.set("Ready to process")
    
    def validate_inputs(self) -> Tuple[bool, str]:
        """Validate user inputs"""
        errors = []
        
        # Check if input directory is provided
        if not self.input_dir.get():
            errors.append("Input directory must be specified")
        elif not os.path.exists(self.input_dir.get()):
            errors.append(f"Input directory does not exist: {self.input_dir.get()}")
        
        # Check template files
        for name, path in [("Sales template", self.sales_template_path.get()),
                          ("Purchases template", self.purchases_template_path.get()),
                          ("Upload template", self.upload_template_path.get())]:
            if path and not os.path.exists(path):
                errors.append(f"{name} file does not exist: {path}")
        
        if errors:
            return False, "\n".join(errors)
        return True, ""
    
    def update_progress(self, value: float, status: str):
        """Update progress bar and status"""
        self.progress_var.set(value)
        self.status_var.set(status)
        self.root.update_idletasks()
    
    def process_data(self):
        """Process the data with selected options"""
        # Validate inputs
        is_valid, error_msg = self.validate_inputs()
        if not is_valid:
            messagebox.showerror("Validation Error", error_msg)
            return
        
        try:
            self.update_progress(0, "Initializing...")
            
            # Create output directory in the input directory
            output_dir = os.path.join(self.input_dir.get(), "output")
            os.makedirs(output_dir, exist_ok=True)
            
            # Initialize extractor
            extractor = CombinedDataExtractor(
                sales_template_path=self.sales_template_path.get(),
                purchases_template_path=self.purchases_template_path.get()
            )
            
            self.update_progress(10, "Extractor initialized...")
            
            # Rename PDF files if requested
            if self.rename_pdfs.get() and self.input_dir.get():
                self.update_progress(20, "Renaming PDF files...")
                pdf_subdir = os.path.join(self.input_dir.get(), "WeeklySales")
                if os.path.exists(pdf_subdir):
                    extractor.rename_pdf_files(pdf_subdir)
                else:
                    # Look for PDFs in the main input directory
                    extractor.rename_pdf_files(self.input_dir.get())
            
            # Process all files
            self.update_progress(30, "Processing data files...")
            results = extractor.process_all_files(
                sales_input_dir=None,  # No sales Excel files (only purchases Excel)
                purchases_input_dir=self.input_dir.get(),  # Excel files for purchases 
                pdf_dir=self.input_dir.get()  # Look for PDFs in main directory or WeeklySales subdirectory
            )
            
            total_records = sum(len(data) for data in results.values())
            
            if total_records == 0:
                messagebox.showwarning("No Data", "No data was extracted from the selected files.")
                self.update_progress(0, "No data processed")
                return
            
            # Save individual files if requested
            if self.generate_individual.get():
                self.update_progress(60, "Saving individual data files...")
                extractor.save_results(results, output_dir)
            
            # Generate combined CSV if requested
            if self.generate_combined.get():
                self.update_progress(70, "Generating combined CSV...")
                # Update the generate_combined_csv to use the new output directory
                self.generate_combined_csv_with_custom_output(output_dir)
            
            # Generate upload CSV if requested
            if self.generate_upload.get():
                self.update_progress(85, "Generating upload CSV...")
                self.generate_upload_csv_with_custom_output(output_dir)
            
            # Create renamed invoices folder if PDF renaming was requested
            if self.rename_pdfs.get() and self.input_dir.get():
                self.update_progress(90, "Creating renamed invoices folder...")
                extractor.create_renamed_invoices_folder(self.input_dir.get(), output_dir)
            
            self.update_progress(100, f"Processing complete! {total_records} records processed")
            
            # Show success message
            messagebox.showinfo("Success", 
                              f"Processing completed successfully!\n\n"
                              f"Records processed: {total_records}\n"
                              f"Output location: {output_dir}")
            
        except Exception as e:
            error_msg = f"An error occurred during processing:\n{str(e)}"
            messagebox.showerror("Processing Error", error_msg)
            self.update_progress(0, "Error occurred")
            print(f"Error details: {e}")
            import traceback
            traceback.print_exc()
    
    def generate_combined_csv_with_custom_output(self, output_dir: str):
        """Generate combined CSV with custom output directory"""
        try:
            # Import the function from generate_combined_csv
            import generate_combined_csv
            
            # Temporarily modify the file paths
            original_generate = generate_combined_csv.generate_combined_csv
            
            def custom_generate():
                # File paths using the custom output directory
                sales_file = os.path.join(output_dir, "pdf_sales_data.csv")
                purchases_file = os.path.join(output_dir, "purchases_data.csv")
                output_file = os.path.join(output_dir, "combined_sales_purchases.csv")
                
                # Load data using the original functions
                sales_records = generate_combined_csv.load_and_standardize_sales_data(sales_file)
                purchases_records = generate_combined_csv.load_and_standardize_purchases_data(purchases_file)
                
                # Combine records
                all_records = sales_records + purchases_records
                
                if not all_records:
                    print("No data to combine")
                    return False
                
                # Sort by date
                all_records = generate_combined_csv.sort_records_by_date(all_records)
                
                # Define field order
                fieldnames = [
                    'Data Type', 'Document Type', 'Document Date', 'Supplier Code',
                    'Empty_Column', 'Document Number', 'Description', 'NC', 'VC',
                    'Locality', 'Net', 'VAT', 'Total'
                ]
                
                # Save combined CSV
                import csv
                with open(output_file, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(all_records)
                
                print(f"Combined CSV saved to: {output_file}")
                return True
            
            return custom_generate()
            
        except Exception as e:
            print(f"Error generating combined CSV: {e}")
            return False
    
    def generate_upload_csv_with_custom_output(self, output_dir: str):
        """Generate upload CSV with custom output directory"""
        try:
            import csv
            import json
            
            # File paths using the custom output directory
            sales_file = os.path.join(output_dir, "pdf_sales_data.csv")
            purchases_file = os.path.join(output_dir, "purchases_data.csv")
            upload_file = os.path.join(output_dir, "upload_data.csv")
            upload_template_path = self.upload_template_path.get()
            
            # Load upload template
            try:
                with open(upload_template_path, 'r', encoding='utf-8') as f:
                    template_list = json.load(f)
                    upload_template = template_list[0] if template_list else {}
            except FileNotFoundError:
                print(f"Upload template file {upload_template_path} not found. Using default structure.")
                upload_template = {
                    "Type": "",
                    "Account Reference": "supplier_code",
                    "Nominal A/C Ref": "NC_extract",
                    "Department Code": "",
                    "Date": "extract_date",
                    "Reference": "extract_reference",
                    "Details": "extract_description",
                    "Net Amount": "extract_net",
                    "Tax Code": "T0",
                    "Tax Amount": "extract_vat"
                }
            
            if not upload_template:
                print("Error: Could not load upload template")
                return False
            
            print("Generating upload format CSV...")
            upload_records = []
            
            # Import date formatting function
            from generate_combined_csv import format_date_to_ddmmyyyy
            
            # Process sales data
            if os.path.exists(sales_file):
                with open(sales_file, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        upload_record = {}
                        
                        for field, mapping in upload_template.items():
                            if mapping == "supplier_code":
                                upload_record[field] = row.get('Supplier Code', '')
                            elif mapping == "NC_extract":
                                upload_record[field] = row.get('NC', '')
                            elif mapping == "extract_date":
                                raw_date = row.get('Document Date', '')
                                upload_record[field] = format_date_to_ddmmyyyy(raw_date)
                            elif mapping == "extract_reference":
                                upload_record[field] = row.get('Document Number', '')
                            elif mapping == "extract_description":
                                upload_record[field] = row.get('Description', '')
                            elif mapping == "extract_net":
                                # For sales: Net = Total, VAT = 0
                                total_str = str(row.get('Total', '')).replace(',', '')
                                try:
                                    total = float(total_str) if total_str else 0
                                    upload_record[field] = f"{total:.2f}" if total > 0 else ''
                                except ValueError:
                                    upload_record[field] = ''
                            elif mapping == "extract_vat":
                                # For sales: VAT is always 0
                                upload_record[field] = "0.00"
                            else:
                                # Static value or empty
                                upload_record[field] = mapping
                        
                        upload_records.append(upload_record)
            
            # Process purchases data
            if os.path.exists(purchases_file):
                with open(purchases_file, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        upload_record = {}
                        
                        for field, mapping in upload_template.items():
                            if mapping == "supplier_code":
                                upload_record[field] = row.get('Supplier Code', '')
                            elif mapping == "NC_extract":
                                upload_record[field] = row.get('NC', '')
                            elif mapping == "extract_date":
                                raw_date = row.get('Document Date', '')
                                upload_record[field] = format_date_to_ddmmyyyy(raw_date)
                            elif mapping == "extract_reference":
                                upload_record[field] = row.get('Document Number', '')
                            elif mapping == "extract_description":
                                upload_record[field] = row.get('Description', '')
                            elif mapping == "extract_net":
                                upload_record[field] = row.get('Net', '')
                            elif mapping == "extract_vat":
                                upload_record[field] = row.get('VAT', '')
                            else:
                                # Static value or empty
                                upload_record[field] = mapping
                        
                        upload_records.append(upload_record)
            
            # Save upload CSV
            if upload_records:
                with open(upload_file, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=list(upload_template.keys()))
                    writer.writeheader()
                    writer.writerows(upload_records)
                
                print(f"Upload CSV saved to: {upload_file}")
                print(f"Upload records: {len(upload_records)}")
                return True
            else:
                print("No records to write to upload CSV")
                return False
                
        except Exception as e:
            print(f"Error generating upload CSV: {e}")
            return False
    
    def run(self):
        """Run the GUI application"""
        self.root.mainloop()


def main():
    """Main function to run the GUI"""
    try:
        app = CMZExtractorGUI()
        app.run()
    except Exception as e:
        print(f"Error starting GUI: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
