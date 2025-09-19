#!/usr/bin/env python3
"""
Combined Data Extractor for CMZ Sales

This script extracts data from both PDF and Excel files, renames PDF files to standard format,
and outputs structured data according to JSON template specifications.
"""

import pandas as pd
import json
import csv
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional

# Optional PDF processing
try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    print("Warning: pdfplumber not available. PDF processing will be disabled.")


class CombinedDataExtractor:
    """Unified extractor for PDF and Excel data across purchases and sales"""
    
    def __init__(self, 
                 sales_template_path: str = "templates/JSON_Template_Sales.json",
                 purchases_template_path: str = "templates/JSON_Template_Purchases.json"):
        """Initialize the extractor with JSON templates"""
        self.sales_template_path = sales_template_path
        self.purchases_template_path = purchases_template_path
        self.sales_template = self._load_template(sales_template_path, "sales")
        self.purchases_template = self._load_template(purchases_template_path, "purchases")
        self.supplier_mapping = self._create_supplier_mapping()
        self.description_mapping = self._create_description_mapping()
    
    def _load_template(self, template_path: str, template_type: str) -> List[Dict[str, str]]:
        """Load the JSON template structure"""
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Template file {template_path} not found. Using default {template_type} structure.")
            return self._get_default_template(template_type)
    
    def _get_default_template(self, template_type: str) -> List[Dict[str, str]]:
        """Return default template structure"""
        if template_type == "sales":
            return [{
                "Document Type": "",
                "Document Date": "extract_date",
                "Supplier Code": "CHAINFGU",
                "Empty_Column": "",
                "Document Number": "extract_reference",
                "Description": "extract_description",
                "NC": "4002",
                "VC": "T0",
                "Locality": "extract_locality",
                "Total": ""
            }]
        else:  # purchases
            return [{
                "Document Type": "SI",
                "Document Date": "extract_date",
                "Supplier Code": "CHAINFGU",
                "Document Number": "extract_reference",
                "Description": "extract_description",
                "NC": "4001",
                "VC": "T0",
                "Net": "extract_net",
                "VAT": "extract_vat"
            }]
    
    def _create_supplier_mapping(self) -> Dict[str, str]:
        """Create a unified mapping from locality to supplier code"""
        return {
            "Carter": "CHAINTAR",
            "Carters": "CHAINTAR",  # Carter and Tarxien are one and the same
            "Fgura": "CHAINFGU",
            "Tarxien": "CHAINTAR", 
            "Zabbar": "CHAINZAB"
        }
    
    def _create_description_mapping(self) -> Dict[str, str]:
        """Create a mapping for description locality names"""
        return {
            "Carter": "Tarxien",
            "Carters": "Tarxien",  # Carter and Tarxien are one and the same
            "Fgura": "Fgura",
            "Tarxien": "Tarxien", 
            "Zabbar": "Zabbar"
        }
    
    def _get_template_for_locality(self, locality: str, data_type: str) -> Dict[str, str]:
        """Get the template item for a specific locality and data type"""
        template = self.sales_template if data_type == "sales" else self.purchases_template
        
        if isinstance(template, list):
            supplier_code = self.supplier_mapping.get(locality, "")
            
            # Find matching template by supplier code
            for item in template:
                if "Supplier Code" in item and item["Supplier Code"] == supplier_code:
                    return item
            
            # Fallback: use first template and update supplier code
            if template:
                template_copy = template[0].copy()
                template_copy["Supplier Code"] = supplier_code
                return template_copy
        
        return {}
    
    def rename_pdf_files(self, pdf_directory: str = "input") -> None:
        """Rename PDF files based on extracted data: Weekly Sales - <Locality>"""
        if not PDF_AVAILABLE:
            print("PDF processing not available. Cannot rename PDF files.")
            return
            
        # Look for PDFs in the main directory and subdirectories
        pdf_directories = []
        if os.path.exists(pdf_directory):
            # Add main directory
            pdf_directories.append(pdf_directory)
            
            # Add subdirectories that might contain PDFs
            try:
                for item in os.listdir(pdf_directory):
                    item_path = os.path.join(pdf_directory, item)
                    if os.path.isdir(item_path):
                        # Check if subdirectory contains PDF files
                        pdf_files_in_subdir = [f for f in os.listdir(item_path) if f.endswith('.pdf')]
                        if pdf_files_in_subdir:
                            pdf_directories.append(item_path)
            except Exception as e:
                print(f"Error scanning subdirectories: {e}")
        
        if not pdf_directories:
            print(f"PDF directory not found: {pdf_directory}")
            return
        
        renamed_count = 0
        total_files = 0
        
        for search_dir in pdf_directories:
            try:
                pdf_files = [f for f in os.listdir(search_dir) if f.endswith('.pdf')]
                if not pdf_files:
                    continue
                    
                print(f"Found {len(pdf_files)} PDF files in {search_dir}...")
                total_files += len(pdf_files)
                
                for filename in pdf_files:
                    old_path = os.path.join(search_dir, filename)
                    
                    try:
                        # Extract data from the PDF to get locality
                        extracted_data = self.extract_from_pdf(old_path)
                        
                        if extracted_data:
                            # Get locality from extracted data
                            locality = extracted_data.get('Locality', '').strip()
                            
                            if locality:
                                new_filename = f"Weekly Sales - {locality}.pdf"
                                new_path = os.path.join(search_dir, new_filename)
                                
                                if filename != new_filename:
                                    if os.path.exists(new_path):
                                        print(f"  Warning: Target file already exists: {new_filename}")
                                        continue
                                    
                                    os.rename(old_path, new_path)
                                    print(f"  Renamed: {filename} -> {new_filename}")
                                    renamed_count += 1
                                else:
                                    print(f"  Already correctly named: {filename}")
                            else:
                                print(f"  Could not extract locality from: {filename}")
                        else:
                            print(f"  No data extracted from: {filename}")
                            
                    except Exception as e:
                        print(f"  Error processing {filename}: {e}")
                        
            except Exception as e:
                print(f"Error processing directory {search_dir}: {e}")
        
        print(f"\nPDF renaming complete. {renamed_count} files renamed out of {total_files} total files.")
    
    def copy_renamed_pdfs_to_output(self, pdf_directory: str = "input", output_dir: str = "output") -> None:
        """Copy PDF files with their updated names to the output folder, organized by locality"""
        if not PDF_AVAILABLE:
            print("PDF processing not available. Cannot copy PDF files.")
            return
        
        # Create main PDF directory in output
        pdf_output_dir = os.path.join(output_dir, "weekly_sales_pdfs")
        os.makedirs(pdf_output_dir, exist_ok=True)
        
        copied_count = 0
        
        # Look for PDFs in the main directory and subdirectories
        pdf_directories = []
        if os.path.exists(pdf_directory):
            # Add main directory
            pdf_directories.append(pdf_directory)
            
            # Add subdirectories that might contain PDFs
            try:
                for item in os.listdir(pdf_directory):
                    item_path = os.path.join(pdf_directory, item)
                    if os.path.isdir(item_path):
                        # Check if subdirectory contains PDF files
                        pdf_files_in_subdir = [f for f in os.listdir(item_path) if f.endswith('.pdf')]
                        if pdf_files_in_subdir:
                            pdf_directories.append(item_path)
            except Exception as e:
                print(f"Error scanning subdirectories: {e}")
        
        print(f"\nCopying PDF files to output directory...")
        
        for search_dir in pdf_directories:
            try:
                pdf_files = [f for f in os.listdir(search_dir) if f.endswith('.pdf')]
                
                for filename in pdf_files:
                    source_path = os.path.join(search_dir, filename)
                    destination_path = os.path.join(pdf_output_dir, filename)
                    
                    try:
                        # Copy the file to output directory
                        shutil.copy2(source_path, destination_path)
                        print(f"  Copied: {filename}")
                        copied_count += 1
                        
                    except Exception as e:
                        print(f"  Error copying {filename}: {e}")
                        
            except Exception as e:
                print(f"Error processing directory {search_dir}: {e}")
        
        print(f"\nPDF copying complete. {copied_count} files copied to {pdf_output_dir}")
    
    def create_renamed_invoices_folder(self, pdf_directory: str = "input", output_dir: str = "output") -> None:
        """Create Renamed Invoices folder with PDFs renamed to 'Weekly Sales - Location WkXX' format"""
        if not PDF_AVAILABLE:
            print("PDF processing not available. Cannot create renamed invoices folder.")
            return
        
        # Create Renamed Invoices directory in output
        renamed_invoices_dir = os.path.join(output_dir, "Renamed Invoices")
        os.makedirs(renamed_invoices_dir, exist_ok=True)
        
        renamed_count = 0
        
        # Look for PDFs in the main directory and subdirectories
        pdf_directories = []
        if os.path.exists(pdf_directory):
            # Add main directory
            pdf_directories.append(pdf_directory)
            
            # Add subdirectories that might contain PDFs
            try:
                for item in os.listdir(pdf_directory):
                    item_path = os.path.join(pdf_directory, item)
                    if os.path.isdir(item_path):
                        # Check if subdirectory contains PDF files
                        pdf_files_in_subdir = [f for f in os.listdir(item_path) if f.endswith('.pdf')]
                        if pdf_files_in_subdir:
                            pdf_directories.append(item_path)
            except Exception as e:
                print(f"Error scanning subdirectories: {e}")
        
        print(f"\nCreating renamed invoices in: {renamed_invoices_dir}")
        
        for search_dir in pdf_directories:
            try:
                pdf_files = [f for f in os.listdir(search_dir) if f.endswith('.pdf')]
                
                for filename in pdf_files:
                    source_path = os.path.join(search_dir, filename)
                    
                    try:
                        # Extract location and week number from filename or PDF content
                        location, week_num = self._extract_location_and_week(source_path, filename)
                        
                        if location and week_num:
                            # Create new filename in the format: Weekly Sales - Location WkXX
                            new_filename = f"Weekly Sales - {location} Wk{week_num.zfill(2)}.pdf"
                            destination_path = os.path.join(renamed_invoices_dir, new_filename)
                            
                            # Copy the file with new name
                            shutil.copy2(source_path, destination_path)
                            print(f"  Created: {new_filename}")
                            renamed_count += 1
                        else:
                            print(f"  Could not extract location/week from: {filename}")
                            
                    except Exception as e:
                        print(f"  Error processing {filename}: {e}")
                        
            except Exception as e:
                print(f"Error processing directory {search_dir}: {e}")
        
        print(f"\nRenamed invoices creation complete. {renamed_count} files created in {renamed_invoices_dir}")
    
    def _extract_location_and_week(self, pdf_path: str, filename: str) -> tuple:
        """Extract location and week number from PDF filename and content"""
        location = ""
        week_num = ""
        filename_week = ""
        
        # First try to extract from filename
        # Pattern 1: Weekly Sales - Location WkXX.pdf
        filename_match = re.search(r'Weekly Sales - (\w+)\s+Wk(\d+)', filename, re.IGNORECASE)
        if filename_match:
            location = filename_match.group(1)
            filename_week = filename_match.group(2)
        else:
            # Pattern 2: Weekly Sales - Location.pdf (without week)
            location_match = re.search(r'Weekly Sales - (\w+)', filename, re.IGNORECASE)
            if location_match:
                location = location_match.group(1)
                
                # Try alternative patterns in filename for week
                week_match = re.search(r'\((\d+)\)', filename)
                if week_match:
                    filename_week = week_match.group(1)
        
        # If still no location, try to extract from PDF content
        if not location:
            location = self._extract_locality_from_pdf_content(pdf_path)
        
        # Try to extract week from PDF content (this is more reliable)
        content_week = self._extract_week_from_pdf_content(pdf_path)
        
        # Prioritize PDF content week over filename week, but use filename as fallback
        if content_week:
            week_num = content_week
            if filename_week and filename_week != content_week:
                print(f"  Note: Filename shows week {filename_week} but PDF content shows week {content_week}. Using PDF content.")
        elif filename_week:
            week_num = filename_week
        
        # Clean up location names to standardize them
        if location:
            location_mapping = {
                "Carter": "Tarxien",    # Carter invoices should be named Tarxien
                "Carters": "Tarxien",  # Carters invoices should be named Tarxien  
                "Fgura": "Fgura",
                "Tarxien": "Tarxien",
                "Zabbar": "Zabbar"
            }
            location = location_mapping.get(location, location)
        
        # If we still don't have a week number, try to infer from common patterns in filename
        if location and not week_num:
            # Try to find week number in the filename using other patterns
            week_patterns = [
                r'wk(\d+)',
                r'week\s*(\d+)',
                r'w(\d+)',
                r'\b(\d{1,2})\b'  # Single or double digit number
            ]
            
            for pattern in week_patterns:
                match = re.search(pattern, filename.lower())
                if match:
                    potential_week = match.group(1)
                    # Only accept reasonable week numbers (1-53)
                    if 1 <= int(potential_week) <= 53:
                        week_num = potential_week
                        break
        
        return location, week_num
    
    def _extract_locality_from_pdf_content(self, pdf_path: str) -> str:
        """Extract locality from PDF content"""
        if not PDF_AVAILABLE:
            return ""
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = pdf.pages[0].extract_text()
                
                # Look for locality pattern in address (e.g., "Tarxien, TXN 9044")
                locality_match = re.search(r'([A-Za-z]+),\s+[A-Z]{2,4}\s+\d+', text)
                if locality_match:
                    return locality_match.group(1)
                
                return ""
        except Exception:
            return ""
    
    def _extract_week_from_pdf_content(self, pdf_path: str) -> str:
        """Extract week number from PDF content"""
        if not PDF_AVAILABLE:
            return ""
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = pdf.pages[0].extract_text()
                
                # Look for week number in invoice description
                week_match = re.search(r'Invoice for week (\d+)', text)
                if week_match:
                    return week_match.group(1)
                
                return ""
        except Exception:
            return ""
    
    def extract_from_pdf(self, pdf_path: str) -> Dict[str, Any]:
        """Extract data from a single PDF file"""
        if not PDF_AVAILABLE:
            print(f"PDF processing not available for {pdf_path}")
            return {}
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = pdf.pages[0].extract_text()
                
                result = {
                    "Document Type": "",
                    "Document Date": "",
                    "Supplier Code": "",
                    "Empty_Column": "",
                    "Document Number": "",
                    "Description": "",
                    "NC": "4001",
                    "VC": "T0",
                    "Locality": "",
                    "Total": ""
                }
                
                # Extract Document Date
                date_match = re.search(r'Document Date\s+(\d{1,2}/\d{1,2}/\d{2,4})', text)
                if date_match:
                    result["Document Date"] = date_match.group(1)
                
                # Extract Document Number
                doc_num_match = re.search(r'Document Number\s+([A-Za-z0-9-]+)', text)
                if doc_num_match:
                    result["Document Number"] = doc_num_match.group(1)
                
                # Extract Locality and set supplier code
                locality_match = re.search(r'([A-Za-z]+),\s+[A-Z]{2,4}\s+\d+', text)
                if locality_match:
                    locality = locality_match.group(1)
                    supplier_code = self.supplier_mapping.get(locality, "")
                    result["Supplier Code"] = supplier_code
                    result["Locality"] = locality
                    
                    # Set NC code based on supplier
                    if supplier_code == "CHAINFGU":
                        result["NC"] = "4002"
                    elif supplier_code == "CHAINTAR":
                        result["NC"] = "4001"
                    elif supplier_code == "CHAINZAB":
                        result["NC"] = "4003"
                
                # Extract Description
                desc_match = re.search(r'Invoice for week (\d+)\s+([^\n]+)', text)
                if desc_match:
                    week_num = desc_match.group(1)
                    period = desc_match.group(2).strip()
                    
                    date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4})\s+(\d{1,2}/\d{1,2}/\d{2,4})', period)
                    if date_match:
                        start_date = date_match.group(1)
                        end_date = date_match.group(2)
                        
                        def format_date(date_str):
                            parts = date_str.split('/')
                            day = parts[0].zfill(2)
                            month = parts[1].zfill(2)
                            year = parts[2]
                            return f"{day}.{month}.{year}"
                        
                        formatted_start = format_date(start_date)
                        formatted_end = format_date(end_date)
                        
                        locality = locality_match.group(1) if locality_match else ""
                        if locality:
                            result["Description"] = f"Chain Supermarket {locality} - Wk{week_num} ({formatted_start}-{formatted_end})"
                        else:
                            result["Description"] = f"Chain Supermarket - Wk{week_num} ({formatted_start}-{formatted_end})"
                
                # Extract Total
                total_match = re.search(r'Invoice for week \d+[^€]*€([0-9,]+\.?\d*)', text)
                if total_match:
                    result["Total"] = total_match.group(1)
                
                return result
                
        except Exception as e:
            print(f"Error processing PDF {pdf_path}: {e}")
            return {}
    
    def _extract_locality_from_filename(self, filename: str) -> str:
        """Extract locality from filename"""
        # For Excel files with 'delicatessen' pattern
        match = re.search(r'delicatessen\s+(\w+)\s+(?:wk)?\d+', filename, re.IGNORECASE)
        if match:
            return match.group(1)
        
        # For PDF files with 'Weekly Sales' pattern
        match = re.search(r'Weekly Sales - (\w+)\s+Week\s+\d+', filename, re.IGNORECASE)
        if match:
            return match.group(1)
        
        return "Unknown"
    
    def _extract_date(self, df: pd.DataFrame) -> str:
        """Extract document date from Excel file"""
        for i in range(min(10, len(df))):
            val = df.iloc[i, 0]
            if pd.notna(val):
                try:
                    if isinstance(val, datetime):
                        return val.strftime('%d/%m/%y')
                    elif isinstance(val, str) and re.match(r'\d{4}-\d{2}-\d{2}', val):
                        date_obj = datetime.strptime(val[:10], '%Y-%m-%d')
                        return date_obj.strftime('%d/%m/%y')
                except:
                    continue
        return ""
    
    def _extract_reference(self, df: pd.DataFrame) -> str:
        """Extract document number/reference from Excel file"""
        for i in range(len(df)):
            for col in range(df.shape[1]):
                val = df.iloc[i, col]
                if pd.notna(val) and isinstance(val, str):
                    match = re.search(r'DEL/\d+/\d{4}', val)
                    if match:
                        return match.group(0)
        return ""
    
    def _extract_description(self, df: pd.DataFrame, locality: str, data_type: str) -> str:
        """Extract description from Excel file"""
        week_number = ""
        for i in range(len(df)):
            val = df.iloc[i, 0]
            if pd.notna(val) and isinstance(val, str):
                if "Commission on Turnover" in val:
                    match = re.search(r'Week (\d+)', val)
                    if match:
                        week_number = match.group(1)
                    break
        
        description_locality = self.description_mapping.get(locality, locality)
        if week_number:
            if data_type == "sales":
                return f"Chain Supermarket {description_locality} - Wk{week_number} (period)"
            else:  # purchases
                return f"Chain Supermarket {description_locality} - Rent for Week {week_number} 2025"
        return ""
    
    def _extract_net(self, df: pd.DataFrame) -> str:
        """Extract net amount from Excel file"""
        for i in range(len(df)):
            val = df.iloc[i, 0]
            if pd.notna(val) and isinstance(val, str) and "Commission on Turnover" in val:
                net_val = df.iloc[i, 6]
                if pd.notna(net_val):
                    return f"{float(net_val):.2f}"
        return ""
    
    def _extract_vat(self, df: pd.DataFrame) -> str:
        """Extract VAT amount from Excel file"""
        for i in range(len(df)):
            val = df.iloc[i, 0]
            if pd.notna(val) and isinstance(val, str) and "VAT @" in val:
                vat_val = df.iloc[i, 6]
                if pd.notna(vat_val):
                    return f"{float(vat_val):.2f}"
        return ""
    
    def _extract_total(self, df: pd.DataFrame) -> str:
        """Extract total amount from Excel file (for sales)"""
        # For sales data, total might be calculated or found elsewhere
        net = self._extract_net(df)
        vat = self._extract_vat(df)
        if net and vat:
            try:
                total = float(net) + float(vat)
                return f"{total:.2f}"
            except:
                pass
        return ""
    
    def _extract_locality(self, df: pd.DataFrame, filename: str) -> str:
        """Extract locality from filename or data"""
        return self._extract_locality_from_filename(filename)
    
    def extract_from_excel(self, filepath: str, data_type: str = "sales") -> Dict[str, Any]:
        """Extract data from a single Excel file"""
        try:
            df = pd.read_excel(filepath, header=None)
            
            filename = os.path.basename(filepath)
            locality = self._extract_locality_from_filename(filename)
            
            template_item = self._get_template_for_locality(locality, data_type)
            extracted_data = {}
            
            for key, value in template_item.items():
                if value.startswith("extract_"):
                    if key == "Document Date":
                        extracted_data[key] = self._extract_date(df)
                    elif key == "Document Number":
                        extracted_data[key] = self._extract_reference(df)
                    elif key == "Description":
                        extracted_data[key] = self._extract_description(df, locality, data_type)
                    elif key == "Net":
                        extracted_data[key] = self._extract_net(df)
                    elif key == "VAT":
                        extracted_data[key] = self._extract_vat(df)
                    elif key == "Total":
                        extracted_data[key] = self._extract_total(df)
                    elif key == "Locality":
                        extracted_data[key] = self._extract_locality(df, filename)
                else:
                    extracted_data[key] = value
            
            return extracted_data
            
        except Exception as e:
            print(f"Error processing Excel file {filepath}: {e}")
            return {}
    
    def process_all_files(self, 
                         sales_input_dir: str = "input", 
                         purchases_input_dir: str = "../CMZ_Purchases/Inputs",
                         pdf_dir: str = "input") -> Dict[str, List[Dict[str, Any]]]:
        """Process all files and return categorized results"""
        results = {"sales": [], "purchases": [], "pdf_sales": []}
        
        # Process PDF files for sales
        # Look for PDFs in the main directory and subdirectories
        pdf_directories = []
        if pdf_dir and os.path.exists(pdf_dir):
            # Add main directory
            pdf_directories.append(pdf_dir)
            
            # Add subdirectories that might contain PDFs
            try:
                for item in os.listdir(pdf_dir):
                    item_path = os.path.join(pdf_dir, item)
                    if os.path.isdir(item_path):
                        # Check if subdirectory contains PDF files
                        pdf_files_in_subdir = [f for f in os.listdir(item_path) if f.endswith('.pdf')]
                        if pdf_files_in_subdir:
                            pdf_directories.append(item_path)
            except Exception as e:
                print(f"Error scanning subdirectories: {e}")
        
        if PDF_AVAILABLE and pdf_directories:
            for search_dir in pdf_directories:
                try:
                    pdf_files = [f for f in os.listdir(search_dir) if f.endswith('.pdf')]
                    if pdf_files:
                        print(f"Processing PDF sales files from: {search_dir}")
                        for filename in pdf_files:
                            filepath = os.path.join(search_dir, filename)
                            print(f"  Processing PDF: {filename}")
                            data = self.extract_from_pdf(filepath)
                            if data:
                                results["pdf_sales"].append(data)
                except Exception as e:
                    print(f"Error processing directory {search_dir}: {e}")
        elif not PDF_AVAILABLE:
            print("PDF processing disabled (pdfplumber not available)")
        else:
            print(f"No PDF directories found or no PDFs available")
        
        # Process Excel files for purchases only (no sales Excel files)
        if purchases_input_dir and os.path.exists(purchases_input_dir):
            print(f"\nProcessing Excel purchases files from: {purchases_input_dir}")
            excel_files = [f for f in os.listdir(purchases_input_dir) 
                          if f.endswith(('.xlsx', '.xls')) and 'delicatessen' in f.lower()]
            
            for filename in excel_files:
                filepath = os.path.join(purchases_input_dir, filename)
                print(f"  Processing Excel: {filename}")
                data = self.extract_from_excel(filepath, "purchases")
                if data:
                    results["purchases"].append(data)
        
        return results
    
    def save_results(self, results: Dict[str, List[Dict[str, Any]]], output_dir: str = "output"):
        """Save results to separate files by category and individual JSON files for each record"""
        os.makedirs(output_dir, exist_ok=True)
        
        for category, data in results.items():
            if not data:
                continue
            
            # Create single directory for all individual JSON files
            individual_json_dir = os.path.join(output_dir, "individual_json")
            os.makedirs(individual_json_dir, exist_ok=True)
            
            # Save individual JSON files for each record
            individual_count = 0
            for i, record in enumerate(data):
                try:
                    # Create filename based on document info
                    doc_num = record.get('Document Number', f'record_{i+1}')
                    supplier = record.get('Supplier Code', 'unknown')
                    date = record.get('Document Date', '').replace('/', '-')
                    
                    # Clean filename (remove invalid characters)
                    safe_doc_num = re.sub(r'[<>:"/\\|?*]', '_', str(doc_num))
                    safe_supplier = re.sub(r'[<>:"/\\|?*]', '_', str(supplier))
                    safe_date = re.sub(r'[<>:"/\\|?*]', '_', str(date))
                    
                    # Create descriptive filename with category prefix
                    if safe_supplier and safe_supplier != 'unknown':
                        filename = f"{category}_{safe_supplier}_{safe_doc_num}_{safe_date}.json"
                    else:
                        filename = f"{category}_{safe_doc_num}_{safe_date}.json"
                    
                    # Ensure filename isn't too long
                    if len(filename) > 100:
                        filename = f"{category}_record_{i+1}.json"
                    
                    individual_json_file = os.path.join(individual_json_dir, filename)
                    
                    with open(individual_json_file, 'w', encoding='utf-8') as f:
                        json.dump([record], f, indent=2, ensure_ascii=False)
                    individual_count += 1
                    
                except Exception as e:
                    print(f"Error saving individual JSON for {category} record {i+1}: {e}")
                    # Fallback filename
                    fallback_filename = f"{category}_record_{i+1}.json"
                    fallback_path = os.path.join(individual_json_dir, fallback_filename)
                    try:
                        with open(fallback_path, 'w', encoding='utf-8') as f:
                            json.dump([record], f, indent=2, ensure_ascii=False)
                        individual_count += 1
                    except Exception as e2:
                        print(f"Error saving fallback JSON for {category} record {i+1}: {e2}")
            
            if individual_count > 0:
                print(f"Individual JSON files saved to {individual_json_dir} ({individual_count} files)")
            
            # Save combined JSON (keep existing functionality)
            json_file = os.path.join(output_dir, f"{category}_data.json")
            try:
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                print(f"Combined JSON results saved to {json_file}")
            except Exception as e:
                print(f"Error saving combined JSON for {category}: {e}")
            
            # Save CSV
            csv_file = os.path.join(output_dir, f"{category}_data.csv")
            try:
                headers = list(data[0].keys())
                with open(csv_file, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=headers)
                    writer.writeheader()
                    writer.writerows(data)
                print(f"CSV results saved to {csv_file}")
            except Exception as e:
                print(f"Error saving CSV for {category}: {e}")
    
    def generate_combined_sales_purchases_csv(self):
        """Generate combined sales and purchases CSV using external module"""
        print("\n4. Generating combined sales and purchases CSV...")
        
        # Import the dedicated module for this functionality
        try:
            from generate_combined_csv import generate_combined_csv, generate_upload_csv
            
            # Generate combined CSV
            combined_success = generate_combined_csv()
            
            # Generate upload CSV
            print("\n5. Generating upload CSV...")
            upload_success = generate_upload_csv()
            
            return combined_success and upload_success
        except ImportError:
            print("Error: generate_combined_csv module not found")
            return False
    
    def print_results(self, results: Dict[str, List[Dict[str, Any]]]):
        """Print extracted results in a formatted way"""
        for category, data in results.items():
            if not data:
                continue
            
            print(f"\n{'='*60}")
            print(f"{category.upper()} DATA SUMMARY")
            print(f"{'='*60}")
            
            for i, record in enumerate(data, 1):
                print(f"\nRecord {i}:")
                for key, value in record.items():
                    print(f"  {key}: {value}")


def main():
    """Main function to run the unified data extraction"""
    print("CMZ Unified Data Extractor")
    print("="*50)
    
    try:
        # Initialize extractor
        extractor = CombinedDataExtractor()
        
        # Rename PDF files to standard format
        print("\n1. Renaming PDF files to standard format...")
        extractor.rename_pdf_files()
        
        # Process all files
        print("\n2. Processing all data files...")
        results = extractor.process_all_files()
        
        # Calculate total records
        total_records = sum(len(data) for data in results.values())
        
        if total_records > 0:
            # Print results
            extractor.print_results(results)
            
            # Save results
            print("\n3. Saving results...")
            extractor.save_results(results)
            
            # Generate combined sales and purchases CSV
            extractor.generate_combined_sales_purchases_csv()
            
            # Copy renamed PDF files to output
            print("\n6. Copying renamed PDF files to output...")
            extractor.copy_renamed_pdfs_to_output()
            
            # Create renamed invoices folder
            print("\n7. Creating renamed invoices folder...")
            extractor.create_renamed_invoices_folder()
            
            print(f"\nSuccessfully processed {total_records} total records")
            for category, data in results.items():
                if data:
                    print(f"  {category}: {len(data)} records")
        else:
            print("\nNo data extracted from any files")
            
    except Exception as e:
        print(f"\nError during execution: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
