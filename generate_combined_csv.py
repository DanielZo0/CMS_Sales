#!/usr/bin/env python3
"""
Generate Combined Sales and Purchases CSV

This script combines sales and purchases data into a single CSV file with standardized fields.
Uses only built-in Python modules - no external dependencies required.
"""

import os
import csv
import json
from typing import Dict, List, Any
from datetime import datetime


def load_and_standardize_sales_data(file_path: str) -> List[Dict[str, Any]]:
    """Load sales data and standardize fields"""
    if not os.path.exists(file_path):
        print(f"Sales file not found: {file_path}")
        return []
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            records = []
            
            for row in reader:
                # For sales data: VAT is zero, Net equals Total
                total_str = str(row.get('Total', '')).replace(',', '')
                try:
                    total = float(total_str)
                    net = total  # Net equals Total (no VAT)
                    vat = 0      # VAT is zero
                except (ValueError, TypeError):
                    total = 0
                    net = 0
                    vat = 0
                
                record = {
                    'Data Type': 'Sales',
                    'Document Type': row.get('Document Type', ''),
                    'Document Date': row.get('Document Date', ''),
                    'Supplier Code': row.get('Supplier Code', ''),
                    'Empty_Column': row.get('Empty_Column', row.get('Empty Column', '')),  # Handle both column names
                    'Document Number': row.get('Document Number', ''),
                    'Description': row.get('Description', ''),
                    'NC': row.get('NC', ''),
                    'VC': row.get('VC', ''),
                    'Locality': row.get('Locality', ''),
                    'Net': f"{net:.2f}" if net > 0 else '',
                    'VAT': f"{vat:.2f}",  # Always show VAT (will be 0.00 for sales)
                    'Total': f"{total:.2f}" if total > 0 else ''
                }
                records.append(record)
            
            print(f"Loaded {len(records)} sales records")
            return records
        
    except Exception as e:
        print(f"Error loading sales data: {e}")
        return []


def load_and_standardize_purchases_data(file_path: str) -> List[Dict[str, Any]]:
    """Load purchases data and standardize fields"""
    if not os.path.exists(file_path):
        print(f"Purchases file not found: {file_path}")
        return []
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            records = []
            
            for row in reader:
                # Calculate Total from Net + VAT
                try:
                    net = float(row.get('Net', 0))
                    vat = float(row.get('VAT', 0))
                    total = net + vat
                except (ValueError, TypeError):
                    net = 0
                    vat = 0
                    total = 0
                
                record = {
                    'Data Type': 'Expenses',
                    'Document Type': row.get('Document Type', ''),
                    'Document Date': row.get('Document Date', ''),
                    'Supplier Code': row.get('Supplier Code', ''),
                    'Empty_Column': row.get('Empty_Column', row.get('Empty Column', '')),  # Handle both column names
                    'Document Number': row.get('Document Number', ''),
                    'Description': row.get('Description', ''),
                    'NC': row.get('NC', ''),
                    'VC': row.get('VC', ''),
                    'Locality': row.get('Locality', ''),
                    'Net': f"{net:.2f}" if net > 0 else '',
                    'VAT': f"{vat:.2f}" if vat > 0 else '',
                    'Total': f"{total:.2f}" if total > 0 else ''
                }
                records.append(record)
            
            print(f"Loaded {len(records)} purchase records")
            return records
        
    except Exception as e:
        print(f"Error loading purchases data: {e}")
        return []


def sort_records_by_date(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Sort records by document date"""
    def parse_date(date_str):
        """Parse date string to sortable format"""
        try:
            # Handle different date formats
            if '/' in date_str:
                parts = date_str.split('/')
                if len(parts) == 3:
                    day, month, year = parts
                    # Convert 2-digit year to 4-digit
                    if len(year) == 2:
                        year = f"20{year}"
                    return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            return date_str
        except:
            return "9999-12-31"  # Put invalid dates at the end
    
    return sorted(records, key=lambda x: (parse_date(x.get('Document Date', '')), x.get('Data Type', '')))


def format_date_to_ddmmyyyy(date_str: str) -> str:
    """Convert date string to dd/mm/yyyy format"""
    if not date_str or date_str.strip() == '':
        return ''
    
    try:
        # Try different input formats
        date_formats = [
            "%d/%m/%y",     # 31/5/25
            "%d/%m/%Y",     # 31/05/2025
            "%d-%m-%y",     # 31-5-25
            "%d-%m-%Y",     # 31-05-2025
            "%Y-%m-%d",     # 2025-05-31
            "%m/%d/%y",     # 5/31/25
            "%m/%d/%Y",     # 05/31/2025
        ]
        
        date_obj = None
        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(date_str.strip(), fmt)
                break
            except ValueError:
                continue
        
        if date_obj:
            # Convert to dd/mm/yyyy format
            return date_obj.strftime("%d/%m/%Y")
        else:
            # If parsing fails, return original
            return date_str
            
    except Exception:
        return date_str


def load_upload_template(template_path: str = "templates/JSON_Template_Upload.json") -> Dict[str, str]:
    """Load the upload template structure"""
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            template_list = json.load(f)
            return template_list[0] if template_list else {}
    except FileNotFoundError:
        print(f"Upload template file {template_path} not found. Using default structure.")
        return {
            "Type": "",
            "Account Reference": "supplier_code",
            "Nominal A/C Ref": "NC_extract",
            "Department Code": "",
            "Date": "extract_date",
            "Refernce": "extract_reference",
            "Details": "extract_description",
            "Net Amount": "extract_net",
            "Tax Code": "T0",
            "Tax Amount": "extract_vat"
        }


def generate_upload_csv(sales_file: str = "output/pdf_sales_data.csv", 
                       purchases_file: str = "output/purchases_data.csv",
                       upload_template_path: str = "templates/JSON_Template_Upload.json") -> bool:
    """Generate a separate upload CSV file based on the upload template format"""
    
    try:
        # Load upload template
        upload_template = load_upload_template(upload_template_path)
        if not upload_template:
            print("Error: Could not load upload template")
            return False
        
        print("Generating upload format CSV...")
        upload_records = []
        
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
            upload_file = "output/upload_data.csv"
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


def convert_to_upload_format(records: List[Dict[str, Any]], template: Dict[str, str]) -> List[Dict[str, str]]:
    """Convert combined records to upload format based on template (deprecated - use generate_upload_csv instead)"""
    upload_records = []
    
    for record in records:
        upload_record = {}
        
        for field, mapping in template.items():
            if mapping == "supplier_code":
                upload_record[field] = record.get('Supplier Code', '')
            elif mapping == "NC_extract":
                upload_record[field] = record.get('NC', '')
            elif mapping == "extract_date":
                upload_record[field] = record.get('Document Date', '')
            elif mapping == "extract_reference":
                upload_record[field] = record.get('Document Number', '')
            elif mapping == "extract_description":
                upload_record[field] = record.get('Description', '')
            elif mapping == "extract_net":
                upload_record[field] = record.get('Net', '')
            elif mapping == "extract_vat":
                upload_record[field] = record.get('VAT', '')
            else:
                # Static value or empty
                upload_record[field] = mapping
        
        upload_records.append(upload_record)
    
    return upload_records


def generate_combined_csv():
    """Generate combined sales and purchases CSV"""
    print("CMZ Combined Sales and Purchases CSV Generator")
    print("=" * 55)
    
    # File paths
    sales_file = "output/pdf_sales_data.csv"
    purchases_file = "output/purchases_data.csv"
    output_file = "output/combined_sales_purchases.csv"
    
    # Load data
    sales_records = load_and_standardize_sales_data(sales_file)
    purchases_records = load_and_standardize_purchases_data(purchases_file)
    
    # Combine records
    all_records = sales_records + purchases_records
    
    if not all_records:
        print("No data to combine")
        return
    
    # Sort by date
    all_records = sort_records_by_date(all_records)
    
    # Define field order
    fieldnames = [
        'Data Type',
        'Document Type',
        'Document Date',
        'Supplier Code',
        'Empty_Column',
        'Document Number',
        'Description',
        'NC',
        'VC',
        'Locality',
        'Net',
        'VAT',
        'Total'
    ]
    
    # Save combined CSV
    try:
        os.makedirs("output", exist_ok=True)
        
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(all_records)
        
        print(f"\nCombined CSV saved to: {output_file}")
        print(f"Total records: {len(all_records)}")
        
        # Print summary
        sales_count = len([r for r in all_records if r['Data Type'] == 'Sales'])
        expenses_count = len([r for r in all_records if r['Data Type'] == 'Expenses'])
        
        print(f"  Sales records: {sales_count}")
        print(f"  Expense records: {expenses_count}")
        
        # Show sample records
        print(f"\nSample records (first 3):")
        print("-" * 80)
        for i, record in enumerate(all_records[:3]):
            print(f"Record {i+1}: {record['Data Type']} - {record['Description'][:50]}...")
            print(f"  Date: {record['Document Date']}, Total: {record['Total']}")
        
        return True
        
    except Exception as e:
        print(f"Error saving combined CSV: {e}")
        return False


def main():
    """Main function to generate both combined and upload CSV files"""
    print("CMZ Data Processing - Generating CSV Files")
    print("=" * 60)
    
    # Generate combined CSV
    print("\n1. Generating Combined CSV...")
    combined_success = generate_combined_csv()
    
    # Generate separate upload CSV
    print("\n2. Generating Upload CSV...")
    upload_success = generate_upload_csv()
    
    # Summary
    print("\n" + "=" * 60)
    print("PROCESSING SUMMARY:")
    print(f"Combined CSV: {'✓ Success' if combined_success else '✗ Failed'}")
    print(f"Upload CSV:   {'✓ Success' if upload_success else '✗ Failed'}")
    
    if combined_success and upload_success:
        print("\n✓ All CSV files generated successfully!")
        return True
    else:
        print("\n✗ Some CSV files failed to generate!")
        return False


if __name__ == "__main__":
    success = main()
    if not success:
        exit(1)
