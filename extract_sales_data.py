#!/usr/bin/env python3
"""
Simple PDF Sales Data Extractor

This script extracts sales data from CMZ weekly sales PDFs and outputs them in JSON format.
"""

import pdfplumber
import json
import re
from pathlib import Path
from typing import Dict, List, Any


def extract_sales_data_from_pdf(pdf_path: str) -> Dict[str, Any]:
    """
    Extract sales data from a single PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        Dictionary containing extracted data
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Extract text from the first page
            text = pdf.pages[0].extract_text()
            
            # Initialize result dictionary
            result = {
                "Document Date": "",
                "Document Number": "",
                "Description": "",
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
            
            # Extract Locality from the address first (needed for description)
            # Look for the locality pattern (e.g., "Tarxien, TXN 9044" or "Fgura, FGR 1242")
            locality_match = re.search(r'([A-Za-z]+),\s+[A-Z]{2,4}\s+\d+', text)
            if locality_match:
                result["Locality"] = locality_match.group(1)
            
            # Extract Description (Invoice description)
            desc_match = re.search(r'Invoice for week (\d+)\s+([^\n]+)', text)
            if desc_match:
                week_num = desc_match.group(1)
                period = desc_match.group(2).strip()
                
                # Extract dates from the period (e.g., "June 8/6/25 14/6/25")
                date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4})\s+(\d{1,2}/\d{1,2}/\d{2,4})', period)
                if date_match:
                    start_date = date_match.group(1)
                    end_date = date_match.group(2)
                    
                    # Convert date format from d/m/yy to dd.mm.yy
                    def format_date(date_str):
                        parts = date_str.split('/')
                        day = parts[0].zfill(2)
                        month = parts[1].zfill(2)
                        year = parts[2]
                        return f"{day}.{month}.{year}"
                    
                    formatted_start = format_date(start_date)
                    formatted_end = format_date(end_date)
                    
                    # Get locality for the description
                    locality = result.get("Locality", "")
                    if locality:
                        result["Description"] = f"Chain Supermarket {locality} - Wk{week_num} ({formatted_start}-{formatted_end})"
                    else:
                        result["Description"] = f"Chain Supermarket - Wk{week_num} ({formatted_start}-{formatted_end})"
                else:
                    result["Description"] = f"Week {week_num} - {period}"
            
            # Extract Total amount - look for the amount after "Invoice for week"
            total_match = re.search(r'Invoice for week \d+[^€]*€([0-9,]+\.?\d*)', text)
            if total_match:
                result["Total"] = f"€{total_match.group(1)}"
            
            return result
            
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return {
            "Document Date": "",
            "Document Number": "",
            "Description": "",
            "Locality": "",
            "Total": ""
        }


def process_weekly_sales_pdfs(input_dir: str = "input/WeeklySales") -> List[Dict[str, Any]]:
    """
    Process all PDF files in the WeeklySales directory.
    
    Args:
        input_dir: Directory containing the PDF files
        
    Returns:
        List of dictionaries containing extracted data
    """
    results = []
    pdf_dir = Path(input_dir)
    
    if not pdf_dir.exists():
        print(f"Directory {input_dir} does not exist")
        return results
    
    # Find all PDF files
    pdf_files = list(pdf_dir.glob("*.pdf"))
    
    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return results
    
    print(f"Found {len(pdf_files)} PDF files to process")
    
    for pdf_file in pdf_files:
        print(f"Processing: {pdf_file.name}")
        data = extract_sales_data_from_pdf(str(pdf_file))
        results.append(data)
    
    return results


def save_to_json(data: List[Dict[str, Any]], output_file: str = "output/weekly_sales_data.json"):
    """
    Save extracted data to JSON file.
    
    Args:
        data: List of dictionaries containing extracted data
        output_file: Output file path
    """
    output_path = Path(output_file)
    output_path.parent.mkdir(exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print(f"Data saved to: {output_path}")


def main():
    """Main function to process PDFs and extract sales data."""
    print("CMZ Weekly Sales Data Extractor")
    print("=" * 40)
    
    # Process all PDFs in the WeeklySales directory
    sales_data = process_weekly_sales_pdfs()
    
    if sales_data:
        # Save to JSON file
        save_to_json(sales_data)
        
        # Display results
        print("\nExtracted Data:")
        print("-" * 40)
        for i, data in enumerate(sales_data, 1):
            print(f"\nRecord {i}:")
            for key, value in data.items():
                print(f"  {key}: {value}")
    else:
        print("No data extracted")


if __name__ == "__main__":
    main()
