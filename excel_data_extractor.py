"""
Excel Data Extractor for CMZ Purchases

This script extracts data from Excel invoice files according to the JSON template structure.
It processes delicatessen invoice files and outputs structured data.
"""

import pandas as pd
import json
import csv
import os
import re
from datetime import datetime
from typing import Dict, Any, List


class ExcelDataExtractor:
    """Extracts data from Excel invoice files according to JSON template"""
    
    def __init__(self, json_template_path: str = "templates/JSON_Template_Sales.json"):
        """Initialize the extractor with JSON template"""
        self.json_template_path = json_template_path
        self.template = self._load_template()
        self.supplier_mapping = self._create_supplier_mapping()
    
    def _load_template(self) -> List[Dict[str, str]]:
        """Load the JSON template structure"""
        with open(self.json_template_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _create_supplier_mapping(self) -> Dict[str, str]:
        """Create a mapping from locality to supplier code"""
        # Hardcoded mapping based on the template structure
        return {
            "Fgura": "CHAINFGU",
            "Tarxien": "CHAINTAR", 
            "Zabbar": "CHAINZAB"
        }
    
    def _get_template_for_locality(self, locality: str) -> Dict[str, str]:
        """Get the template item for a specific locality"""
        if isinstance(self.template, list):
            # Get supplier code for this locality
            supplier_code = self.supplier_mapping.get(locality, "")
            
            # Find template item with matching supplier code
            for item in self.template:
                if "Supplier Code" in item and item["Supplier Code"] == supplier_code:
                    return item
            
            # If no match found, return first template item
            if self.template:
                return self.template[0]
        
        # Fallback to empty dict
        return {}
    
    def _extract_locality_from_filename(self, filename: str) -> str:
        """Extract locality from filename (e.g., 'Fgura', 'Tarxien', 'Zabbar')"""
        # Extract locality from filename pattern: "delicatessen [Locality] wk[number].xlsx"
        match = re.search(r'delicatessen\s+(\w+)\s+wk\d+', filename, re.IGNORECASE)
        return match.group(1) if match else "Unknown"
    
    def _extract_date(self, df: pd.DataFrame) -> str:
        """Extract document date from the Excel file"""
        # Look for date in column 0, typically around row 6
        for i in range(min(10, len(df))):
            val = df.iloc[i, 0]
            if pd.notna(val):
                try:
                    # Try to parse as datetime
                    if isinstance(val, datetime):
                        return val.strftime('%d/%m/%y')
                    elif isinstance(val, str) and re.match(r'\d{4}-\d{2}-\d{2}', val):
                        # Parse the date and reformat to dd/mm/yy
                        date_obj = datetime.strptime(val[:10], '%Y-%m-%d')
                        return date_obj.strftime('%d/%m/%y')
                except:
                    continue
        return ""
    
    def _extract_reference(self, df: pd.DataFrame) -> str:
        """Extract document number/reference from the Excel file"""
        # Look for invoice number pattern like "DEL/22/2025"
        for i in range(len(df)):
            for col in range(df.shape[1]):
                val = df.iloc[i, col]
                if pd.notna(val) and isinstance(val, str):
                    # Look for pattern like "DEL/XX/YYYY"
                    match = re.search(r'DEL/\d+/\d{4}', val)
                    if match:
                        return match.group(0)
        return ""
    
    def _extract_description(self, df: pd.DataFrame, locality: str) -> str:
        """Extract description from the Excel file"""
        # Look for commission description around row 25 to extract week number
        week_number = ""
        for i in range(len(df)):
            val = df.iloc[i, 0]
            if pd.notna(val) and isinstance(val, str):
                if "Commission on Turnover" in val:
                    # Extract week number from the description
                    match = re.search(r'Week (\d+)', val)
                    if match:
                        week_number = match.group(1)
                    break
        
        # Format description as requested: "Chain Supermarket [Locality] - Rent for Week [Number] 2025"
        if week_number:
            return f"Chain Supermarket {locality} - Rent for Week {week_number} 2025"
        return ""
    
    def _extract_net(self, df: pd.DataFrame) -> str:
        """Extract net amount from the Excel file"""
        # Look for commission amount (net) around row 25
        for i in range(len(df)):
            val = df.iloc[i, 0]
            if pd.notna(val) and isinstance(val, str) and "Commission on Turnover" in val:
                # Get the net amount from column 6
                net_val = df.iloc[i, 6]
                if pd.notna(net_val):
                    return f"{float(net_val):.2f}"
        return ""
    
    def _extract_vat(self, df: pd.DataFrame) -> str:
        """Extract VAT amount from the Excel file"""
        # Look for VAT amount around row 26
        for i in range(len(df)):
            val = df.iloc[i, 0]
            if pd.notna(val) and isinstance(val, str) and "VAT @" in val:
                # Get the VAT amount from column 6
                vat_val = df.iloc[i, 6]
                if pd.notna(vat_val):
                    return f"{float(vat_val):.2f}"
        return ""
    
    def extract_from_file(self, filepath: str) -> Dict[str, Any]:
        """Extract data from a single Excel file"""
        try:
            df = pd.read_excel(filepath, header=None)
            
            # Extract data according to template
            extracted_data = {}
            
            # Locality (from filename) - needed for supplier code lookup and description formatting
            filename = os.path.basename(filepath)
            locality = self._extract_locality_from_filename(filename)
            
            # Get template for this locality to include constant values
            template_item = self._get_template_for_locality(locality)
            
            # Add all fields from template (including constants)
            for key, value in template_item.items():
                if value.startswith("extract_"):
                    # Extract dynamic values
                    if key == "Document Date":
                        extracted_data[key] = self._extract_date(df)
                    elif key == "Document Number":
                        extracted_data[key] = self._extract_reference(df)
                    elif key == "Description":
                        extracted_data[key] = self._extract_description(df, locality)
                    elif key == "Net":
                        extracted_data[key] = self._extract_net(df)
                    elif key == "VAT":
                        extracted_data[key] = self._extract_vat(df)
                else:
                    # Use constant values directly (Type, NC, VC, Supplier Code)
                    extracted_data[key] = value
            
            return extracted_data
            
        except Exception as e:
            print(f"Error processing {filepath}: {e}")
            return {}
    
    def extract_all_files(self, input_directory: str = "Inputs") -> List[Dict[str, Any]]:
        """Extract data from all Excel files in the input directory"""
        results = []
        
        if not os.path.exists(input_directory):
            print(f"Input directory {input_directory} not found")
            return results
        
        # Find all Excel files
        excel_files = [f for f in os.listdir(input_directory) 
                      if f.endswith(('.xlsx', '.xls')) and 'delicatessen' in f.lower()]
        
        print(f"Found {len(excel_files)} Excel files to process:")
        for file in excel_files:
            print(f"  - {file}")
        
        # Process each file
        for filename in excel_files:
            filepath = os.path.join(input_directory, filename)
            print(f"\nProcessing: {filename}")
            
            data = self.extract_from_file(filepath)
            if data:
                results.append(data)
                print(f"  Extracted: {data}")
            else:
                print(f"  Failed to extract data from {filename}")
        
        return results
    
    def save_results(self, results: List[Dict[str, Any]], output_file: str = "extracted_rent.json"):
        """Save extracted results to JSON file"""
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            print(f"\nResults saved to {output_file}")
        except Exception as e:
            print(f"Error saving results: {e}")
    
    def save_csv_results(self, results: List[Dict[str, Any]], output_file: str = "extracted_rent.csv"):
        """Save extracted results to CSV file with headers from template"""
        try:
            if not results:
                print("No data to save to CSV")
                return
            
            # Get headers from the first template item (all templates have same structure)
            template_item = self.template[0] if self.template else {}
            headers = list(template_item.keys())
            
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=headers)
                writer.writeheader()
                writer.writerows(results)
            
            print(f"\nCSV results saved to {output_file}")
        except Exception as e:
            print(f"Error saving CSV results: {e}")
    
    def print_results(self, results: List[Dict[str, Any]]):
        """Print extracted results in a formatted way"""
        print("\n" + "="*60)
        print("EXTRACTED DATA SUMMARY")
        print("="*60)
        
        for i, data in enumerate(results, 1):
            print(f"\nRecord {i}:")
            for key, value in data.items():
                print(f"  {key}: {value}")


def main():
    """Main function to run the data extraction"""
    print("CMZ Purchases - Excel Data Extractor")
    print("="*40)
    
    # Initialize extractor
    extractor = ExcelDataExtractor()
    
    # Extract data from all files
    results = extractor.extract_all_files()
    
    if results:
        # Print results
        extractor.print_results(results)
        
        # Save results to JSON
        extractor.save_results(results)
        
        # Save results to CSV
        extractor.save_csv_results(results)
        
        print(f"\nSuccessfully processed {len(results)} files")
    else:
        print("No data extracted from any files")


if __name__ == "__main__":
    main()
