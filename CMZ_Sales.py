#!/usr/bin/env python3
"""
CMZ Sales PDF Data Extraction Script

This script extracts sales data from PDF files and processes it for analysis.
Supports multiple PDF processing libraries for different types of PDFs.
"""

import os
import sys
import pandas as pd
from pathlib import Path
import logging
from typing import List, Dict, Any, Optional
import argparse
from datetime import datetime

# PDF processing libraries
try:
    import PyPDF2
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False
    print("Warning: PyPDF2 not available. Install with: pip install PyPDF2")

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False
    print("Warning: pdfplumber not available. Install with: pip install pdfplumber")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('cmz_sales.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class CMZSalesExtractor:
    """Main class for extracting sales data from PDF files."""
    
    def __init__(self, input_dir: str = "input", output_dir: str = "output"):
        """
        Initialize the CMZ Sales Extractor.
        
        Args:
            input_dir: Directory containing PDF files to process
            output_dir: Directory to save extracted data
        """
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.extracted_data = []
        
        # Create directories if they don't exist
        self.input_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
        
        logger.info(f"Initialized CMZ Sales Extractor")
        logger.info(f"Input directory: {self.input_dir}")
        logger.info(f"Output directory: {self.output_dir}")
    
    def extract_text_pypdf2(self, pdf_path: Path) -> str:
        """Extract text using PyPDF2 library."""
        if not PYPDF2_AVAILABLE:
            raise ImportError("PyPDF2 is not available")
        
        text = ""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
        except Exception as e:
            logger.error(f"Error extracting text with PyPDF2 from {pdf_path}: {e}")
            raise
        
        return text
    
    def extract_text_pdfplumber(self, pdf_path: Path) -> str:
        """Extract text using pdfplumber library."""
        if not PDFPLUMBER_AVAILABLE:
            raise ImportError("pdfplumber is not available")
        
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
        except Exception as e:
            logger.error(f"Error extracting text with pdfplumber from {pdf_path}: {e}")
            raise
        
        return text
    
    def extract_tables_pdfplumber(self, pdf_path: Path) -> List[pd.DataFrame]:
        """Extract tables using pdfplumber library."""
        if not PDFPLUMBER_AVAILABLE:
            raise ImportError("pdfplumber is not available")
        
        tables = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    page_tables = page.extract_tables()
                    for table_num, table in enumerate(page_tables):
                        if table:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            df['source_page'] = page_num + 1
                            df['table_number'] = table_num + 1
                            tables.append(df)
        except Exception as e:
            logger.error(f"Error extracting tables with pdfplumber from {pdf_path}: {e}")
            raise
        
        return tables
    
    def process_pdf(self, pdf_path: Path) -> Dict[str, Any]:
        """
        Process a single PDF file and extract relevant data.
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            Dictionary containing extracted data
        """
        logger.info(f"Processing PDF: {pdf_path.name}")
        
        result = {
            'filename': pdf_path.name,
            'file_path': str(pdf_path),
            'processed_at': datetime.now().isoformat(),
            'text_content': '',
            'tables': [],
            'extraction_method': '',
            'error': None
        }
        
        try:
            # Try pdfplumber first (better for tables and complex layouts)
            if PDFPLUMBER_AVAILABLE:
                result['text_content'] = self.extract_text_pdfplumber(pdf_path)
                result['tables'] = self.extract_tables_pdfplumber(pdf_path)
                result['extraction_method'] = 'pdfplumber'
                logger.info(f"Successfully extracted data using pdfplumber from {pdf_path.name}")
            
            # Fallback to PyPDF2 if pdfplumber fails or is not available
            elif PYPDF2_AVAILABLE:
                result['text_content'] = self.extract_text_pypdf2(pdf_path)
                result['extraction_method'] = 'PyPDF2'
                logger.info(f"Successfully extracted data using PyPDF2 from {pdf_path.name}")
            
            else:
                raise ImportError("No PDF processing libraries available")
                
        except Exception as e:
            result['error'] = str(e)
            logger.error(f"Failed to process {pdf_path.name}: {e}")
        
        return result
    
    def process_all_pdfs(self) -> List[Dict[str, Any]]:
        """
        Process all PDF files in the input directory.
        
        Returns:
            List of dictionaries containing extracted data from all PDFs
        """
        pdf_files = list(self.input_dir.glob("*.pdf"))
        
        if not pdf_files:
            logger.warning(f"No PDF files found in {self.input_dir}")
            return []
        
        logger.info(f"Found {len(pdf_files)} PDF files to process")
        
        results = []
        for pdf_file in pdf_files:
            result = self.process_pdf(pdf_file)
            results.append(result)
            self.extracted_data.append(result)
        
        return results
    
    def save_results(self, results: List[Dict[str, Any]], format: str = 'excel') -> None:
        """
        Save extracted results to file.
        
        Args:
            results: List of extracted data dictionaries
            format: Output format ('excel', 'csv', 'json')
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if format.lower() == 'excel':
            output_file = self.output_dir / f"cmz_sales_extracted_{timestamp}.xlsx"
            self._save_to_excel(results, output_file)
        elif format.lower() == 'csv':
            output_file = self.output_dir / f"cmz_sales_extracted_{timestamp}.csv"
            self._save_to_csv(results, output_file)
        elif format.lower() == 'json':
            output_file = self.output_dir / f"cmz_sales_extracted_{timestamp}.json"
            self._save_to_json(results, output_file)
        else:
            raise ValueError(f"Unsupported format: {format}")
        
        logger.info(f"Results saved to: {output_file}")
    
    def _save_to_excel(self, results: List[Dict[str, Any]], output_file: Path) -> None:
        """Save results to Excel file with multiple sheets."""
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Summary sheet
            summary_data = []
            for result in results:
                summary_data.append({
                    'filename': result['filename'],
                    'processed_at': result['processed_at'],
                    'extraction_method': result['extraction_method'],
                    'has_error': result['error'] is not None,
                    'error_message': result['error'] or '',
                    'table_count': len(result['tables'])
                })
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Text content sheet
            text_data = []
            for result in results:
                if result['text_content']:
                    text_data.append({
                        'filename': result['filename'],
                        'text_content': result['text_content']
                    })
            
            if text_data:
                text_df = pd.DataFrame(text_data)
                text_df.to_excel(writer, sheet_name='Text_Content', index=False)
            
            # Tables sheet
            all_tables = []
            for result in results:
                for table in result['tables']:
                    table['source_file'] = result['filename']
                    all_tables.append(table)
            
            if all_tables:
                combined_tables = pd.concat(all_tables, ignore_index=True)
                combined_tables.to_excel(writer, sheet_name='Tables', index=False)
    
    def _save_to_csv(self, results: List[Dict[str, Any]], output_file: Path) -> None:
        """Save results to CSV file."""
        # Flatten results for CSV format
        flattened_data = []
        for result in results:
            flattened_data.append({
                'filename': result['filename'],
                'processed_at': result['processed_at'],
                'extraction_method': result['extraction_method'],
                'has_error': result['error'] is not None,
                'error_message': result['error'] or '',
                'text_content': result['text_content'],
                'table_count': len(result['tables'])
            })
        
        df = pd.DataFrame(flattened_data)
        df.to_csv(output_file, index=False)
    
    def _save_to_json(self, results: List[Dict[str, Any]], output_file: Path) -> None:
        """Save results to JSON file."""
        import json
        
        # Convert DataFrames to dictionaries for JSON serialization
        json_results = []
        for result in results:
            json_result = result.copy()
            json_result['tables'] = [table.to_dict('records') for table in result['tables']]
            json_results.append(json_result)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(json_results, f, indent=2, ensure_ascii=False)


def main():
    """Main function to run the CMZ Sales PDF extraction."""
    parser = argparse.ArgumentParser(description='Extract data from CMZ Sales PDF files')
    parser.add_argument('--input', '-i', default='input', 
                       help='Input directory containing PDF files (default: input)')
    parser.add_argument('--output', '-o', default='output',
                       help='Output directory for extracted data (default: output)')
    parser.add_argument('--format', '-f', choices=['excel', 'csv', 'json'], 
                       default='excel', help='Output format (default: excel)')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        # Initialize extractor
        extractor = CMZSalesExtractor(args.input, args.output)
        
        # Process all PDFs
        results = extractor.process_all_pdfs()
        
        if results:
            # Save results
            extractor.save_results(results, args.format)
            logger.info(f"Successfully processed {len(results)} PDF files")
        else:
            logger.warning("No PDF files were processed")
    
    except Exception as e:
        logger.error(f"Error in main execution: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
