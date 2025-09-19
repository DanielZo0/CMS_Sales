# CMZ Combined Sales Data Extractor

A unified Python script for processing both PDF and Excel sales data files with automatic file organization and data extraction.

## Features

- **Dual Format Support**: Process both PDF and Excel files in a single run
- **PDF Processing**: Extract sales data from CMZ weekly sales PDFs
- **Excel Processing**: Extract data from Excel files according to JSON template structure
- **File Standardization**: Rename PDF files to standard format "Weekly Sales - <Locality> Week <Number>"
- **Invoice Organization**: Create renamed copies of invoice files with standardized naming format
- **Multi-Format Output**: Export results to both JSON and CSV formats
- **Multi-Locality Support**: Handle Fgura, Tarxien, Zabbar, and Carters locations
- **Robust Error Handling**: Graceful fallback if dependencies are missing

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. Clone or download this repository
2. Navigate to the project directory
3. Install required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the combined script to process all data files:

```bash
python combined_data_extractor.py
```

The script will automatically:
1. Rename PDF files in `input/WeeklySales/` to standard format
2. Extract data from PDF files (if pdfplumber is available)
3. Extract data from Excel files with "delicatessen" pattern
4. Create individual JSON files for each invoice
5. Generate renamed copies of invoices with standardized naming
6. Combine all results into unified output files

## File Structure

```
CMZ_Sales/
├── input/
│   ├── WeeklySales/              # PDF files for processing and renaming
│   ├── output/                   # Generated output files
│   │   ├── individual_json/      # Individual JSON files per invoice
│   │   ├── renamed_invoices/     # Renamed invoice copies
│   │   ├── *.csv                 # CSV output files
│   │   └── *.json                # Combined JSON output files
│   └── templates/                # JSON template files
├── combined_data_extractor.py    # Main unified script
├── gui_extractor.py             # GUI interface
├── main.py                      # Alternative main entry point
├── requirements.txt             # Dependencies
└── README.md                    # This file
```

## Dependencies

- **pandas**: Data manipulation and Excel file processing
- **openpyxl**: Excel file reading support
- **pdfplumber**: PDF text extraction (optional)
- **python-dotenv**: Environment variable management

## Data Processing

The script processes:
- **PDF files**: Extracts sales data using text pattern matching
- **Excel files**: Extracts data according to JSON template structure
- **File organization**: Standardizes PDF filenames for consistency
- **Invoice management**: Creates organized copies with standardized naming

### Invoice Organization

The system automatically creates a `renamed_invoices` folder containing copies of all invoice JSON files with standardized names. This feature:

- Extracts location information from the invoice data (Locality field)
- Extracts week numbers from the description field (e.g., "Wk19", "Wk20")
- Creates files named: `Weekly Sales - [Location] Wk'[XX]'.json`
- Maintains original files unchanged in the `individual_json` folder
- Provides easy identification and organization of invoices by location and week

### Extracted Fields

- Document Type, Date, Number
- Supplier codes (locality-specific)
- Description (formatted with locality and week information)
- Net amount, VAT amount, Total
- Automatic locality detection from filenames

## Output Files

### Main Output Files
- `input/output/pdf_sales_data.json`: Combined sales data in JSON format
- `input/output/pdf_sales_data.csv`: Combined sales data in CSV format
- `input/output/purchases_data.json`: Combined purchases data in JSON format
- `input/output/purchases_data.csv`: Combined purchases data in CSV format
- `input/output/combined_sales_purchases.csv`: Unified sales and purchases data
- `input/output/upload_data.csv`: Data formatted for upload according to upload template

### Individual Files
- `input/output/individual_json/`: Individual JSON files for each processed invoice
- `input/output/renamed_invoices/`: Renamed copies of invoices with format "Weekly Sales - Location Wk'xx'"

### Invoice Naming Convention
Invoice files in the `renamed_invoices` folder follow the format:
- `Weekly Sales - Fgura Wk'19'.json`
- `Weekly Sales - Tarxien Wk'20'.json`
- `Weekly Sales - Zabbar Wk'21'.json`

Where:
- **Location** is extracted from the invoice data (Fgura, Tarxien, Zabbar)
- **Week number** is extracted from the description field (e.g., Wk'19', Wk'20')

## Error Handling

- Works even if PDF processing libraries are unavailable
- Individual file processing errors don't stop the batch
- Clear error messages and progress reporting
- Automatic fallback to default templates if JSON template missing