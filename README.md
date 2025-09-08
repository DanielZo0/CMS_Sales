# CMZ Sales PDF Data Extraction

A simple Python script for extracting sales data from CMZ weekly sales PDF files.

## Features

- Extract specific sales data from CMZ weekly sales PDFs
- Support for pdfplumber library for reliable PDF processing
- Export results to JSON format matching the required template
- Simple and focused functionality

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. Clone or download this repository
2. Navigate to the project directory
3. Create and activate a virtual environment:

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

4. Install required dependencies:

```bash
pip install -r requirements.txt
```

### Quick Setup (Windows)

For Windows users, you can use the provided batch file:

```bash
# Run the activation script
activate_env.bat
```

## Usage

### Sales Data Extraction

1. Place your CMZ weekly sales PDF files in the `input/WeeklySales/` directory
2. Run the extraction script:

```bash
python extract_sales_data.py
```

This script will:
- Process all PDF files in `input/WeeklySales/` directory
- Extract Document Date, Document Number, Description, Locality, and Total
- Output results in JSON format matching the template structure
- Save results to `output/weekly_sales_data.json`

### Output Format

The script extracts the following fields:
- **Document Date**: Date from the PDF
- **Document Number**: Reference number from the PDF
- **Description**: Formatted as "Chain Supermarket [Locality] - Wk[Week] ([StartDate]-[EndDate])"
- **Locality**: Location name (e.g., Tarxien, Fgura, Zabbar)
- **Total**: Sales amount in euros

### Example Output

```json
[
  {
    "Document Date": "30/6/25",
    "Document Number": "Chaintrx-2025224",
    "Description": "Chain Supermarket Tarxien - Wk24 (08.06.25-14.06.25)",
    "Locality": "Tarxien",
    "Total": "€13,769.01"
  }
]
```

## Project Structure

```
CMZ_Sales/
├── extract_sales_data.py # Main extraction script
├── requirements.txt      # Python dependencies
├── README.md            # This file
├── .gitignore           # Git ignore file
├── input/               # Directory for input files
│   ├── WeeklySales/     # PDF files to process
│   └── JSON_Template_Sales.json # Template structure
└── output/              # Directory for extracted data
```

## Output Format

### JSON Format (.json)
- Clean data structure matching the template
- One record per PDF file processed
- All required fields extracted and formatted

## Dependencies

- **pdfplumber**: PDF processing and text extraction
- **pandas**: Data manipulation (for potential future enhancements)
- **python-dotenv**: Environment variable management

## Error Handling

The script includes error handling:
- Missing dependencies are detected and reported
- Individual PDF processing errors don't stop the entire batch
- All errors are logged to console output
- Processing continues even if individual files fail

## Troubleshooting

### Common Issues

1. **"No PDF processing libraries available"**
   - Install dependencies: `pip install -r requirements.txt`

2. **"No PDF files found in input directory"**
   - Ensure PDF files are in the `input` directory
   - Check file extensions are `.pdf`

3. **Permission errors**
   - Ensure write permissions for output directory
   - Close any open Excel files before running

### PDF Compatibility

- **pdfplumber**: Optimized for CMZ weekly sales PDF format
- Handles the specific structure of CMZ sales invoices
- Extracts text and parses structured data reliably

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source. Please check the license file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review log files for error details
3. Create an issue in the repository
