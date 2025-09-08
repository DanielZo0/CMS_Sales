# CMZ Sales PDF Data Extraction

A Python script for extracting sales data from PDF files using multiple PDF processing libraries.

## Features

- Extract text content from PDF files
- Extract tables from PDF files
- Support for multiple PDF processing libraries (PyPDF2, pdfplumber)
- Export results to Excel, CSV, or JSON formats
- Comprehensive logging and error handling
- Command-line interface for easy automation

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

### Basic Usage

1. Place your PDF files in the `input` directory
2. Run the script:

```bash
python CMZ_Sales.py
```

3. Extracted data will be saved in the `output` directory

### Command Line Options

```bash
python CMZ_Sales.py [OPTIONS]

Options:
  -i, --input DIR     Input directory containing PDF files (default: input)
  -o, --output DIR    Output directory for extracted data (default: output)
  -f, --format FORMAT Output format: excel, csv, or json (default: excel)
  -v, --verbose       Enable verbose logging
  -h, --help          Show help message
```

### Examples

```bash
# Process PDFs from custom input directory
python CMZ_Sales.py --input /path/to/pdfs --output /path/to/results

# Export results as CSV
python CMZ_Sales.py --format csv

# Enable verbose logging
python CMZ_Sales.py --verbose
```

## Project Structure

```
CMZ_Sales/
├── CMZ_Sales.py          # Main extraction script
├── requirements.txt      # Python dependencies
├── README.md            # This file
├── .gitignore           # Git ignore file
├── input/               # Directory for input PDF files
└── output/              # Directory for extracted data
```

## Output Formats

### Excel Format (.xlsx)
- **Summary Sheet**: Overview of processed files
- **Text_Content Sheet**: Extracted text from all PDFs
- **Tables Sheet**: All extracted tables combined

### CSV Format (.csv)
- Single file with flattened data structure
- Includes filename, processing info, and content

### JSON Format (.json)
- Complete data structure with nested tables
- Preserves original data relationships

## Dependencies

- **PyPDF2**: Basic PDF text extraction
- **pdfplumber**: Advanced PDF processing with table extraction
- **pandas**: Data manipulation and analysis
- **openpyxl**: Excel file handling
- **python-dotenv**: Environment variable management

## Error Handling

The script includes comprehensive error handling:
- Missing dependencies are detected and reported
- Individual PDF processing errors don't stop the entire batch
- All errors are logged to both console and log file
- Error information is included in output files

## Logging

Logs are written to:
- Console output (INFO level by default)
- `cmz_sales.log` file (all levels)

Use `--verbose` flag for detailed debug information.

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

- **pdfplumber**: Best for PDFs with tables and complex layouts
- **PyPDF2**: Good for simple text extraction
- Some PDFs may require specific processing methods

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
