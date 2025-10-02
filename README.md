# File Processor - As-Built Extraction Tool

A Python-based file processing tool designed to organize architect drawings and documents based on accommodation schedules. This tool processes Excel spreadsheets to match flat/house references with section drawings and organizes them into a structured directory system.

## Features

- **Automated File Organization**: Organizes files by House Type with standardized naming
- **Excel Integration**: Reads accommodation schedules and architect spreadsheets
- **Intelligent Matching**: Matches flat references with section titles using pattern matching
- **Comprehensive Reporting**: Generates detailed Excel reports with processing statistics
- **Error Handling**: Robust error handling with detailed logging
- **File Validation**: Checks for file existence and handles missing files gracefully

## Project Structure

```
file_processor/
├── file_processor.py          # Main processing script
├── usage_guide.md            # Detailed usage instructions
├── validate_data.py          # Data validation utilities
├── view_excel_report.py      # Report viewing utilities
├── requirements.txt          # Python dependencies
├── README.md                # This file
└── .gitignore               # Git ignore rules
```

## Output Structure

Files are organized as follows:
```
processed/
├── HT A 3B4P/
│   ├── sections_HTA3B4P_C0033_50.026.pdf
│   ├── sections_HTA3B4P_C0033_50.027.pdf
│   └── ...
├── HT B 2B3P/
│   ├── sections_HTB2B3P_C0033_50.060.pdf
│   └── ...
└── ...
```

**File Naming Convention**: `drawingtype_housetype_originalname.pdf`
- `drawingtype`: Type of drawing (e.g., "sections")
- `housetype`: House type with spaces removed (e.g., "HTA3B4P")
- `originalname`: Original filename for uniqueness

## Prerequisites

- Python 3.7+
- Required Python packages (see requirements.txt)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/imransal/file_processor.git
cd file_processor
```

2. Create a virtual environment:
```bash
python -m venv .venv
```

3. Activate the virtual environment:
```bash
# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate
```

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Setup

1. Create the following directory structure:
```
project_root/
├── spreadsheet/
│   ├── accomodation_schedule.xlsx
│   └── architect_spreadsheet.xlsx
└── architect/
    └── [PDF files to be processed]
```

2. Ensure your Excel files have the correct format:
   - **accommodation_schedule.xlsx**: Flat/House references in Column D
   - **architect_spreadsheet.xlsx**: Filenames in Column A, Titles in Column B

## Usage

### Basic Usage

Run the main processing script:
```bash
python file_processor.py
```

### Advanced Usage

For detailed usage instructions, see [usage_guide.md](usage_guide.md).

## Input Data Format

### Accommodation Schedule (Column D)
- Contains flat/house references like: `HT A 3B4P`, `FT A 1B2P`, `HT B 2B3P`

### Architect Spreadsheet
- **Column A**: Filenames (e.g., `C0033_50.026`)
- **Column B**: Titles (e.g., `Sections - HT A 3B4P`)

## Output

### Processed Files
- Files are copied to `processed/` directory
- Organized by House Type
- Renamed with standardized convention

### Reports
- Detailed Excel report generated after processing
- Includes statistics, successful matches, errors, and unused files
- Named: `processing_report_YYYYMMDD_HHMMSS.xlsx`

### Logs
- Comprehensive logging to `file_processing.log`
- Console output for real-time monitoring

## Processing Statistics

The tool provides comprehensive statistics including:
- Total flat references processed
- Successful matches and copies
- Files not found
- Copy errors
- Success rate percentage

## Error Handling

- **File Not Found**: Logs missing files and continues processing
- **Copy Errors**: Handles file system errors gracefully
- **Data Validation**: Validates input data format
- **Detailed Reporting**: All errors tracked in Excel report

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For questions or issues, please:
1. Check the [usage_guide.md](usage_guide.md) for detailed instructions
2. Review the generated log files for error details
3. Open an issue on GitHub

## Version History

- **v2.0**: Current version with House Type organization and standardized naming
- **v1.0**: Initial version with Sections-based organization

## Author

Created for as-built extraction and document organization workflows.