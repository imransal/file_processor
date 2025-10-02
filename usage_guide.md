# File Processing Script Usage Guide

## Overview
This script processes accommodation and architect spreadsheets to automatically organize files based on flat/house references.

## Files Created
1. `file_processor.py` - Main processing script
2. `validate_data.py` - Data validation and testing script
3. `usage_guide.md` - This file

## Before Running
1. Ensure your spreadsheets are in the correct location:
   - `C:\Users\IS19\Documents\as_built_extraction\spreadsheet\accomodation_schedule.xlsx`
   - `C:\Users\IS19\Documents\as_built_extraction\spreadsheet\architect_spreadsheet.xlsx`

2. Make sure the architect files are in:
   - `C:\Users\IS19\Documents\as_built_extraction\architect\`

## Step-by-Step Usage

### Step 1: Validate Your Data First
Run the validation script to examine your spreadsheets and test the matching logic:

```powershell
cd "C:\Users\IS19\Documents\as_built_extraction"
C:/Users/IS19/Documents/as_built_extraction/.venv/Scripts/python.exe validate_data.py
```

This will:
- Show the structure of both spreadsheets
- Display sample data from each column
- Test the matching logic with a few examples
- Check if referenced files actually exist

### Step 2: Review the Output
Look at the validation output to ensure:
- Column D contains the flat references (like "HT A 3B4P")
- Column B in architect spreadsheet contains titles (like "Sections - HT A 3B4P")
- Column A in architect spreadsheet contains filenames
- Files exist in the architect directory

### Step 3: Run the Main Processing Script
Once validation looks good, run the main script:

```powershell
cd "C:\Users\IS19\Documents\as_built_extraction"
C:/Users/IS19/Documents/as_built_extraction/.venv/Scripts/python.exe file_processor.py
```

## Expected Output Structure
The script will create folders in `C:\Users\IS19\Documents\as_built_extraction\processed\`:

```
processed/
└── Sections/
    ├── HT A 3B4P/
    │   └── [matching files]
    ├── HT B 2B3P/
    │   └── [matching files]
    └── ... (other flat references)
```

## What the Script Does
1. **Reads** flat/house references from Column D of `accomodation_schedule.xlsx`
2. **Searches** for matching titles in Column B of `architect_spreadsheet.xlsx` that start with "Sections - [flat_ref]"
3. **Gets** the corresponding filename from Column A
4. **Finds** the actual file in the architect directory
5. **Copies** the file to the organized folder structure in processed directory

## Logs and Monitoring
- The script creates a `file_processing.log` file with detailed information
- Progress is also displayed in the terminal
- Statistics are shown at the end (matches found, files copied, errors, etc.)

## Troubleshooting

### Common Issues
1. **"Column not found"** - Check if your spreadsheets have the expected column structure
2. **"File not found"** - Verify that filenames in the spreadsheet match actual files in the architect directory
3. **"No matches found"** - Check if the title format in architect spreadsheet matches "Sections - [flat_ref]"

### Customization
If you need to modify the matching logic:
- Edit the `target_pattern` in the `find_matching_sections()` method
- Adjust column indices if your spreadsheet structure is different
- Modify the output folder structure in `create_output_structure()`

## Future Extensions
The script is designed to be easily extended for other document types:
- Change "Sections" to "Roof Plan" for roof plans
- Add multiple document types in a single run
- Modify the folder structure as needed

## Excel Reports
The script now generates comprehensive Excel reports with multiple sheets:

### Report Sheets:
1. **Summary** - Overall statistics and success rates
2. **Successfully Processed** - All files that were copied with full paths
3. **No Matches Found** - Flat references that had no matching sections
4. **Files Not Found** - Matches found but files missing from architect directory
5. **Copy Errors** - Any files that failed to copy
6. **Unused Section Files** - Section files not matched to any flat reference
7. **All Flat References** - Complete list with processing status
8. **All Section Files** - Complete list with usage status
9. **Summary by Flat Ref** - Processing summary grouped by flat reference
10. **Report Info** - Metadata about the report generation

### Using the Excel Report:
- Open in Excel for sorting, filtering, and analysis
- Use pivot tables for custom summaries
- Filter by status to see specific categories
- Copy data for presentations or further analysis

### Viewing Reports:
Run the viewer script to see report contents in the terminal:
```powershell
C:/Users/IS19/Documents/as_built_extraction/.venv/Scripts/python.exe view_excel_report.py
```

## Support
Check the log file `file_processing.log` for detailed error messages and processing information.