# As-Built Extraction Tool - Complete Documentation

**A comprehensive Python system for automated organization of architectural drawings based on accommodation schedules.**

This tool processes Excel spreadsheets containing flat/house references and automatically organizes PDF drawings into structured folders with standardized naming. Designed for construction documentation workflows where hundreds of architectural drawings need to be matched and organized efficiently.

---

## Table of Contents

1. [System Overview](#system-overview)
2. [Quick Start](#quick-start)
3. [Detailed Setup](#detailed-setup)
4. [Input File Requirements](#input-file-requirements)
5. [How to Use](#how-to-use)
6. [Understanding the Output](#understanding-the-output)
7. [Excel Report Guide](#excel-report-guide)
8. [Troubleshooting](#troubleshooting)
9. [Technical Details](#technical-details)
10. [For Future Developers](#for-future-developers)

---

## System Overview

### What This Tool Does

The system automates the organization of architectural drawings by:
1. Reading flat/house type references from an accommodation schedule (Excel)
2. Finding matching drawing files in an architect spreadsheet (Excel)
3. Copying and organizing PDFs into folders by house type
4. Generating comprehensive Excel reports with statistics

**Example**: If your accommodation schedule lists "HT A 3B4P" and your architect spreadsheet has "Sections - HT A 3B4P" with filename "C0033_50.026.pdf", the tool will:
- Create folder: `processed/HT A 3B4P/`
- Copy file as: `sections_HTA3B4P_C0033_50.026.pdf`

### Key Features

✓ **Automated Matching**: Intelligent pattern matching for flat types (FT) and house types (HT)  
✓ **Multiple Drawing Types**: Handles both Sections and Floor Plans  
✓ **Standardized Naming**: Consistent filename format for easy identification  
✓ **Comprehensive Reports**: Excel reports with 10+ sheets covering all aspects  
✓ **Error Handling**: Continues processing even if some files are missing  
✓ **Validation Tools**: Built-in scripts to verify data before processing  
✓ **Detailed Logging**: Complete audit trail in log file  

### File Structure

```
as_built_extraction/
├── file_processor.py           # Main processing script
├── validate_data.py            # Validation and testing utility
├── view_excel_report.py        # Report viewer utility
├── requirements.txt            # Python dependencies
├── usage_guide.md             # Detailed usage guide
├── README.md                  # This documentation
├── .gitignore                 # Git ignore configuration
├── architect/                 # Source PDF files (468 files)
│   ├── C0033_50.001.pdf
│   ├── C0033_50.002.pdf
│   └── ...
├── spreadsheet/               # Input Excel files
│   ├── accomodation_schedule.xlsx    # Flat/house references
│   ├── architect_spreadsheet.xlsx    # Drawing index
│   └── drawing_register_external.xlsx
└── processed/                 # Organized output (created by script)
    ├── HT A 3B4P/            # Organized by house type
    │   ├── sections_HTA3B4P_C0033_50.026.pdf
    │   ├── floorplans_HTA3B4P_C0033_50.027.pdf
    │   └── ...
    ├── FT A&B/               # Flat types with combined references
    └── ...
```

---

## Quick Start

**For users who just want to run the tool:**

```powershell
# 1. Navigate to the project directory
cd "C:\Users\IS19\Documents\as_built_extraction"

# 2. Activate virtual environment
.venv\Scripts\activate

# 3. (Optional) Validate your data first
python validate_data.py

# 4. Run the main processor
python file_processor.py

# 5. (Optional) View the generated report
python view_excel_report.py
```

The tool will:
- Process all flat references from your accommodation schedule
- Match them with drawings in the architect spreadsheet
- Copy files to the `processed/` folder with organized structure
- Generate an Excel report with detailed results

---

## Detailed Setup

### Prerequisites

- **Python 3.7 or higher** (Python 3.8+ recommended)
- **Windows, macOS, or Linux** (paths may need adjustment for non-Windows)
- **Administrator/write access** to the project directory

### First-Time Installation

#### 1. Install Python

Download from [python.org](https://www.python.org/downloads/) if not already installed.

Verify installation:
```powershell
python --version
```

#### 2. Set Up Virtual Environment

A virtual environment keeps dependencies isolated from your system Python.

```powershell
# Navigate to project directory
cd "C:\Users\IS19\Documents\as_built_extraction"

# Create virtual environment
python -m venv .venv

# Activate it (Windows)
.venv\Scripts\activate

# Your prompt should now show (.venv) prefix
```

For macOS/Linux:
```bash
source .venv/bin/activate
```

#### 3. Install Dependencies

```powershell
pip install -r requirements.txt
```

This installs:
- `pandas` - Excel file processing
- `openpyxl` - Excel file reading/writing

#### 4. Verify Installation

```powershell
python -c "import pandas; import openpyxl; print('Dependencies OK')"
```

### Directory Preparation

Ensure you have these folders:
- `spreadsheet/` - Contains your Excel files
- `architect/` - Contains all PDF files to be organized
- `processed/` - Will be created automatically by the script

---

## Input File Requirements

### 1. Accommodation Schedule (accomodation_schedule.xlsx)

**Location**: `spreadsheet/accomodation_schedule.xlsx`

**Required Structure**:
- **Column D** must contain flat/house references
- Header row typically says "Flat / House Ref" or similar

**Example Data in Column D**:
```
HT A 3B4P
HT A 4B7P
HT B 2B3P
FT A 1B2P
FT C 2B3P
```

**Supported Formats**:
- House Types: `HT A 3B4P`, `HT B 2B3P`, `HT C3H5 3B5P`
- Flat Types: `FT A 1B2P`, `FT E 1B2P`, `FT H 2B4P`

### 2. Architect Spreadsheet (architect_spreadsheet.xlsx)

**Location**: `spreadsheet/architect_spreadsheet.xlsx`

**Required Structure**:
- **Column A**: Filenames (without .pdf extension)
- **Column B**: Drawing titles with house/flat type references

**Example Data**:

| Column A (Filename) | Column B (Title) |
|-------------------|------------------|
| C0033_50.026 | Sections - HT A 3B4P |
| C0033_50.027 | Floor Plans - HT A 3B4P |
| C0033_50.060 | Sections - HT B 2B3P |
| C0033_50.250 | Sections - Flat Types A & B |
| C0033_50.270 | Floor Plans - FT E 1B2P |

**Title Format Rules**:
- Must include drawing type: "Sections" or "Floor Plans" (or singular "Section"/"Floor Plan")
- Followed by: " - " (space-dash-space)
- Then the flat/house reference matching accommodation schedule

**Combined References**:
- Flat types can be grouped: `"Sections - Flat Types A & B"` → Creates folder `FT A&B`
- House types can be grouped: `"Floor Plans - HT B 2B3P & HT D 3B4P"` → Creates folder with full name

### 3. PDF Files (architect/ folder)

**Location**: `architect/`

**Requirements**:
- Files must be named exactly as listed in Column A of architect_spreadsheet.xlsx
- Files must have `.pdf` extension (case-insensitive)
- Example: If spreadsheet lists "C0033_50.026", file must be "C0033_50.026.pdf"

---

## How to Use

### Step 1: Prepare Your Data

1. **Verify Excel files are in correct location**:
   - `spreadsheet/accomodation_schedule.xlsx`
   - `spreadsheet/architect_spreadsheet.xlsx`

2. **Check Column D** in accommodation schedule contains flat references

3. **Check Columns A & B** in architect spreadsheet contain filenames and titles

4. **Verify PDF files** are in `architect/` folder

### Step 2: Validate Before Processing (Recommended)

```powershell
python validate_data.py
```

**What this does**:
- Shows structure of both spreadsheets
- Displays sample flat references from Column D
- Lists matching sections from architect spreadsheet
- Identifies potential issues before processing

**Review the output**:
- ✓ Column D contains expected flat references
- ✓ Column B contains titles in format "Sections - [reference]" or "Floor Plans - [reference]"
- ✓ Filenames in Column A match files in architect directory

### Step 3: Run the Main Processor

```powershell
python file_processor.py
```

**What happens**:

1. **Loading Phase**: Reads both Excel files
2. **Extraction Phase**: Gets all unique flat references from accommodation schedule
3. **Matching Phase**: Finds corresponding drawings in architect spreadsheet
4. **Organization Phase**: Creates folders and copies files with new names
5. **Reporting Phase**: Generates Excel report and displays statistics

**Console Output Example**:
```
2025-10-15 12:03:00 - INFO - Loading accommodation schedule spreadsheet...
2025-10-15 12:03:01 - INFO - Loaded 150 rows from accommodation schedule
2025-10-15 12:03:01 - INFO - Found 28 unique flat references
2025-10-15 12:03:02 - INFO - Match found: HT A 3B4P -> Sections - HT A 3B4P -> C0033_50.026
2025-10-15 12:03:02 - INFO - Successfully copied: C0033_50.026.pdf
...
2025-10-15 12:03:08 - INFO - ==================================================
2025-10-15 12:03:08 - INFO - PROCESSING STATISTICS
2025-10-15 12:03:08 - INFO - Total flat references processed: 28
2025-10-15 12:03:08 - INFO - Successful matches found: 122
2025-10-15 12:03:08 - INFO - No matches found: 0
2025-10-15 12:03:08 - INFO - Copy errors: 0
2025-10-15 12:03:08 - INFO - Success rate: 100.0%
```

### Step 4: Review Results

After processing completes, you'll have:

1. **Organized files** in `processed/` folder
2. **Excel report**: `processing_report_YYYYMMDD_HHMMSS.xlsx`
3. **Log file**: `file_processing.log`

### Step 5: View Excel Report (Optional)

```powershell
python view_excel_report.py
```

This shows a terminal preview of the most recent Excel report.

---

## Understanding the Output

### Output Folder Structure

```
processed/
├── HT A 3B4P/                 # One folder per house type
│   ├── sections_HTA3B4P_C0033_50.026.pdf
│   ├── sections_HTA3B4P_C0033_50.027.pdf
│   └── floorplans_HTA3B4P_C0033_50.028.pdf
├── HT B 2B3P/
│   ├── sections_HTB2B3P_C0033_50.060.pdf
│   └── floorplans_HTB2B3P_C0033_50.061.pdf
├── FT A&B/                    # Combined flat types
│   ├── sections_FTA&B_C0033_50.250.pdf
│   └── sections_FTA&B_C0033_50.251.pdf
└── HT C3H5 3B5P/              # House types with complex codes
    └── ...
```

### File Naming Convention

**Format**: `drawingtype_housetype_originalname.pdf`

**Components**:
- `drawingtype`: "sections" or "floorplans"
- `housetype`: House/flat type with spaces removed (e.g., "HTA3B4P", "FTA&B")
- `originalname`: Original filename for uniqueness (e.g., "C0033_50.026")

**Examples**:
- Original: `C0033_50.026.pdf`, Title: "Sections - HT A 3B4P"  
  → Output: `sections_HTA3B4P_C0033_50.026.pdf`

- Original: `C0033_50.250.pdf`, Title: "Sections - Flat Types A & B"  
  → Output: `sections_FTA&B_C0033_50.250.pdf`

**Why this naming**:
- Quickly identify drawing type from filename
- Know which house type without checking folder
- Maintain uniqueness with original filename
- Sortable and searchable

### Log File (file_processing.log)

Located in project root, contains:
- All console messages plus additional details
- Timestamps for each operation
- Full error messages and stack traces
- Useful for debugging or auditing

**Tip**: The log file can grow large over multiple runs. Delete or archive periodically.

---

## Excel Report Guide

The tool generates a comprehensive Excel workbook with 10 sheets providing complete processing details.

**Report Filename**: `processing_report_YYYYMMDD_HHMMSS.xlsx`  
Example: `processing_report_20251015_120308.xlsx`

### Sheet 1: Summary

**Purpose**: High-level statistics overview

| Metric | Value |
|--------|-------|
| Total Flat References in Accommodation Schedule | 28 |
| Total Section Files in Architect Spreadsheet | 468 |
| Successful Matches and Copies | 122 |
| Flat References with No Matches | 0 |
| Matches Found but Files Missing | 0 |
| Copy Errors | 0 |
| Unused Section Files | 346 |
| Overall Success Rate (%) | 100.0 |

**Use this to**: Quickly assess processing success and identify issues

### Sheet 2: Successfully Processed

**Purpose**: Complete list of all files that were successfully copied

Columns:
- **Flat Reference**: Original reference from accommodation schedule
- **Original Filename**: Source filename
- **New Filename**: Renamed destination filename
- **Title**: Full title from architect spreadsheet
- **Source Path**: Full source path
- **Destination Path**: Full destination path  
- **Destination Folder**: Output folder name

**Use this to**: 
- Verify specific files were processed
- Find where files were copied to
- Audit the renaming process

### Sheet 3: No Matches Found

**Purpose**: Flat references that had no matching drawings

Columns:
- **Flat Reference**: The unmatched reference
- **Reason**: Why no match was found

**Use this to**:
- Identify missing drawings in architect spreadsheet
- Update accommodation schedule if references are incorrect
- Request missing drawings from architect

### Sheet 4: Files Not Found

**Purpose**: Matches were found in spreadsheet but PDF doesn't exist

Columns:
- **Flat Reference**
- **Title**: Matched title from spreadsheet
- **Filename**: Expected filename
- **Expected Path**: Where file should be
- **Reason**: Error description

**Use this to**:
- Identify missing PDF files
- Check for filename spelling errors
- Request missing files

### Sheet 5: Copy Errors

**Purpose**: Files that failed to copy due to file system errors

Columns:
- **Flat Reference**
- **Title**
- **Filename**
- **Source Path**
- **Error**: Full error message

**Use this to**: 
- Troubleshoot permission issues
- Check disk space
- Identify corrupted files

### Sheet 6: Unused Section Files

**Purpose**: Drawings in architect spreadsheet not matched to any flat reference

Columns:
- **Filename**
- **Title**
- **Status**: "UNUSED"

**Use this to**:
- Identify drawings that weren't needed
- Check if accommodation schedule is missing entries
- Verify architect spreadsheet accuracy

### Sheet 7: All Flat References

**Purpose**: Complete list of all references from accommodation schedule

Columns:
- **No.**: Sequential number
- **Flat Reference**
- **Status**: "PROCESSED" or "NOT PROCESSED"
- **Files Copied**: Count of files copied for this reference

**Use this to**: See complete processing status for all references

### Sheet 8: All Section Files

**Purpose**: Complete list of all drawings from architect spreadsheet

Columns:
- **No.**
- **Filename**
- **Title**  
- **Status**: "USED" or "UNUSED"

**Use this to**: Understand which drawings were utilized

### Sheet 9: Summary by Flat Ref

**Purpose**: Processing summary grouped by flat reference

Columns:
- **Flat Reference**
- **Files Copied**: Count
- **Status**: "SUCCESS" or "NO MATCH"
- **Output Folder**: Where files were copied

**Use this to**:
- Quick reference for specific flat types
- Identify which flat types had issues
- Find output locations

### Sheet 10: Report Info

**Purpose**: Metadata about the report generation

Contains:
- Report generation date/time
- Script version
- Document types processed
- Source spreadsheet names

---

## Troubleshooting

### Common Issues and Solutions

#### Issue 1: "Column not found" error

**Error message**: `Error extracting flat references: list index out of range`

**Cause**: Column D doesn't exist or spreadsheet structure is different

**Solution**:
1. Open `accomodation_schedule.xlsx` in Excel
2. Verify flat references are in Column D (4th column)
3. If in different column, update line 109 in `file_processor.py`:
   ```python
   flat_refs_raw = self.accommodation_df.iloc[:, 3]  # Change 3 to your column index (0-based)
   ```

#### Issue 2: "File not found" errors

**Error message**: `Source file not found: C:\...\architect\C0033_50.026.pdf`

**Causes**:
- Filename in spreadsheet doesn't match actual file
- File extension is missing or different
- Filename has extra spaces or characters

**Solutions**:
1. Check architect spreadsheet Column A exactly matches PDF filenames
2. Verify all PDFs have `.pdf` extension
3. Look for hidden characters or extra spaces in spreadsheet
4. Check file actually exists in `architect/` folder

#### Issue 3: No matches found for flat references

**Error message**: `No match found for: HT A 3B4P`

**Causes**:
- Title format in architect spreadsheet is different
- Spelling differences between spreadsheets
- Drawing type not "Sections" or "Floor Plans"

**Solutions**:
1. Check architect spreadsheet Column B contains titles like:
   - "Sections - HT A 3B4P" or
   - "Floor Plans - HT A 3B4P"
2. Verify space-dash-space (" - ") separator exists
3. Match spelling exactly between both spreadsheets
4. Run `validate_data.py` to see mismatches

#### Issue 4: Script crashes or hangs

**Solutions**:
1. Check Python version: `python --version` (need 3.7+)
2. Reinstall dependencies: `pip install -r requirements.txt --force-reinstall`
3. Check spreadsheet files aren't open in Excel (closes file locks)
4. Review `file_processing.log` for specific error
5. Try processing with smaller test dataset first

#### Issue 5: Permission denied errors

**Error**: `PermissionError: [Errno 13] Permission denied`

**Solutions**:
1. Close Excel if spreadsheet is open
2. Run Command Prompt/PowerShell as Administrator
3. Check folder permissions for `processed/` directory
4. Verify antivirus isn't blocking file operations

#### Issue 6: Excel report not generating

**Solutions**:
1. Check disk space (reports can be large)
2. Verify `openpyxl` is installed: `pip install openpyxl`
3. Close any existing report files with same name
4. Check file_processing.log for specific error

### Getting Help

1. **Check the log file**: `file_processing.log` contains detailed error messages
2. **Run validation**: `python validate_data.py` shows spreadsheet structure
3. **Review usage guide**: `usage_guide.md` has additional examples
4. **Check this README**: Search for your error message

---

## Technical Details

### Code Architecture

**FileProcessor Class** (`file_processor.py`):
- Main processing engine
- Methods:
  - `load_spreadsheets()` - Reads Excel files into pandas DataFrames
  - `get_flat_references()` - Extracts unique references from Column D
  - `find_matching_drawings()` - Pattern matching logic
  - `extract_folder_name_from_title()` - Parses titles to get folder names
  - `create_output_structure()` - Creates directory structure
  - `generate_new_filename()` - Applies naming convention
  - `copy_file()` - Handles file copying with error handling
  - `generate_detailed_report()` - Creates Excel report

### Matching Logic

The tool uses intelligent pattern matching:

**For House Types (HT)**:
- Exact match: "HT A 3B4P"
- Base match: "HT A" (matches variations like "HT A 3B4P", "HT A 4B7P")
- With bed/person: "HT C3H5 3B5P"
- Combined: "HT B 2B3P & HT D 3B4P"

**For Flat Types (FT)**:
- Full match: "FT A 1B2P"
- Letter match: "FT A" (matches "Flat Type A", "Flat Types A & B", "FT A 1B2P")
- Abbreviated: "Flat Types A & B" → "FT A&B"
- Combined: "FT A 1B2P & FT B 1B2P" → "FT A&B"

### Dependencies

```
pandas>=1.3.0        # Excel file processing and data manipulation
openpyxl>=3.0.7      # Excel file format support (.xlsx)
```

### Paths and Configuration

Hardcoded paths in `file_processor.py` (lines 39-46):
```python
self.base_path = Path(r"C:\Users\IS19\Documents\as_built_extraction")
self.spreadsheet_path = self.base_path / "spreadsheet"
self.architect_path = self.base_path / "architect"
self.processed_path = self.base_path / "processed"
```

**To use in different location**: Update `base_path` in line 39

### Performance

- Processing ~470 PDF files takes approximately 5-10 seconds
- Excel report generation adds ~2-3 seconds
- Memory usage: ~50-100 MB for typical datasets
- Disk space: Output roughly equals input size (files are copied, not moved)

---

## For Future Developers

### Modifying for Different Use Cases

#### Change to Different Drawing Types

To process "Roof Plans" instead of "Sections":

1. Update drawing type search in `find_matching_drawings()` (line 156):
   ```python
   # Change from:
   drawing_mask = titles.str.contains("Sections|Floor Plans", case=False, na=False, regex=True)
   # To:
   drawing_mask = titles.str.contains("Roof Plans", case=False, na=False, regex=True)
   ```

2. Update drawing type determination (line 207):
   ```python
   if "Roof Plans" in title:
       drawing_type = "roofplans"
   ```

#### Use Different Spreadsheet Columns

To read flat references from Column E instead of Column D:

Update line 109:
```python
flat_refs_raw = self.accommodation_df.iloc[:, 4]  # 4 = Column E (0-indexed)
```

#### Change Output Folder Structure

To organize by drawing type first, then house type:

Modify `create_output_structure()` (line 329):
```python
def create_output_structure(self, folder_name, drawing_type):
    output_path = self.processed_path / drawing_type / folder_name
    output_path.mkdir(parents=True, exist_ok=True)
    return output_path
```

#### Customize File Naming

To use different naming format, modify `generate_new_filename()` (line 342):

Current format: `drawingtype_housetype_originalname.pdf`

Example alternative: `housetype_drawingtype_originalname.pdf`
```python
new_filename = f"{clean_folder}_{drawing_type}_{orig_name}{file_ext}"
```

### Adding New Features

#### Add Progress Bar

Install tqdm: `pip install tqdm`

Add to imports:
```python
from tqdm import tqdm
```

Wrap processing loop (around line 469):
```python
for match in tqdm(matches, desc="Processing files"):
    # existing code
```

#### Add Email Notifications

Install: `pip install yagmail`

Add to end of `process_files()`:
```python
import yagmail
yag = yagmail.SMTP('your-email@gmail.com')
yag.send(
    to='recipient@example.com',
    subject='Processing Complete',
    contents=f'Processed {self.stats["matched"]} files successfully'
)
```

#### Support Additional File Types

To process DWG files in addition to PDF:

1. Update file search in `copy_file()` (line 378):
   ```python
   for ext in ['.pdf', '.dwg', '.dxf']:
       source_file = self.architect_path / (source_filename + ext)
       if source_file.exists():
           break
   ```

### Testing Changes

Always test with a small subset first:

1. Create test folders:
   ```powershell
   mkdir test_architect
   mkdir test_spreadsheet
   ```

2. Copy 5-10 sample files

3. Create test spreadsheet with just those references

4. Update paths in script temporarily

5. Run and verify output

6. Once confirmed, apply to full dataset

### Version Control Best Practices

When making changes:

1. **Document changes** in version history section of README
2. **Update usage_guide.md** if workflows change
3. **Test thoroughly** with sample data
4. **Keep backups** of original spreadsheets and architect folder
5. **Tag versions** in git: `git tag v2.1`

### Maintenance Tasks

**Regular**:
- Clear old report files (keep last 5)
- Archive or delete log file when large
- Verify backup of source PDFs exists

**After Major Updates**:
- Update version number in `generate_detailed_report()` (line 569)
- Update this README with changes
- Test with full dataset

### Support Contacts

For handoff questions:
- Check Git commit history for context
- Review `file_processing.log` for usage patterns
- Test with `validate_data.py` when troubleshooting

---

## License

This project is licensed under the MIT License - see the LICENSE file for details.

---

## Quick Reference Card

### Essential Commands

```powershell
# Activate environment
.venv\Scripts\activate

# Validate data
python validate_data.py

# Process files
python file_processor.py

# View report
python view_excel_report.py
```

### Key Files

- **Input**: `spreadsheet/accomodation_schedule.xlsx`, `spreadsheet/architect_spreadsheet.xlsx`
- **Source**: `architect/*.pdf`
- **Output**: `processed/[HouseType]/*.pdf`
- **Report**: `processing_report_*.xlsx`
- **Log**: `file_processing.log`

### Common Paths to Update

| Line | File | What to Change |
|------|------|----------------|
| 39 | file_processor.py | Base path location |
| 109 | file_processor.py | Accommodation schedule column |
| 156 | file_processor.py | Drawing types to match |
| 342 | file_processor.py | File naming format |

---

**Documentation Version**: 2.0  
**Last Updated**: February 2026  
**For**: Project Handoff