#!/usr/bin/env python3
"""
File Processing Script for As-Built Extraction

This script processes accommodation schedule and architect spreadsheets to:
1. Read flat/house references from accommodation_schedule.xlsx (Column D)
2. Match them with titles in architect_spreadsheet.xlsx (Column B) that start with "Sections"
3. Find corresponding filenames in architect_spreadsheet.xlsx (Column A)
4. Copy matching files from the architect directory to organized folder structure in processed directory
5. Files are organized by House Type and renamed with format: drawingtype_housetype.pdf

Example output structure:
- processed/HT A 3B4P/sections_HTA3B4P.pdf

Author: GitHub Copilot
Date: October 2025
"""

import pandas as pd
import os
import shutil
from pathlib import Path
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('file_processing.log'),
        logging.StreamHandler()
    ]
)

class FileProcessor:
    def __init__(self):
        # Define base paths
        self.base_path = Path(r"C:\Users\IS19\Documents\as_built_extraction")
        self.spreadsheet_path = self.base_path / "spreadsheet"
        self.architect_path = self.base_path / "architect"
        self.processed_path = self.base_path / "processed"
        
        # Define spreadsheet files
        self.accommodation_file = self.spreadsheet_path / "accomodation_schedule.xlsx"
        self.architect_file = self.spreadsheet_path / "architect_spreadsheet.xlsx"
        
        # Statistics and detailed tracking
        self.stats = {
            'processed': 0,
            'matched': 0,
            'not_found': 0,
            'copy_errors': 0
        }
        
        # Detailed tracking for reporting
        self.detailed_results = {
            'successful_matches': [],      # Files successfully processed
            'no_matches_found': [],        # Flat refs with no matching sections
            'file_not_found': [],          # Matches found but file missing
            'copy_errors': [],             # Files that failed to copy
            'unused_section_files': [],    # Section files not matched to any flat ref
            'all_flat_refs': [],           # All flat references processed
            'all_section_files': []        # All section files in architect spreadsheet
        }
    
    def load_spreadsheets(self):
        """Load both Excel spreadsheets into pandas DataFrames"""
        try:
            logging.info("Loading accommodation schedule spreadsheet...")
            self.accommodation_df = pd.read_excel(self.accommodation_file)
            logging.info(f"Loaded {len(self.accommodation_df)} rows from accommodation schedule")
            
            logging.info("Loading architect spreadsheet...")
            self.architect_df = pd.read_excel(self.architect_file)
            logging.info(f"Loaded {len(self.architect_df)} rows from architect spreadsheet")
            
            # Display column info for debugging
            logging.info(f"Accommodation columns: {list(self.accommodation_df.columns)}")
            logging.info(f"Architect columns: {list(self.architect_df.columns)}")
            
            return True
            
        except Exception as e:
            logging.error(f"Error loading spreadsheets: {str(e)}")
            return False
    
    def get_flat_references(self):
        """Extract flat/house references from Column D of accommodation schedule"""
        try:
            # Assuming Column D contains the flat references
            # Get the 4th column (index 3) which should be Column D
            flat_refs_raw = self.accommodation_df.iloc[:, 3].dropna().unique()
            
            # Filter out header and non-data entries
            flat_refs = []
            for ref in flat_refs_raw:
                ref_str = str(ref).strip()
                # Skip header rows and empty strings
                if ref_str and ref_str != "Flat / House Ref" and not ref_str.startswith("Flat"):
                    flat_refs.append(ref_str)
            
            logging.info(f"Found {len(flat_refs)} unique flat references")
            logging.info(f"Sample references: {list(flat_refs[:5])}")
            
            return flat_refs
            
        except Exception as e:
            logging.error(f"Error extracting flat references: {str(e)}")
            return []
    
    def find_matching_sections(self, flat_refs):
        """Find matching sections in architect spreadsheet for given flat references"""
        matches = []
        
        try:
            # Get Column B (Title) - assuming it's the 2nd column (index 1)
            titles = self.architect_df.iloc[:, 1]  # Column B
            filenames = self.architect_df.iloc[:, 0]  # Column A
            
            # Store all flat refs for reporting
            self.detailed_results['all_flat_refs'] = list(flat_refs)
            
            # Find all section files for reporting
            section_mask = titles.str.contains("Sections", case=False, na=False)
            for idx in section_mask[section_mask].index:
                self.detailed_results['all_section_files'].append({
                    'filename': filenames.iloc[idx],
                    'title': titles.iloc[idx],
                    'matched': False  # Will be updated when matches are found
                })
            
            for flat_ref in flat_refs:
                self.stats['processed'] += 1
                
                # Look for titles that start with "Sections" and contain the flat reference
                target_pattern = f"Sections - {flat_ref}"
                
                # Find matching rows
                matching_rows = titles.str.contains(target_pattern, case=False, na=False)
                
                if matching_rows.any():
                    # Get all matches (there might be multiple)
                    for idx in matching_rows[matching_rows].index:
                        filename = filenames.iloc[idx]
                        title = titles.iloc[idx]
                        
                        match_info = {
                            'flat_ref': flat_ref,
                            'title': title,
                            'filename': filename,
                            'row_index': idx
                        }
                        
                        matches.append(match_info)
                        
                        # Mark this section file as matched
                        for section_file in self.detailed_results['all_section_files']:
                            if section_file['filename'] == filename and section_file['title'] == title:
                                section_file['matched'] = True
                                break
                        
                        logging.info(f"Match found: {flat_ref} -> {title} -> {filename}")
                        self.stats['matched'] += 1
                else:
                    logging.warning(f"No match found for: {flat_ref}")
                    self.stats['not_found'] += 1
                    self.detailed_results['no_matches_found'].append({
                        'flat_ref': flat_ref,
                        'reason': 'No matching section found in architect spreadsheet'
                    })
            
            # Find unused section files
            for section_file in self.detailed_results['all_section_files']:
                if not section_file['matched']:
                    self.detailed_results['unused_section_files'].append(section_file)
            
            logging.info(f"Total matches found: {len(matches)}")
            return matches
            
        except Exception as e:
            logging.error(f"Error finding matches: {str(e)}")
            return []
    
    def create_output_structure(self, flat_ref):
        """Create the output folder structure: HT A 3B4P > files"""
        try:
            output_path = self.processed_path / flat_ref
            output_path.mkdir(parents=True, exist_ok=True)
            
            logging.info(f"Created directory structure: {output_path}")
            return output_path
            
        except Exception as e:
            logging.error(f"Error creating directory structure: {str(e)}")
            return None
    
    def generate_new_filename(self, original_filename, flat_ref, drawing_type="sections"):
        """Generate new filename in format drawingtype_housetype_originalname.pdf"""
        try:
            # Remove spaces from flat_ref and convert to format like HTA3B4P
            house_type = flat_ref.replace(" ", "")
            
            # Get the original filename without extension for uniqueness
            orig_name = original_filename
            if "." in original_filename:
                orig_name = original_filename.rsplit(".", 1)[0]  # Remove extension
                file_ext = "." + original_filename.rsplit(".", 1)[1]
            else:
                file_ext = ".pdf"  # Default to PDF
            
            # Create new filename: drawingtype_housetype_originalname.ext
            new_filename = f"{drawing_type}_{house_type}_{orig_name}{file_ext}"
            
            logging.info(f"Generated new filename: {original_filename} -> {new_filename}")
            return new_filename
            
        except Exception as e:
            logging.error(f"Error generating filename for {original_filename}: {str(e)}")
            return original_filename
    
    def copy_file(self, source_filename, destination_path, match_info):
        """Copy file from architect directory to destination"""
        try:
            # Find the file in architect directory
            # Try with .pdf extension if not already present
            if not source_filename.lower().endswith('.pdf'):
                source_filename_with_ext = source_filename + '.pdf'
            else:
                source_filename_with_ext = source_filename
            
            source_file = self.architect_path / source_filename_with_ext
            
            if not source_file.exists():
                logging.error(f"Source file not found: {source_file}")
                self.stats['copy_errors'] += 1
                
                # Track file not found
                self.detailed_results['file_not_found'].append({
                    'flat_ref': match_info['flat_ref'],
                    'title': match_info['title'],
                    'filename': source_filename,
                    'expected_path': str(source_file),
                    'reason': 'File not found in architect directory'
                })
                return False
            
            # Generate new filename
            new_filename = self.generate_new_filename(source_filename_with_ext, match_info['flat_ref'])
            
            # Create destination file path with new filename
            dest_file = destination_path / new_filename
            
            # Copy the file
            shutil.copy2(source_file, dest_file)
            logging.info(f"Successfully copied: {source_filename_with_ext} -> {dest_file}")
            
            # Track successful copy
            self.detailed_results['successful_matches'].append({
                'flat_ref': match_info['flat_ref'],
                'title': match_info['title'],
                'original_filename': source_filename_with_ext,
                'new_filename': new_filename,
                'source_path': str(source_file),
                'destination_path': str(dest_file),
                'destination_folder': str(destination_path)
            })
            
            return True
            
        except Exception as e:
            logging.error(f"Error copying file {source_filename}: {str(e)}")
            self.stats['copy_errors'] += 1
            
            # Track copy error
            self.detailed_results['copy_errors'].append({
                'flat_ref': match_info['flat_ref'],
                'title': match_info['title'],
                'filename': source_filename,
                'source_path': str(source_file) if 'source_file' in locals() else 'Unknown',
                'error': str(e)
            })
            return False
    
    def process_files(self):
        """Main processing function"""
        logging.info("Starting file processing...")
        
        # Load spreadsheets
        if not self.load_spreadsheets():
            logging.error("Failed to load spreadsheets. Exiting.")
            return False
        
        # Get flat references
        flat_refs = self.get_flat_references()
        if not flat_refs:
            logging.error("No flat references found. Exiting.")
            return False
        
        # Find matching sections
        matches = self.find_matching_sections(flat_refs)
        if not matches:
            logging.warning("No matches found.")
            return True
        
        # Process each match
        for match in matches:
            flat_ref = match['flat_ref']
            filename = match['filename']
            
            # Create output directory structure
            output_dir = self.create_output_structure(flat_ref)
            if output_dir is None:
                continue
            
            # Copy the file
            self.copy_file(filename, output_dir, match)
        
        # Print statistics
        self.print_statistics()
        
        # Generate detailed report
        self.generate_detailed_report()
        
        logging.info("File processing completed!")
        return True
    
    def print_statistics(self):
        """Print processing statistics"""
        logging.info("=" * 50)
        logging.info("PROCESSING STATISTICS")
        logging.info("=" * 50)
        logging.info(f"Total flat references processed: {self.stats['processed']}")
        logging.info(f"Successful matches found: {self.stats['matched']}")
        logging.info(f"No matches found: {self.stats['not_found']}")
        logging.info(f"Copy errors: {self.stats['copy_errors']}")
        logging.info(f"Success rate: {(self.stats['matched'] / max(1, self.stats['processed'])) * 100:.1f}%")
        logging.info("=" * 50)

    def generate_detailed_report(self):
        """Generate comprehensive detailed report as Excel spreadsheet"""
        from datetime import datetime
        
        report_filename = f"processing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = self.base_path / report_filename
        
        try:
            with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
                
                # 1. Summary Statistics Sheet
                summary_data = {
                    'Metric': [
                        'Total Flat References in Accommodation Schedule',
                        'Total Section Files in Architect Spreadsheet',
                        'Successful Matches and Copies',
                        'Flat References with No Matches',
                        'Matches Found but Files Missing',
                        'Copy Errors',
                        'Unused Section Files',
                        'Overall Success Rate (%)'
                    ],
                    'Value': [
                        len(self.detailed_results['all_flat_refs']),
                        len(self.detailed_results['all_section_files']),
                        len(self.detailed_results['successful_matches']),
                        len(self.detailed_results['no_matches_found']),
                        len(self.detailed_results['file_not_found']),
                        len(self.detailed_results['copy_errors']),
                        len(self.detailed_results['unused_section_files']),
                        round((len(self.detailed_results['successful_matches']) / max(1, len(self.detailed_results['all_flat_refs']))) * 100, 1)
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # 2. Successfully Processed Files Sheet
                if self.detailed_results['successful_matches']:
                    processed_data = []
                    for item in self.detailed_results['successful_matches']:
                        processed_data.append({
                            'Flat Reference': item['flat_ref'],
                            'Original Filename': item['original_filename'],
                            'New Filename': item['new_filename'],
                            'Title': item['title'],
                            'Source Path': item['source_path'],
                            'Destination Path': item['destination_path'],
                            'Destination Folder': item['destination_folder']
                        })
                    processed_df = pd.DataFrame(processed_data)
                    processed_df.to_excel(writer, sheet_name='Successfully Processed', index=False)
                else:
                    # Create empty sheet with headers
                    empty_df = pd.DataFrame(columns=['Flat Reference', 'Original Filename', 'New Filename', 'Title', 'Source Path', 'Destination Path', 'Destination Folder'])
                    empty_df.to_excel(writer, sheet_name='Successfully Processed', index=False)
                
                # 3. No Matches Found Sheet
                if self.detailed_results['no_matches_found']:
                    no_match_data = []
                    for item in self.detailed_results['no_matches_found']:
                        no_match_data.append({
                            'Flat Reference': item['flat_ref'],
                            'Reason': item['reason']
                        })
                    no_match_df = pd.DataFrame(no_match_data)
                    no_match_df.to_excel(writer, sheet_name='No Matches Found', index=False)
                else:
                    empty_df = pd.DataFrame(columns=['Flat Reference', 'Reason'])
                    empty_df.to_excel(writer, sheet_name='No Matches Found', index=False)
                
                # 4. Files Not Found Sheet
                if self.detailed_results['file_not_found']:
                    not_found_data = []
                    for item in self.detailed_results['file_not_found']:
                        not_found_data.append({
                            'Flat Reference': item['flat_ref'],
                            'Title': item['title'],
                            'Filename': item['filename'],
                            'Expected Path': item['expected_path'],
                            'Reason': item['reason']
                        })
                    not_found_df = pd.DataFrame(not_found_data)
                    not_found_df.to_excel(writer, sheet_name='Files Not Found', index=False)
                else:
                    empty_df = pd.DataFrame(columns=['Flat Reference', 'Title', 'Filename', 'Expected Path', 'Reason'])
                    empty_df.to_excel(writer, sheet_name='Files Not Found', index=False)
                
                # 5. Copy Errors Sheet
                if self.detailed_results['copy_errors']:
                    error_data = []
                    for item in self.detailed_results['copy_errors']:
                        error_data.append({
                            'Flat Reference': item['flat_ref'],
                            'Title': item['title'],
                            'Filename': item['filename'],
                            'Source Path': item['source_path'],
                            'Error': item['error']
                        })
                    error_df = pd.DataFrame(error_data)
                    error_df.to_excel(writer, sheet_name='Copy Errors', index=False)
                else:
                    empty_df = pd.DataFrame(columns=['Flat Reference', 'Title', 'Filename', 'Source Path', 'Error'])
                    empty_df.to_excel(writer, sheet_name='Copy Errors', index=False)
                
                # 6. Unused Section Files Sheet
                if self.detailed_results['unused_section_files']:
                    unused_data = []
                    for item in self.detailed_results['unused_section_files']:
                        unused_data.append({
                            'Filename': item['filename'],
                            'Title': item['title'],
                            'Status': 'UNUSED'
                        })
                    unused_df = pd.DataFrame(unused_data)
                    unused_df.to_excel(writer, sheet_name='Unused Section Files', index=False)
                else:
                    empty_df = pd.DataFrame(columns=['Filename', 'Title', 'Status'])
                    empty_df.to_excel(writer, sheet_name='Unused Section Files', index=False)
                
                # 7. All Flat References Sheet
                flat_ref_data = []
                for i, flat_ref in enumerate(self.detailed_results['all_flat_refs'], 1):
                    processed = any(item['flat_ref'] == flat_ref for item in self.detailed_results['successful_matches'])
                    files_count = sum(1 for item in self.detailed_results['successful_matches'] if item['flat_ref'] == flat_ref)
                    
                    flat_ref_data.append({
                        'No.': i,
                        'Flat Reference': flat_ref,
                        'Status': 'PROCESSED' if processed else 'NOT PROCESSED',
                        'Files Copied': files_count if processed else 0
                    })
                flat_ref_df = pd.DataFrame(flat_ref_data)
                flat_ref_df.to_excel(writer, sheet_name='All Flat References', index=False)
                
                # 8. All Section Files Sheet
                section_data = []
                for i, section_file in enumerate(self.detailed_results['all_section_files'], 1):
                    section_data.append({
                        'No.': i,
                        'Filename': section_file['filename'],
                        'Title': section_file['title'],
                        'Status': 'USED' if section_file['matched'] else 'UNUSED'
                    })
                section_df = pd.DataFrame(section_data)
                section_df.to_excel(writer, sheet_name='All Section Files', index=False)
                
                # 9. Processing Summary by Flat Reference
                by_flat_ref = {}
                for item in self.detailed_results['successful_matches']:
                    flat_ref = item['flat_ref']
                    if flat_ref not in by_flat_ref:
                        by_flat_ref[flat_ref] = []
                    by_flat_ref[flat_ref].append(item)
                
                summary_by_ref = []
                for flat_ref in self.detailed_results['all_flat_refs']:
                    files = by_flat_ref.get(flat_ref, [])
                    summary_by_ref.append({
                        'Flat Reference': flat_ref,
                        'Files Copied': len(files),
                        'Status': 'SUCCESS' if files else 'NO MATCH',
                        'Output Folder': files[0]['destination_folder'] if files else 'N/A'
                    })
                
                summary_by_ref_df = pd.DataFrame(summary_by_ref)
                summary_by_ref_df.to_excel(writer, sheet_name='Summary by Flat Ref', index=False)
                
                # Add metadata sheet
                metadata = {
                    'Report Information': [
                        'Generated Date/Time',
                        'Script Version',
                        'Document Type Processed',
                        'Source Spreadsheets'
                    ],
                    'Value': [
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'File Processor v2.0',
                        'Sections',
                        'accomodation_schedule.xlsx, architect_spreadsheet.xlsx'
                    ]
                }
                metadata_df = pd.DataFrame(metadata)
                metadata_df.to_excel(writer, sheet_name='Report Info', index=False)
            
            logging.info(f"Detailed Excel report generated: {report_path}")
            return True
            
        except Exception as e:
            logging.error(f"Error generating detailed report: {str(e)}")
            return False


def main():
    """Main function to run the file processor"""
    processor = FileProcessor()
    
    try:
        success = processor.process_files()
        if success:
            logging.info("Processing completed successfully!")
        else:
            logging.error("Processing failed!")
            
    except KeyboardInterrupt:
        logging.info("Processing interrupted by user")
    except Exception as e:
        logging.error(f"Unexpected error: {str(e)}")


if __name__ == "__main__":
    main()