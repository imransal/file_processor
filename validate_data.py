#!/usr/bin/env python3
"""
Test and Validation Script for File Processing

This script helps you examine your spreadsheets and test the matching logic
before running the full file processing script.

Author: GitHub Copilot
Date: October 2025
"""

import pandas as pd
from pathlib import Path
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def examine_spreadsheets():
    """Examine the structure and content of both spreadsheets"""
    base_path = Path(r"C:\Users\IS19\Documents\as_built_extraction")
    spreadsheet_path = base_path / "spreadsheet"
    
    # Load accommodation schedule
    print("=" * 60)
    print("EXAMINING ACCOMMODATION SCHEDULE")
    print("=" * 60)
    
    try:
        accommodation_file = spreadsheet_path / "accomodation_schedule.xlsx"
        df_acc = pd.read_excel(accommodation_file)
        
        print(f"Shape: {df_acc.shape}")
        print(f"Columns: {list(df_acc.columns)}")
        print("\nFirst 5 rows:")
        print(df_acc.head())
        
        # Show Column D specifically
        print(f"\nColumn D (index 3) - 'Flat / House Ref' examples:")
        if len(df_acc.columns) > 3:
            col_d_values = df_acc.iloc[:, 3].dropna().head(10)
            for i, value in enumerate(col_d_values):
                print(f"  {i+1}: {value}")
        else:
            print("  Column D not found - check column structure")
            
    except Exception as e:
        print(f"Error loading accommodation schedule: {e}")
    
    # Load architect spreadsheet
    print("\n" + "=" * 60)
    print("EXAMINING ARCHITECT SPREADSHEET")
    print("=" * 60)
    
    try:
        architect_file = spreadsheet_path / "architect_spreadsheet.xlsx"
        df_arch = pd.read_excel(architect_file)
        
        print(f"Shape: {df_arch.shape}")
        print(f"Columns: {list(df_arch.columns)}")
        print("\nFirst 5 rows:")
        print(df_arch.head())
        
        # Show Column A (Filename) and B (Title) specifically
        print(f"\nColumn A - 'Filename' examples:")
        if len(df_arch.columns) > 0:
            col_a_values = df_arch.iloc[:, 0].dropna().head(10)
            for i, value in enumerate(col_a_values):
                print(f"  {i+1}: {value}")
        
        print(f"\nColumn B - 'Title' examples:")
        if len(df_arch.columns) > 1:
            col_b_values = df_arch.iloc[:, 1].dropna().head(10)
            for i, value in enumerate(col_b_values):
                print(f"  {i+1}: {value}")
                
        # Look for "Sections" entries specifically
        print(f"\nEntries containing 'Sections':")
        if len(df_arch.columns) > 1:
            sections_mask = df_arch.iloc[:, 1].str.contains("Sections", case=False, na=False)
            sections_entries = df_arch[sections_mask].head(10)
            for idx, row in sections_entries.iterrows():
                filename = row.iloc[0] if pd.notna(row.iloc[0]) else "N/A"
                title = row.iloc[1] if pd.notna(row.iloc[1]) else "N/A"
                print(f"  {filename} -> {title}")
                
    except Exception as e:
        print(f"Error loading architect spreadsheet: {e}")


def test_matching_logic():
    """Test the matching logic with sample data"""
    print("\n" + "=" * 60)
    print("TESTING MATCHING LOGIC")
    print("=" * 60)
    
    base_path = Path(r"C:\Users\IS19\Documents\as_built_extraction")
    spreadsheet_path = base_path / "spreadsheet"
    
    try:
        # Load both files
        accommodation_file = spreadsheet_path / "accomodation_schedule.xlsx"
        architect_file = spreadsheet_path / "architect_spreadsheet.xlsx"
        
        df_acc = pd.read_excel(accommodation_file)
        df_arch = pd.read_excel(architect_file)
        
        # Get flat references (Column D)
        if len(df_acc.columns) > 3:
            flat_refs = df_acc.iloc[:, 3].dropna().unique()[:5]  # Test with first 5
            print(f"Testing with flat references: {list(flat_refs)}")
        else:
            print("Cannot access Column D in accommodation schedule")
            return
        
        # Test matching
        if len(df_arch.columns) > 1:
            titles = df_arch.iloc[:, 1]  # Column B
            filenames = df_arch.iloc[:, 0]  # Column A
            
            for flat_ref in flat_refs:
                print(f"\nLooking for matches with: '{flat_ref}'")
                target_pattern = f"Sections - {flat_ref}"
                print(f"  Target pattern: '{target_pattern}'")
                
                # Find matches
                matching_rows = titles.str.contains(target_pattern, case=False, na=False)
                
                if matching_rows.any():
                    print(f"  ✓ Found {matching_rows.sum()} match(es):")
                    for idx in matching_rows[matching_rows].index:
                        filename = filenames.iloc[idx]
                        title = titles.iloc[idx]
                        print(f"    - {filename} -> {title}")
                else:
                    print(f"  ✗ No matches found")
                    
                    # Try to find partial matches for debugging
                    partial_matches = titles.str.contains(flat_ref, case=False, na=False)
                    if partial_matches.any():
                        print(f"    But found partial matches:")
                        for idx in partial_matches[partial_matches].index[:3]:  # Show first 3
                            title = titles.iloc[idx]
                            print(f"      - {title}")
        
    except Exception as e:
        print(f"Error in matching logic test: {e}")


def check_file_availability():
    """Check if files mentioned in architect spreadsheet actually exist"""
    print("\n" + "=" * 60)
    print("CHECKING FILE AVAILABILITY")
    print("=" * 60)
    
    base_path = Path(r"C:\Users\IS19\Documents\as_built_extraction")
    spreadsheet_path = base_path / "spreadsheet"
    architect_path = base_path / "architect"
    
    try:
        architect_file = spreadsheet_path / "architect_spreadsheet.xlsx"
        df_arch = pd.read_excel(architect_file)
        
        if len(df_arch.columns) > 0:
            filenames = df_arch.iloc[:, 0].dropna().unique()
            
            print(f"Checking {len(filenames)} unique filenames...")
            
            found_count = 0
            missing_count = 0
            
            for filename in filenames[:20]:  # Check first 20 files
                file_path = architect_path / filename
                if file_path.exists():
                    found_count += 1
                    print(f"  ✓ Found: {filename}")
                else:
                    missing_count += 1
                    print(f"  ✗ Missing: {filename}")
            
            print(f"\nSummary (first 20 files):")
            print(f"  Found: {found_count}")
            print(f"  Missing: {missing_count}")
            
            if missing_count > 0:
                print(f"\nNote: Some files may have different naming conventions.")
                print(f"You may need to adjust the matching logic.")
        
    except Exception as e:
        print(f"Error checking file availability: {e}")


def main():
    """Main function to run all tests"""
    print("FILE PROCESSING VALIDATION SCRIPT")
    print("This script will help you understand your data before processing")
    
    examine_spreadsheets()
    test_matching_logic()
    check_file_availability()
    
    print(f"\n" + "=" * 60)
    print("VALIDATION COMPLETE")
    print("=" * 60)
    print("Review the output above to ensure the matching logic is working correctly.")
    print("If everything looks good, you can run the main file_processor.py script.")


if __name__ == "__main__":
    main()