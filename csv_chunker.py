#!/usr/bin/env python3
"""
CSV Email Chunker

This script takes a large email CSV file and splits it into separate CSV files
organized by year and month. Each output file contains emails for a specific month.
"""

import csv
import os
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import argparse

class EmailCSVChunker:
    def __init__(self, input_csv_path, output_directory, filename_prefix="emails"):
        self.input_csv_path = Path(input_csv_path)
        self.output_directory = Path(output_directory)
        self.filename_prefix = filename_prefix
        
        # Create output directory if it doesn't exist
        self.output_directory.mkdir(exist_ok=True, parents=True)
        
    def chunk_by_month(self):
        """Split CSV file into monthly chunks"""
        print(f"Processing {self.input_csv_path}...")
        
        # Dictionary to store rows grouped by year-month
        monthly_data = defaultdict(list)
        headers = None
        total_rows = 0
        skipped_rows = 0
        
        # Read and group emails by month
        with open(self.input_csv_path, 'r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            headers = reader.fieldnames
            
            for row in reader:
                total_rows += 1
                
                # Extract year-month from date_parsed field
                year_month = self._extract_year_month(row.get('date_parsed', ''))
                
                if year_month:
                    monthly_data[year_month].append(row)
                else:
                    skipped_rows += 1
                    print(f"Skipped row {total_rows}: invalid date '{row.get('date_parsed', '')}'")
        
        print(f"Processed {total_rows} total rows")
        print(f"Skipped {skipped_rows} rows with invalid dates")
        print(f"Found emails for {len(monthly_data)} different months")
        
        # Write separate CSV files for each month
        for year_month, rows in monthly_data.items():
            self._write_monthly_csv(year_month, headers, rows)
        
        print(f"\nCompleted! Created {len(monthly_data)} monthly CSV files in {self.output_directory}")
        
        # Print summary of files created
        self._print_summary(monthly_data)
    
    def _extract_year_month(self, date_string):
        """Extract YYYY_MM format from date_parsed field"""
        if not date_string:
            return None
            
        try:
            # Handle various date formats
            if 'T' in date_string:
                # ISO format: 2024-05-21T13:12:19
                dt = datetime.fromisoformat(date_string.replace('Z', '+00:00'))
            else:
                # Try other common formats
                for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
                    try:
                        dt = datetime.strptime(date_string, fmt)
                        break
                    except ValueError:
                        continue
                else:
                    return None
            
            return f"{dt.year}_{dt.month:02d}"
            
        except (ValueError, TypeError) as e:
            return None
    
    def _write_monthly_csv(self, year_month, headers, rows):
        """Write CSV file for a specific month"""
        filename = f"{self.filename_prefix}_{year_month}.csv"
        filepath = self.output_directory / filename
        
        with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=headers)
            writer.writeheader()
            writer.writerows(rows)
        
        print(f"Created {filename} with {len(rows)} emails")
    
    def _print_summary(self, monthly_data):
        """Print summary of created files"""
        print(f"\n{'Month':<12} {'Count':<8} {'Filename'}")
        print("-" * 50)
        
        # Sort by year-month for better display
        sorted_months = sorted(monthly_data.keys())
        
        for year_month in sorted_months:
            count = len(monthly_data[year_month])
            filename = f"{self.filename_prefix}_{year_month}.csv"
            
            # Format display month
            year, month = year_month.split('_')
            month_name = datetime(int(year), int(month), 1).strftime('%b %Y')
            
            print(f"{month_name:<12} {count:<8} {filename}")

def main():
    """Main function for command line usage"""
    parser = argparse.ArgumentParser(description='Split email CSV into monthly files')
    parser.add_argument('input_csv', help='Path to input CSV file')
    parser.add_argument('output_directory', help='Directory to save monthly CSV files')
    parser.add_argument('--prefix', default='emails', help='Filename prefix (default: emails)')
    
    args = parser.parse_args()
    
    # Validate input file
    if not os.path.exists(args.input_csv):
        print(f"Error: Input CSV file not found: {args.input_csv}")
        return
    
    # Create chunker and process
    chunker = EmailCSVChunker(args.input_csv, args.output_directory, args.prefix)
    chunker.chunk_by_month()

if __name__ == "__main__":
    main()