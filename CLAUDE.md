# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
This is a three-part email processing utility: converts Outlook for Mac archive files (.olm) to individual email files (.eml), then converts those to CSV database format, and finally chunks the CSV into monthly files for easier AI analysis.

## Commands
- **Convert OLM to EML**: `python3 olm_to_eml_converter.py path/to/file.olm output_directory`
- **Convert EML to CSV**: `python3 eml_to_csv_converter.py eml_directory output.csv`
- **Chunk CSV by month**: `python3 csv_chunker.py input.csv output_directory [--prefix emails]`
- **Full workflow example**: 
  ```bash
  python3 olm_to_eml_converter.py "~/Documents/EmailArchive.olm" ./converted_emails
  python3 eml_to_csv_converter.py ./converted_emails ./emails_database.csv
  python3 csv_chunker.py ./emails_database.csv ./monthly_emails
  ```
- **Check script help**: Each script supports `--help` flag for detailed usage information

## Architecture
Three complementary Python scripts with no external dependencies:

### OLM to EML Converter
- **OLMToEMLConverter class**: Treats OLM files as ZIP archives
- **Dual directory support**: Searches both `Local/` and `Accounts/` directories
- **Multiple format support**: Handles .olk15Message, .olk14Message, and message_*.xml files
- **Outlook XML parsing**: Correctly parses OPFMessageCopy* XML elements from Outlook for Mac
- **Two-stage processing**: Extract ZIP contents, then process message files

### EML to CSV Converter  
- **EMLToCSVConverter class**: Parses EML files and extracts structured data
- **16-column output**: Includes subject parsing, thread grouping, and summary fields
- **Notion AI optimized**: CSV structure with summary_input field for AI analysis
- **Advanced features**: Subject prefix parsing, thread ID generation, truncation indicators
- **Robust parsing**: Handles both multipart and single-part messages

### CSV Chunker
- **EmailCSVChunker class**: Splits large CSV files into monthly chunks
- **Date-based grouping**: Organizes emails by year and month from date_parsed field
- **Flexible naming**: Configurable filename prefix (default: emails_YYYY_MM.csv)
- **Import-friendly**: Creates manageable file sizes for Notion AI processing
- **Summary reporting**: Shows email counts per month and files created

## Key Technical Details
- **Input**: .olm files (Outlook for Mac archives) → .eml files → .csv database → monthly .csv files
- **Python version**: 3.6+ (tested with 3.12.1)
- **Dependencies**: None (standard library only)
- **Libraries used**: zipfile, xml.etree.ElementTree, email, pathlib, csv, datetime, collections, argparse, tempfile
- **Error handling**: All scripts include try/except blocks for robust file processing
- **Command-line interface**: All scripts use argparse for consistent CLI experience

## Core Methods

### OLM Converter
- `convert()` - Main conversion orchestrator
- `_extract_olm()` - Extracts OLM ZIP contents to temp directory
- `_process_messages()` - Finds and processes message files in Local/Accounts directories
- `_extract_email_from_xml()` - Parses Outlook XML with OPFMessageCopy* elements
- `_convert_message_to_eml()` - Converts individual messages to EML format

### CSV Converter
- `convert()` - Main CSV conversion orchestrator
- `_parse_eml_file()` - Extracts structured data from individual EML files
- `_parse_subject_prefix()` - Separates RE/FWD prefixes from subject
- `_generate_thread_id()` - Creates thread grouping from message references
- `_create_summary_input()` - Builds combined metadata block for AI analysis
- `_extract_body()` - Handles both text and HTML body content

### CSV Chunker
- `chunk_by_month()` - Main chunking orchestrator
- `_extract_year_month()` - Parses various date formats to extract YYYY_MM
- `_write_monthly_csv()` - Creates individual CSV files for each month
- `_print_summary()` - Displays breakdown of emails per month

## CSV Output Structure (16 columns)
1. **subject** - Clean subject without prefixes
2. **subject_prefix** - RE, FWD, etc. 
3. **from_name/from_email** - Sender information
4. **to_name/to_email** - Recipient information
5. **date_parsed** - ISO format timestamp
6. **body_text** - Plain text with truncation indicator
7. **summary_input** - Combined metadata + body for Notion AI
8. **tags** - Empty for manual/AI categorization
9. **attachments** - Semicolon-separated filenames
10. **thread_id** - Groups related messages
11. **message_id/in_reply_to/references** - Threading metadata
12. **filename** - Original EML filename

## Known Limitations
- OLM format is proprietary and complex
- Large files may require significant processing time
- Character encoding issues with special characters

## Data Flow and Processing Pipeline
1. **OLM Extraction**: Treats OLM as ZIP, extracts to temp directory, processes both Local/ and Accounts/ subdirectories
2. **Message Discovery**: Searches for .olk15Message, .olk14Message, and message_*.xml files recursively
3. **XML Parsing**: Extracts email data from OPFMessageCopy* elements with namespace handling
4. **EML Generation**: Creates RFC-compliant EML files with proper headers and body encoding
5. **CSV Conversion**: Parses EML files, extracts 16 structured fields, handles both text and HTML bodies
6. **Monthly Chunking**: Groups by date_parsed field, creates year_month.csv files for batch processing

## Important File Patterns
- **OLM message files**: `*.olk15Message`, `*.olk14Message`, `message_*.xml`
- **Mac resource forks**: Files starting with `._` are automatically excluded
- **Output naming**: 
  - EML: `message_00001.eml` (5-digit zero-padded)
  - CSV chunks: `emails_YYYY_MM.csv` (customizable prefix)

## Production Notes
- CSV converter automatically excludes Mac resource fork files (._filename.eml)
- Email addresses are extracted from XML attributes, not text content
- HTML body content is converted to plain text with basic cleaning
- Thread IDs use message references to group conversation chains
- All output is UTF-8 encoded for maximum compatibility
- Temporary directories are automatically cleaned up after processing
- Large OLM files are processed in memory-efficient streaming fashion