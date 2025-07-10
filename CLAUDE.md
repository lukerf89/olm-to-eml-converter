# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
This is a two-part email processing utility: converts Outlook for Mac archive files (.olm) to individual email files (.eml), then converts those to CSV database format for AI analysis.

## Commands
- **Convert OLM to EML**: `python olm_to_eml_converter.py path/to/file.olm output_directory`
- **Convert EML to CSV**: `python eml_to_csv_converter.py eml_directory output.csv`
- **Full workflow example**: 
  ```bash
  python olm_to_eml_converter.py "~/Documents/EmailArchive.olm" ./converted_emails
  python eml_to_csv_converter.py ./converted_emails ./emails_database.csv
  ```

## Architecture
Two complementary Python scripts with no external dependencies:

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

## Key Technical Details
- **Input**: .olm files (Outlook for Mac archives) → .eml files → .csv database
- **Python version**: 3.6+
- **Dependencies**: None (standard library only)
- **Libraries used**: zipfile, xml.etree.ElementTree, email, pathlib, csv

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