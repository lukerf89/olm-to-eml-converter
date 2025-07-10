# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
This is an OLM to EML converter - a Python utility that converts Outlook for Mac archive files (.olm) to individual email files (.eml) for easier processing by AI tools.

## Commands
- **Run the converter**: `python olm_to_eml_converter.py path/to/file.olm output_directory`
- **Test with sample**: `python olm_to_eml_converter.py ~/Documents/MyEmails.olm ./converted_emails`

## Architecture
This is a single-file Python script with no external dependencies. The main architecture consists of:

- **OLMToEMLConverter class**: Core conversion logic that treats OLM files as ZIP archives
- **Two-stage processing**: First extracts the ZIP contents, then processes message files
- **Dual format support**: Handles both XML and binary message formats (.olk15Message, .olk14Message)
- **Standard library only**: Uses zipfile, xml.etree.ElementTree, email, and pathlib

## Key Technical Details
- **Input**: .olm files (Outlook for Mac archives, actually ZIP files)
- **Output**: Individual .eml files with standard email headers
- **Python version**: 3.6+
- **Dependencies**: None (standard library only)

## Core Methods
- `convert()` - Main conversion orchestrator
- `_extract_olm()` - Extracts OLM ZIP contents to temp directory
- `_process_messages()` - Finds and processes .olk15Message/.olk14Message files
- `_convert_message_to_eml()` - Converts individual messages to EML format
- `_extract_email_from_xml()` - Handles XML format messages
- `_extract_email_from_binary()` - Handles binary format messages

## Known Limitations
- OLM format is proprietary and complex
- Some message formats may not be fully supported
- Large files may require significant processing time
- Character encoding issues with special characters