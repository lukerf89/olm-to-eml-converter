# Project Context for Claude

## Project Overview
This is an OLM to EML converter - a Python utility that converts Outlook for Mac archive files (.olm) to individual email files (.eml) for easier processing by AI tools.

## Key Files
- `olm_to_eml_converter.py` - Main conversion script
- `README.md` - User documentation and usage instructions

## Project Structure
```
eml_converter/
├── olm_to_eml_converter.py    # Main conversion script
├── README.md                  # Documentation
└── CLAUDE.md                 # This file
```

## Technical Details
- **Language**: Python 3.6+
- **Dependencies**: None (uses only standard library)
- **Input**: .olm files (Outlook for Mac archives)
- **Output**: Individual .eml files

## Key Components

### OLMToEMLConverter Class
Main class that handles the conversion process:
- Extracts OLM files (ZIP archives)
- Processes message files (.olk15Message, .olk14Message)
- Converts to standard EML format
- Handles both XML and binary message formats

### Core Methods
- `convert()` - Main conversion orchestrator
- `_extract_olm()` - Extracts OLM ZIP contents
- `_process_messages()` - Finds and processes message files
- `_convert_message_to_eml()` - Converts individual messages
- `_extract_email_from_xml()` - Extracts from XML format
- `_extract_email_from_binary()` - Extracts from binary format

## Usage Pattern
```bash
python olm_to_eml_converter.py input.olm output_directory
```

## Common Tasks
- **Testing**: Run with sample OLM files
- **Debugging**: Check extraction and conversion processes
- **Enhancement**: Add support for additional message formats
- **Documentation**: Update README with new features

## Known Limitations
- OLM format is proprietary and complex
- Some message formats may not be fully supported
- Large files may require significant processing time
- Character encoding issues with special characters

## Future Enhancements
- Support for more message formats
- Better error handling and logging
- Progress indicators for large files
- Batch processing capabilities
- Integration with email analysis tools

## Testing Notes
- Requires actual OLM files for testing
- Test with various Outlook for Mac versions
- Verify EML output compatibility with email clients
- Check handling of attachments and complex formatting