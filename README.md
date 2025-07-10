# OLM to EML Converter

A Python script to convert Outlook for Mac (.olm) archive files to individual .eml files for easy reading by AI tools like Notion AI.

## Features

- Extracts emails from OLM (Outlook for Mac) archive files
- Converts each email to standard EML format
- Handles both XML and binary message formats
- Preserves email headers (From, To, Subject, Date)
- Extracts email body content

## Requirements

- Python 3.6+
- No external dependencies (uses only standard library)

## Usage

### Command Line

```bash
python olm_to_eml_converter.py path/to/your/file.olm output_directory
```

### Example

```bash
python olm_to_eml_converter.py ~/Documents/MyEmails.olm ./converted_emails
```

### Programmatic Usage

```python
from olm_to_eml_converter import OLMToEMLConverter

converter = OLMToEMLConverter("path/to/file.olm", "output_directory")
converter.convert()
```

## Output

The script creates individual .eml files with names like:
- `message_00001.eml`
- `message_00002.eml`
- etc.

Each .eml file contains:
- Standard email headers (From, To, Subject, Date)
- Email body content
- Compatible with email clients and AI tools

## How It Works

1. **Extraction**: OLM files are ZIP archives containing email data
2. **Processing**: Looks for message files (`.olk15Message` or `.olk14Message`)
3. **Conversion**: Extracts email content from XML or binary format
4. **Output**: Creates standard EML files for each email

## Notes

- OLM files are specific to Outlook for Mac
- The script handles various message formats within OLM files
- Some complex formatting may be simplified during conversion
- Large OLM files may take time to process

## Troubleshooting

- **"Invalid OLM file"**: Ensure the file is a valid OLM archive
- **"No Local directory found"**: The OLM structure may be different
- **Encoding issues**: Some special characters may not display correctly

## Using with AI Tools

After conversion, you can:
1. Upload the .eml files to Notion AI or similar tools
2. Use the provided CSV conversion script to create a spreadsheet
3. Process emails programmatically for analysis