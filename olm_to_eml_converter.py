#!/usr/bin/env python3
"""
OLM to EML Converter

This script converts Outlook for Mac (.olm) archive files to individual .eml files.
OLM files are essentially ZIP archives containing email data in a structured format.
"""

import os
import zipfile
import xml.etree.ElementTree as ET
import email
import email.utils
import re
from pathlib import Path
import base64
import tempfile
import shutil
from datetime import datetime

class OLMToEMLConverter:
    def __init__(self, olm_file_path, output_dir):
        self.olm_file_path = olm_file_path
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
    def convert(self):
        """Convert OLM file to EML files"""
        print(f"Converting {self.olm_file_path} to EML files...")
        
        # Create temporary directory for extraction
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            # Extract OLM file (it's a ZIP archive)
            self._extract_olm(temp_path)
            
            # Find and process message files
            self._process_messages(temp_path)
            
        print(f"Conversion complete! EML files saved to: {self.output_dir}")
    
    def _extract_olm(self, temp_path):
        """Extract OLM file contents"""
        try:
            with zipfile.ZipFile(self.olm_file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_path)
                print(f"Extracted OLM contents to temporary directory")
        except zipfile.BadZipFile:
            raise ValueError(f"Invalid OLM file: {self.olm_file_path}")
    
    def _process_messages(self, temp_path):
        """Process extracted message files"""
        message_count = 0
        
        # Look for message files in both Local and Accounts directories
        search_dirs = []
        local_dir = temp_path / "Local"
        accounts_dir = temp_path / "Accounts"
        
        if local_dir.exists():
            search_dirs.append(local_dir)
        if accounts_dir.exists():
            search_dirs.append(accounts_dir)
        
        if not search_dirs:
            print("No Local or Accounts directory found in OLM file")
            return
        
        # Find all message files
        for search_dir in search_dirs:
            for root, dirs, files in os.walk(search_dir):
                for file in files:
                    if (file.endswith('.olk15Message') or 
                        file.endswith('.olk14Message') or 
                        file.startswith('message_') and file.endswith('.xml')):
                        message_path = Path(root) / file
                        try:
                            self._convert_message_to_eml(message_path, message_count)
                            message_count += 1
                        except Exception as e:
                            print(f"Error processing {message_path}: {e}")
        
        print(f"Processed {message_count} messages")
    
    def _convert_message_to_eml(self, message_path, message_count):
        """Convert a single message file to EML format"""
        try:
            # Read the message file
            with open(message_path, 'rb') as f:
                content = f.read()
            
            # Try to parse as XML first
            try:
                # Parse the message XML
                root = ET.fromstring(content)
                eml_content = self._extract_email_from_xml(root)
            except ET.ParseError:
                # If XML parsing fails, try to extract email content directly
                eml_content = self._extract_email_from_binary(content)
            
            if eml_content:
                # Save as EML file
                eml_filename = f"message_{message_count:05d}.eml"
                eml_path = self.output_dir / eml_filename
                
                with open(eml_path, 'w', encoding='utf-8') as f:
                    f.write(eml_content)
                
                print(f"Converted: {eml_filename}")
        
        except Exception as e:
            print(f"Error converting {message_path}: {e}")
    
    def _extract_email_from_xml(self, root):
        """Extract email content from Outlook XML structure"""
        import html
        
        subject = ""
        sender = ""
        recipient = ""
        date = ""
        body = ""
        message_id = ""
        
        # Parse Outlook-specific XML elements
        for elem in root.iter():
            if elem.tag == 'OPFMessageCopySubject':
                subject = elem.text or ""
            elif elem.tag == 'OPFMessageCopyDisplayTo':
                recipient = elem.text or ""
            elif elem.tag == 'OPFMessageCopyFromAddresses':
                sender = elem.text or ""
            elif elem.tag == 'OPFMessageCopySentTime':
                date = elem.text or ""
            elif elem.tag == 'OPFMessageCopyBody':
                # This contains HTML-encoded content
                if elem.text:
                    body = html.unescape(elem.text)
            elif elem.tag == 'OPFMessageCopyHTMLBody':
                # Alternative HTML body location
                if elem.text and not body:
                    body = html.unescape(elem.text)
            elif elem.tag == 'OPFMessageCopyMessageID':
                message_id = elem.text or ""
        
        # If we didn't get sender from FromAddresses, try other sources
        if not sender:
            for elem in root.iter():
                if elem.tag == 'OPFMessageCopySenderAddress':
                    sender = elem.text or ""
                    break
        
        # Clean up the body content - convert HTML to plain text if needed
        if body and body.startswith('<'):
            # Simple HTML to text conversion
            body = re.sub(r'<[^>]+>', '', body)
            body = html.unescape(body)
            body = re.sub(r'\s+', ' ', body).strip()
        
        # Format date if it's in ISO format
        if date and 'T' in date:
            try:
                from datetime import datetime
                dt = datetime.fromisoformat(date.replace('Z', '+00:00'))
                date = dt.strftime('%a, %d %b %Y %H:%M:%S %z')
            except:
                pass
        
        # Create EML format with proper headers
        eml_content = f"""From: {sender}
To: {recipient}
Subject: {subject}
Date: {date}
Message-ID: {message_id}

{body}
"""
        return eml_content
    
    def _extract_email_from_binary(self, content):
        """Extract email content from binary data"""
        try:
            # Try to decode as UTF-8
            text_content = content.decode('utf-8', errors='ignore')
            
            # Look for email-like patterns
            subject_match = re.search(r'Subject:\s*([^\n\r]+)', text_content, re.IGNORECASE)
            from_match = re.search(r'From:\s*([^\n\r]+)', text_content, re.IGNORECASE)
            to_match = re.search(r'To:\s*([^\n\r]+)', text_content, re.IGNORECASE)
            date_match = re.search(r'Date:\s*([^\n\r]+)', text_content, re.IGNORECASE)
            
            subject = subject_match.group(1) if subject_match else "No Subject"
            sender = from_match.group(1) if from_match else "Unknown Sender"
            recipient = to_match.group(1) if to_match else "Unknown Recipient"
            date = date_match.group(1) if date_match else datetime.now().strftime("%a, %d %b %Y %H:%M:%S %z")
            
            # Try to extract body content
            # Look for content after headers
            body_start = text_content.find('\n\n')
            if body_start != -1:
                body = text_content[body_start + 2:].strip()
            else:
                body = "Content could not be extracted"
            
            # Create EML format
            eml_content = f"""From: {sender}
To: {recipient}
Subject: {subject}
Date: {date}

{body}
"""
            return eml_content
            
        except Exception as e:
            print(f"Error extracting from binary: {e}")
            return None

def main():
    """Main function to run the converter"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Convert OLM files to EML format')
    parser.add_argument('olm_file', help='Path to the OLM file')
    parser.add_argument('output_dir', help='Output directory for EML files')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.olm_file):
        print(f"Error: OLM file not found: {args.olm_file}")
        return
    
    converter = OLMToEMLConverter(args.olm_file, args.output_dir)
    converter.convert()

if __name__ == "__main__":
    main()