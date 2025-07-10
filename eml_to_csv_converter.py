#!/usr/bin/env python3
"""
EML to CSV Converter

This script converts EML files to a CSV database format suitable for Notion AI.
It extracts key email fields and creates a structured CSV file.
"""

import os
import csv
import email
import email.utils
from pathlib import Path
from datetime import datetime
import re

class EMLToCSVConverter:
    def __init__(self, eml_directory, csv_output_path):
        self.eml_directory = Path(eml_directory)
        self.csv_output_path = Path(csv_output_path)
        
    def convert(self):
        """Convert all EML files in directory to CSV"""
        print(f"Converting EML files from {self.eml_directory} to CSV...")
        
        # Get all EML files
        eml_files = list(self.eml_directory.glob("*.eml"))
        if not eml_files:
            print("No EML files found in directory")
            return
        
        print(f"Found {len(eml_files)} EML files")
        
        # Create CSV with headers
        with open(self.csv_output_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'filename',
                'subject',
                'from_name',
                'from_email',
                'to_name',
                'to_email',
                'date',
                'date_parsed',
                'body_text',
                'body_html',
                'attachments',
                'message_id',
                'in_reply_to',
                'references'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            # Process each EML file
            processed_count = 0
            for eml_file in eml_files:
                try:
                    email_data = self._parse_eml_file(eml_file)
                    if email_data:
                        writer.writerow(email_data)
                        processed_count += 1
                        if processed_count % 100 == 0:
                            print(f"Processed {processed_count} emails...")
                except Exception as e:
                    print(f"Error processing {eml_file}: {e}")
            
            print(f"Successfully converted {processed_count} emails to CSV")
    
    def _parse_eml_file(self, eml_file):
        """Parse a single EML file and extract email data"""
        try:
            with open(eml_file, 'r', encoding='utf-8', errors='ignore') as f:
                msg = email.message_from_file(f)
        except Exception as e:
            print(f"Error reading {eml_file}: {e}")
            return None
        
        # Extract basic fields
        subject = self._clean_text(msg.get('Subject', ''))
        from_field = msg.get('From', '')
        to_field = msg.get('To', '')
        date_field = msg.get('Date', '')
        message_id = msg.get('Message-ID', '')
        in_reply_to = msg.get('In-Reply-To', '')
        references = msg.get('References', '')
        
        # Parse From field
        from_name, from_email = self._parse_email_address(from_field)
        
        # Parse To field (take first recipient if multiple)
        to_name, to_email = self._parse_email_address(to_field)
        
        # Parse date
        date_parsed = self._parse_date(date_field)
        
        # Extract body content
        body_text, body_html = self._extract_body(msg)
        
        # Get attachments info
        attachments = self._get_attachments_info(msg)
        
        return {
            'filename': eml_file.name,
            'subject': subject,
            'from_name': from_name,
            'from_email': from_email,
            'to_name': to_name,
            'to_email': to_email,
            'date': date_field,
            'date_parsed': date_parsed,
            'body_text': self._clean_text(body_text),
            'body_html': self._clean_text(body_html),
            'attachments': attachments,
            'message_id': message_id,
            'in_reply_to': in_reply_to,
            'references': references
        }
    
    def _parse_email_address(self, email_field):
        """Parse email address field into name and email"""
        if not email_field:
            return '', ''
        
        try:
            # Handle multiple addresses - take the first one
            addresses = email.utils.getaddresses([email_field])
            if addresses:
                name, email_addr = addresses[0]
                return name.strip(), email_addr.strip()
        except Exception:
            pass
        
        # Fallback: try to extract email with regex
        email_match = re.search(r'[\w\.-]+@[\w\.-]+', email_field)
        if email_match:
            return '', email_match.group()
        
        return email_field.strip(), ''
    
    def _parse_date(self, date_field):
        """Parse date field into ISO format"""
        if not date_field:
            return ''
        
        try:
            # Parse the date
            date_tuple = email.utils.parsedate_tz(date_field)
            if date_tuple:
                timestamp = email.utils.mktime_tz(date_tuple)
                return datetime.fromtimestamp(timestamp).isoformat()
        except Exception:
            pass
        
        return date_field
    
    def _extract_body(self, msg):
        """Extract text and HTML body from email message"""
        body_text = ''
        body_html = ''
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == 'text/plain':
                    try:
                        body_text = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except Exception:
                        pass
                elif content_type == 'text/html':
                    try:
                        body_html = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except Exception:
                        pass
        else:
            # Single part message
            content_type = msg.get_content_type()
            try:
                payload = msg.get_payload(decode=True)
                if payload:
                    content = payload.decode('utf-8', errors='ignore')
                    if content_type == 'text/html':
                        body_html = content
                    else:
                        body_text = content
            except Exception:
                pass
        
        return body_text, body_html
    
    def _get_attachments_info(self, msg):
        """Get information about attachments"""
        attachments = []
        
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        attachments.append(filename)
        
        return '; '.join(attachments) if attachments else ''
    
    def _clean_text(self, text):
        """Clean text for CSV output"""
        if not text:
            return ''
        
        # Remove excessive whitespace and newlines
        text = re.sub(r'\s+', ' ', text)
        text = text.strip()
        
        # Limit length for CSV readability
        if len(text) > 5000:
            text = text[:5000] + '...'
        
        return text

def main():
    """Main function for command line usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Convert EML files to CSV format')
    parser.add_argument('eml_directory', help='Directory containing EML files')
    parser.add_argument('csv_output', help='Output CSV file path')
    
    args = parser.parse_args()
    
    converter = EMLToCSVConverter(args.eml_directory, args.csv_output)
    converter.convert()

if __name__ == "__main__":
    main()