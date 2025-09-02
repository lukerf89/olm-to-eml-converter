#!/usr/bin/env python3
"""
Invoice Email Processor
Processes .eml files to generate training data for invoice classification AI system.
Creates two CSV files: invoice_classification_data.csv and vendor_patterns.csv
"""

import os
import csv
import email
import email.utils
import re
from pathlib import Path
from collections import defaultdict
from datetime import datetime
import argparse

class InvoiceEmailProcessor:
    def __init__(self, input_directory, output_directory):
        self.input_directory = Path(input_directory)
        self.output_directory = Path(output_directory)
        self.output_directory.mkdir(exist_ok=True, parents=True)
        
        # Classification patterns
        self.invoice_keywords = [
            'invoice', 'bill', 'statement', 'payment due',
            'amount due', 'remit payment', 'balance due', 'amount owed',
            'invoice attached', 'billing', 'payment terms'
        ]
        
        self.shipping_keywords = [
            'shipped', 'tracking', 'delivery', 'carrier',
            'tracking number', 'expected delivery', 'in transit',
            'has been shipped', 'order shipped', 'shipment confirmation',
            'ups', 'fedex', 'usps', 'dhl'
        ]
        
        self.po_keywords = [
            'purchase order', 'po confirmation', 'order confirmed',
            'order accepted', 'po #', 'order acknowledgment',
            'order received', 'order processed', 'confirmation number'
        ]
        
        self.false_positive_keywords = [
            'newsletter', 'promotion', 'sale', 'marketing',
            'unsubscribe', 'catalog', 'new products', 'follow us',
            'discount', 'special offer', 'limited time'
        ]
        
        # Pattern compilations
        self.invoice_subject_patterns = [
            re.compile(r'invoice\s*#?\s*\d+', re.IGNORECASE),
            re.compile(r'bill\s*#?\s*\d+', re.IGNORECASE),
            re.compile(r'statement\s*#?\s*\d+', re.IGNORECASE),
            re.compile(r'inv\s*[-_]?\s*\d+', re.IGNORECASE)
        ]
        
        self.invoice_attachment_patterns = [
            re.compile(r'inv[_\-]?\d+\.pdf', re.IGNORECASE),
            re.compile(r'invoice.*\.pdf', re.IGNORECASE),
            re.compile(r'bill.*\.pdf', re.IGNORECASE),
            re.compile(r'statement.*\.pdf', re.IGNORECASE)
        ]
        
        self.shipping_attachment_patterns = [
            re.compile(r'tracking.*\.pdf', re.IGNORECASE),
            re.compile(r'shipment.*\.pdf', re.IGNORECASE),
            re.compile(r'delivery.*\.pdf', re.IGNORECASE)
        ]
        
        self.po_attachment_patterns = [
            re.compile(r'po[_\-]?\d+\.pdf', re.IGNORECASE),
            re.compile(r'purchase.*order.*\.pdf', re.IGNORECASE),
            re.compile(r'confirmation.*\.pdf', re.IGNORECASE)
        ]
        
        # Storage for processing
        self.processed_emails = []
        self.vendor_data = defaultdict(lambda: {
            'count': 0,
            'domains': set(),
            'subject_patterns': [],
            'attachment_patterns': []
        })
        
    def process_all_emails(self):
        """Process all .eml files in the input directory"""
        all_eml_files = list(self.input_directory.glob("*.eml"))
        # Filter out Mac resource fork files
        eml_files = [f for f in all_eml_files if not f.name.startswith('._')]
        
        if not eml_files:
            print("No .eml files found in input directory")
            return
        
        print(f"Found {len(eml_files)} .eml files to process")
        
        for eml_file in eml_files:
            try:
                self.process_email(eml_file)
            except Exception as e:
                print(f"Error processing {eml_file.name}: {e}")
                continue
        
        # Generate output files
        self.generate_classification_csv()
        self.generate_vendor_patterns_csv()
        
        print(f"\nProcessing complete!")
        print(f"- Processed {len(self.processed_emails)} emails")
        print(f"- Found {len(self.vendor_data)} unique vendors")
        print(f"- Output files saved to {self.output_directory}")
        
    def process_email(self, eml_file):
        """Process a single .eml file"""
        with open(eml_file, 'rb') as f:
            msg = email.message_from_binary_file(f)
        
        # Extract email data
        email_data = self.extract_email_data(msg, eml_file.name)
        
        # Classify email
        email_data['email_type'] = self.classify_email(email_data)
        
        # Sanitize data
        email_data = self.sanitize_email_data(email_data)
        
        # Extract vendor info
        vendor_name = self.extract_vendor_name(email_data)
        if vendor_name:
            self.update_vendor_data(vendor_name, email_data)
        
        # Store processed email
        self.processed_emails.append(email_data)
        
        print(f"Processed: {eml_file.name} - Type: {email_data['email_type']}")
        
    def extract_email_data(self, msg, filename):
        """Extract relevant data from email message"""
        # Get sender info
        from_header = msg.get('From', '')
        from_name, from_email = email.utils.parseaddr(from_header)
        
        # Get subject
        subject = msg.get('Subject', '')
        
        # Get date
        date_str = msg.get('Date', '')
        
        # Extract body
        body = self.extract_body(msg)
        
        # Extract attachments
        attachments = self.extract_attachments(msg)
        
        return {
            'filename': filename,
            'from_name': from_name,
            'from_email': from_email,
            'subject': subject,
            'date': date_str,
            'body': body[:1000],  # Limit body length
            'attachments': attachments,
            'attachment_names': [a['filename'] for a in attachments]
        }
        
    def extract_body(self, msg):
        """Extract email body text"""
        body = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                        break
                    except:
                        continue
                elif part.get_content_type() == "text/html" and not body:
                    try:
                        html = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                        # Simple HTML stripping
                        body = re.sub('<[^<]+?>', '', html)
                    except:
                        continue
        else:
            try:
                body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
            except:
                body = str(msg.get_payload())
        
        return body.strip()
        
    def extract_attachments(self, msg):
        """Extract attachment information"""
        attachments = []
        
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        attachments.append({
                            'filename': filename,
                            'content_type': part.get_content_type()
                        })
        
        return attachments
        
    def classify_email(self, email_data):
        """Classify email based on patterns and keywords"""
        subject = email_data['subject'].lower()
        body = email_data['body'].lower()
        attachment_names = [a.lower() for a in email_data['attachment_names']]
        
        # Score each category
        scores = {
            'INVOICE': 0,
            'SHIPPING': 0,
            'PURCHASE_ORDER': 0,
            'OTHER': 0
        }
        
        # Check attachments (highest confidence)
        for attachment in attachment_names:
            for pattern in self.invoice_attachment_patterns:
                if pattern.search(attachment):
                    scores['INVOICE'] += 3
                    
            for pattern in self.shipping_attachment_patterns:
                if pattern.search(attachment):
                    scores['SHIPPING'] += 3
                    
            for pattern in self.po_attachment_patterns:
                if pattern.search(attachment):
                    scores['PURCHASE_ORDER'] += 3
        
        # Check subject patterns
        for pattern in self.invoice_subject_patterns:
            if pattern.search(subject):
                scores['INVOICE'] += 2
        
        # Check keywords in subject and body
        content = subject + " " + body
        
        for keyword in self.invoice_keywords:
            if keyword in content:
                scores['INVOICE'] += 1
                
        for keyword in self.shipping_keywords:
            if keyword in content:
                scores['SHIPPING'] += 1
                
        for keyword in self.po_keywords:
            if keyword in content:
                scores['PURCHASE_ORDER'] += 1
                
        for keyword in self.false_positive_keywords:
            if keyword in content:
                scores['OTHER'] += 1
        
        # Determine classification
        max_score = max(scores.values())
        if max_score == 0:
            return 'OTHER'
        
        # Handle ties or ambiguous cases
        top_categories = [cat for cat, score in scores.items() if score == max_score]
        if len(top_categories) > 1:
            # Priority: INVOICE > SHIPPING > PURCHASE_ORDER > OTHER
            if 'INVOICE' in top_categories:
                return 'INVOICE'
            elif 'SHIPPING' in top_categories:
                return 'SHIPPING'
            elif 'PURCHASE_ORDER' in top_categories:
                return 'PURCHASE_ORDER'
            else:
                return 'OTHER'
        
        return top_categories[0]
        
    def sanitize_email_data(self, email_data):
        """Sanitize sensitive information"""
        # Sanitize subject
        subject = email_data['subject']
        subject = re.sub(r'\d+', 'XXX', subject)
        subject = re.sub(r'\$[\d,]+\.?\d*', '$XXX.XX', subject)
        subject = re.sub(r'\d{1,2}/\d{1,2}/\d{2,4}', 'XX/XX/XXXX', subject)
        email_data['subject_sanitized'] = subject
        
        # Sanitize attachments
        attachments_sanitized = []
        for attachment in email_data['attachment_names']:
            att = re.sub(r'\d+', 'XXX', attachment)
            attachments_sanitized.append(att)
        email_data['attachments_sanitized'] = attachments_sanitized
        
        # Extract domain from email
        if '@' in email_data['from_email']:
            domain = '@' + email_data['from_email'].split('@')[1]
            email_data['from_domain'] = domain
        else:
            email_data['from_domain'] = '@unknown.com'
        
        # Extract keywords from body
        email_data['body_keywords'] = self.extract_keywords(email_data)
        
        return email_data
        
    def extract_keywords(self, email_data):
        """Extract relevant keywords from email content"""
        content = (email_data['subject'] + " " + email_data['body']).lower()
        
        # Priority keywords based on classification
        priority_keywords = []
        
        if email_data['email_type'] == 'INVOICE':
            priority_keywords = [kw for kw in self.invoice_keywords if kw in content]
        elif email_data['email_type'] == 'SHIPPING':
            priority_keywords = [kw for kw in self.shipping_keywords if kw in content]
        elif email_data['email_type'] == 'PURCHASE_ORDER':
            priority_keywords = [kw for kw in self.po_keywords if kw in content]
        else:
            # For OTHER, find any relevant business terms
            all_keywords = self.invoice_keywords + self.shipping_keywords + self.po_keywords
            priority_keywords = [kw for kw in all_keywords if kw in content]
        
        # Limit to 10 keywords
        return ', '.join(priority_keywords[:10])
        
    def extract_vendor_name(self, email_data):
        """Extract vendor name from email"""
        # Try from name first
        if email_data['from_name']:
            # Clean common suffixes
            vendor = email_data['from_name']
            vendor = re.sub(r'\s*(LLC|Inc|Corp|Co|Ltd)\.?$', '', vendor, flags=re.IGNORECASE)
            return vendor.strip()
        
        # Try domain name
        domain = email_data['from_domain'].replace('@', '')
        if domain and domain != 'unknown.com':
            # Extract company name from domain
            parts = domain.split('.')
            if len(parts) > 1:
                company = parts[0]
                # Clean common prefixes
                company = re.sub(r'^(mail|info|ar|accounting|noreply)\.', '', company)
                return company.capitalize()
        
        return None
        
    def update_vendor_data(self, vendor_name, email_data):
        """Update vendor pattern data"""
        vendor = self.vendor_data[vendor_name]
        vendor['count'] += 1
        vendor['domains'].add(email_data['from_domain'])
        
        # Store subject pattern
        if email_data['subject_sanitized']:
            vendor['subject_patterns'].append(email_data['subject_sanitized'])
        
        # Store attachment patterns
        for att in email_data.get('attachments_sanitized', []):
            if att:
                vendor['attachment_patterns'].append(att)
                
    def generate_classification_csv(self):
        """Generate the invoice_classification_data.csv file"""
        output_file = self.output_directory / 'invoice_classification_data.csv'
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Email_Type', 'From', 'Subject', 'Attachments', 'Body_Keywords']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            
            for email_data in self.processed_emails:
                row = {
                    'Email_Type': email_data['email_type'],
                    'From': email_data['from_domain'],
                    'Subject': email_data['subject_sanitized'],
                    'Attachments': '; '.join(email_data['attachments_sanitized']) if email_data['attachments_sanitized'] else 'none',
                    'Body_Keywords': email_data['body_keywords']
                }
                writer.writerow(row)
        
        print(f"Created: {output_file}")
        
    def generate_vendor_patterns_csv(self):
        """Generate the vendor_patterns.csv file"""
        output_file = self.output_directory / 'vendor_patterns.csv'
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Vendor_Name', 'Email_Domain', 'Subject_Pattern', 'Attachment_Pattern', 'Email_Count']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            
            for vendor_name, vendor_info in self.vendor_data.items():
                # Get most common patterns
                subject_pattern = ''
                if vendor_info['subject_patterns']:
                    # Find most common subject pattern
                    from collections import Counter
                    subject_counter = Counter(vendor_info['subject_patterns'])
                    subject_pattern = subject_counter.most_common(1)[0][0]
                
                attachment_pattern = ''
                if vendor_info['attachment_patterns']:
                    # Find most common attachment pattern
                    from collections import Counter
                    att_counter = Counter(vendor_info['attachment_patterns'])
                    attachment_pattern = att_counter.most_common(1)[0][0]
                
                row = {
                    'Vendor_Name': vendor_name,
                    'Email_Domain': ', '.join(vendor_info['domains']),
                    'Subject_Pattern': subject_pattern,
                    'Attachment_Pattern': attachment_pattern,
                    'Email_Count': vendor_info['count']
                }
                writer.writerow(row)
        
        print(f"Created: {output_file}")
        
    def print_statistics(self):
        """Print processing statistics"""
        print("\n=== Processing Statistics ===")
        
        # Count by type
        type_counts = defaultdict(int)
        for email in self.processed_emails:
            type_counts[email['email_type']] += 1
        
        print("\nEmail Type Distribution:")
        for email_type, count in sorted(type_counts.items()):
            print(f"  {email_type}: {count}")
        
        print(f"\nTotal Vendors Found: {len(self.vendor_data)}")
        print("\nTop 5 Vendors by Email Count:")
        sorted_vendors = sorted(self.vendor_data.items(), key=lambda x: x[1]['count'], reverse=True)[:5]
        for vendor_name, info in sorted_vendors:
            print(f"  {vendor_name}: {info['count']} emails")

def main():
    """Main function for command line usage"""
    parser = argparse.ArgumentParser(description='Process invoice emails for AI classification training')
    parser.add_argument('input_directory', help='Directory containing .eml files')
    parser.add_argument('output_directory', help='Directory for output CSV files')
    parser.add_argument('--stats', action='store_true', help='Print processing statistics')
    
    args = parser.parse_args()
    
    # Create processor and run
    processor = InvoiceEmailProcessor(args.input_directory, args.output_directory)
    processor.process_all_emails()
    
    if args.stats:
        processor.print_statistics()

if __name__ == "__main__":
    main()