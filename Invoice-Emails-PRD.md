# PRD: EML Email Analysis Tool for Invoice Classification Training Data

**Project**: Email Sample Data Generator for Freckled Hen Inventory Automation  
**Purpose**: Generate training data for Gemini AI email classification system  

---

## **Project Overview**

Create a tool that processes .eml email files and generates two CSV files containing sanitized email data for training the invoice classification AI system.

### **Core Functionality**
1. **Email Parsing**: Extract metadata and content from .eml files
2. **Data Sanitization**: Protect sensitive information per business requirements  
3. **Classification**: Intelligently categorize emails based on content patterns
4. **Pattern Extraction**: Generate vendor communication patterns
5. **CSV Generation**: Create structured output files

---

## **Input Specifications**

### **Email File Format**
- **File Type**: `.eml` (standard email format)
- **Source**: Gmail exports or email client saves
- **Expected Volume**: 60-80 email files
- **Content**: Mix of invoice, shipping, PO, and false positive emails

### **Expected Email Types Distribution**
```
INVOICE (pure billing): 15-20 files
SHIPPING (delivery only): 15-20 files  
PURCHASE_ORDER (confirmations): 10-15 files
OTHER/FALSE_POSITIVE: 10-15 files
EDGE_CASES: 5-10 files
```

---

## **Output Specifications**

### **File 1: `invoice_classification_data.csv`**

#### **Required Columns**
| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `Email_Type` | String | Classification category | "INVOICE" |
| `From` | String | Sanitized sender domain | "@creativecoop.com" |
| `Subject` | String | Sanitized subject line | "Invoice #XXX from Creative Co-Op" |
| `Attachments` | String | Sanitized attachment names | "INV-XXX.pdf" |
| `Body_Keywords` | String | Key identifying phrases | "invoice, payment due, attached" |

#### **Data Sanitization Rules**
- **Numbers**: Replace all numbers with "XXX" (preserve pattern length)
- **Email Domains**: Extract domain only, include "@" prefix
- **Company Names**: Keep vendor names, sanitize customer references
- **Dates**: Replace with "XX/XX/XXXX" format
- **Amounts**: Replace with "$XXX.XX" pattern
- **Order IDs**: Replace with "XXX" preserving prefix/suffix patterns

### **File 2: `vendor_patterns.csv`**

#### **Required Columns**
| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `Vendor_Name` | String | Extracted vendor name | "Creative Co-Op" |
| `Email_Domain` | String | Sender domain | "@creativecoop.com" |
| `Subject_Pattern` | String | Common subject format | "Invoice #XXX" |
| `Attachment_Pattern` | String | Typical attachment naming | "INV-XXX.pdf" |
| `Email_Count` | Integer | Frequency in dataset | 3 |

---

## **Business Logic Requirements**

### **Email Classification Logic**

#### **INVOICE Detection**
```typescript
// Primary indicators (high confidence)
const invoiceKeywords = [
  'invoice', 'bill', 'statement', 'payment due', 
  'amount due', 'remit payment', 'balance due'
];

// Subject patterns
const invoiceSubjectPatterns = [
  /invoice\s*#?\d+/i,
  /bill\s*#?\d+/i,
  /statement\s*#?\d+/i
];

// Attachment patterns  
const invoiceAttachmentPatterns = [
  /inv[_-]?\d+\.pdf/i,
  /invoice.*\.pdf/i,
  /bill.*\.pdf/i
];
```

#### **SHIPPING Detection**
```typescript
const shippingKeywords = [
  'shipped', 'tracking', 'delivery', 'carrier',
  'tracking number', 'expected delivery', 'in transit'
];

const shippingSubjectPatterns = [
  /shipped/i,
  /tracking/i,
  /delivery/i,
  /order.*shipped/i
];
```

#### **PURCHASE_ORDER Detection**
```typescript
const poKeywords = [
  'purchase order', 'po confirmation', 'order confirmed',
  'order accepted', 'po #', 'order acknowledgment'
];

const poAttachmentPatterns = [
  /po[_-]?\d+\.pdf/i,
  /purchase.*order.*\.pdf/i,
  /confirmation.*\.pdf/i
];
```

#### **OTHER/FALSE_POSITIVE Detection**
```typescript
const falsePositiveKeywords = [
  'newsletter', 'promotion', 'sale', 'marketing',
  'unsubscribe', 'catalog', 'new products', 'follow us'
];

// No attachments or promotional attachments
const promotionalAttachments = [
  /catalog\.pdf/i,
  /flyer\.pdf/i,
  /promotion\.pdf/i
];
```

### **Classification Decision Tree**
1. **Check attachment patterns** (highest confidence)
2. **Analyze subject line keywords** (medium confidence)
3. **Scan body content** (lower confidence, more false positives)
4. **Apply business rules** (domain reputation, sender history)

---

## **Data Processing Requirements**

### **Email Parsing**
```typescript
interface ParsedEmail {
  from: string;           // Full sender address
  subject: string;        // Complete subject line
  body: string;          // Email body content
  attachments: Attachment[];
  date: Date;
  messageId: string;
}

interface Attachment {
  filename: string;
  contentType: string;
  size: number;
}
```

### **Keyword Extraction Algorithm**
```typescript
// Extract 5-10 most relevant keywords per email
function extractKeywords(emailContent: EmailContent): string[] {
  // 1. Remove common words (the, and, or, etc.)
  // 2. Weight business-relevant terms higher
  // 3. Include attachment-related keywords
  // 4. Limit to 10 words max, comma-separated
  // 5. Prioritize classification-relevant terms
}
```

### **Vendor Name Extraction**
```typescript
// Extract vendor name from sender or subject
function extractVendorName(email: ParsedEmail): string {
  // Priority order:
  // 1. Company name in "from" display name
  // 2. Domain name (before .com)  
  // 3. Subject line company reference
  // 4. Email signature parsing
}
```

---

## **Error Handling Requirements**

### **File Processing Errors**
- **Corrupted .eml files**: Log error, continue processing others
- **Missing attachments**: Mark as "none" in CSV
- **Encoding issues**: Attempt UTF-8, fallback to ASCII
- **Large files**: Warn if >10MB, process if possible

### **Classification Errors**
- **Ambiguous emails**: Classify as "EDGE_CASES"
- **Multiple indicators**: Use highest confidence classification
- **No clear indicators**: Classify as "OTHER"

### **Output Validation**
- **Required fields**: Ensure no empty required columns
- **CSV format**: Validate proper escaping and encoding
- **File creation**: Verify both output files created successfully

---

## **Success Criteria**

### **Functional Requirements**
- [ ] Successfully parse all .eml files without crashing
- [ ] Generate both required CSV files with correct schema
- [ ] Achieve target distribution of email types
- [ ] Properly sanitize all sensitive information
- [ ] Extract meaningful vendor patterns

### **Quality Requirements**  
- [ ] Classification accuracy >85% on manual review
- [ ] Zero sensitive data leakage (manual audit)
- [ ] All vendor domains correctly extracted
- [ ] Keyword extraction captures relevant terms
- [ ] Processing completes within performance targets

### **Business Requirements**
- [ ] Output files ready for Gemini AI training
- [ ] Sufficient examples for each email type
- [ ] Vendor patterns support automation rules
- [ ] Data format matches existing '[Invoice Emails - Output.csv](<Invoice Emails - Output.csv>) structure
- [ ] Tool reusable for future data collection

---

## **Testing Requirements**

### **Test Data**
- Provide 5-10 sample .eml files representing each email type
- Include edge cases: corrupted files, missing attachments, unusual formats
- Test with realistic vendor email samples

### **Validation Process**
1. **Manual review** of 20% of classifications
2. **Privacy audit** of sanitized data
3. **Format validation** of output CSV files
4. **Performance testing** with full dataset

---

**Deliverable**: Functional application that processes .eml files and generates the two specified CSV files according to all requirements above.