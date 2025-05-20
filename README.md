# PDF Document Splitter

A Python tool for intelligently splitting PDF documents into sections based on document type detection. This tool is particularly useful for processing import/export documentation where multiple document types are combined in a single PDF.

## Features

- **Intelligent Document Type Detection**
  - Exact matching for precise document identification
  - Fuzzy matching for similar document variations
  - Support for multiple document formats and variations

- **Two-Phase Processing**
  1. First Phase: Identifies documents with exact matches (80% confidence threshold)
  2. Second Phase: Processes remaining pages with fuzzy matching (90% confidence threshold)
  3. Creates an "Others" section for unmatched pages

- **Multiple Output Formats**
  - PDF: Preserves original document formatting
  - DOCX: Creates editable Word documents
  - TXT: Plain text output

- **Smart Document Grouping**
  - Groups similar document types (e.g., different variations of declaration certificates)
  - Maintains separation between distinct document types
  - Prevents false matches between different document categories

## Supported Document Types

The tool recognizes the following document types and their variations:

1. **Self Declaration Cum Undertaking Certificate**
   - Declaration Cum Undertaking Certificate
   - Self Declaration Certificate
   - Declaration Certificate

2. **Certificate of Origin**
   - Origin Certificate

3. **Certificate of Chemical Analysis Report**
   - Chemical Analysis Report
   - Chemical Analysis Certificate

4. **Form 6** (Exact match only)
5. **Form 9** (Exact match only)
6. **Bill of Exchange**
7. **Commercial Invoice**
8. **Packing List**
9. **Pre-Shipment Inspection Certificate**
10. **Insurance Policy**
11. **Transboundary Movement Document**
12. **Bill of Lading**

## Installation

1. Install required Python packages:
```bash
pip install PyPDF2 pdfplumber python-docx fuzzywuzzy python-Levenshtein
```

2. Clone or download this repository

## Usage

1. Run the script:
```bash
python document_splitter_v1.py
```

2. When prompted, enter the path to your PDF file

3. The script will:
   - Process the PDF in two phases
   - Create separate files for each identified document type
   - Generate an "Others" file for unmatched pages
   - Save all files in a directory named after the input file

## How It Works

### 1. Document Processing

The tool uses a two-phase approach to process documents:

#### Phase 1: Exact Matching
- Scans the first 5 lines of each page for exact document type matches
- Uses pattern matching and normalized text comparison
- Requires 80% confidence for a match
- Creates initial document sections

#### Phase 2: Fuzzy Matching
- Processes remaining unmatched pages
- Uses fuzzy string matching to identify similar document types
- Requires 90% confidence for a match
- Adds matched pages to existing sections or creates new ones

### 2. Document Section Management

Each document section contains:
- Heading (document type)
- Page numbers
- Content
- Confidence score
- Match type (exact or fuzzy)

### 3. Output Generation

The tool can save sections in three formats:
- PDF: Preserves original formatting and layout
- DOCX: Creates editable Word documents with headings
- TXT: Plain text with basic formatting

## Code Structure

### Main Classes

1. **DocumentSection**
   - Data class representing a document section
   - Stores metadata and content for each section

2. **DocumentSplitter**
   - Main class handling document processing
   - Implements document type detection and section management
   - Handles file I/O and format conversion

### Key Methods

- `process_pdf()`: Main processing pipeline
- `find_heading()`: Document type detection
- `extract_page_content()`: Text and table extraction
- `save_sections()`: Output file generation

## Customization

### Adjusting Matching Thresholds

You can modify the confidence thresholds in the `DocumentSplitter` initialization:
```python
splitter = DocumentSplitter(min_confidence=80, fuzzy_threshold=90)
```

### Adding New Document Types

To add new document types, update the `document_groups` dictionary in the `DocumentSplitter` class:
```python
self.document_groups = {
    "New Document Type": [
        "Exact Name",
        "Alternative Name",
        "Another Variation"
    ],
    # ... existing document types ...
}
```

## Error Handling

The tool includes comprehensive error handling:
- Logs all operations and errors
- Continues processing even if individual sections fail
- Provides detailed error messages for troubleshooting

## Best Practices

1. **Input PDFs**
   - Use good quality, text-based PDFs when possible
   - Ensure documents are properly scanned if using physical copies
   - Verify PDFs are not password-protected

2. **Document Organization**
   - Keep similar document types together in the input PDF
   - Use consistent naming conventions
   - Include clear document type indicators

3. **Output Management**
   - Review the "Others" section for missed documents
   - Verify fuzzy matches for accuracy
   - Check generated files for completeness

## Troubleshooting

Common issues and solutions:

1. **Missing Documents**
   - Check the "Others" section
   - Verify document type names match exactly
   - Adjust confidence thresholds if needed

2. **Incorrect Grouping**
   - Review document type variations
   - Adjust fuzzy matching threshold
   - Check for similar document names

3. **Format Issues**
   - Verify PDF is not corrupted
   - Check for password protection
   - Ensure sufficient disk space

## Contributing

Feel free to:
- Report bugs
- Suggest improvements
- Add new document types
- Enhance matching algorithms

## License

[Your License Here]

## Contact

[Your Contact Information]