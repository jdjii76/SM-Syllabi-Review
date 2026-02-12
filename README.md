# Syllabus Text Extractor

A PyQt6-based GUI application for extracting specific text sections from syllabus documents (TXT, PDF, DOCX) and exporting them to Excel.

## Features

- **Multi-format Support**: Load syllabi from .txt, .pdf, and .docx files
- **Automatic Course Info Extraction**: Extracts course code and title from the syllabus heading (e.g., "SM 2200 HISTORY AND CONTEMPORARY ASPECTS OF SPORT")
- **Multiple File Handling**: Load and process multiple syllabi at once
- **Manual Text Selection**: Drag to highlight and select any text from loaded documents
- **Predefined Sections**: Quick checkboxes for common syllabus sections:
  - Course Information
  - Instructor Information
  - Course Description
  - Prerequisites
  - Credit Hours
  - Learning Outcomes
  - Course Materials
  - Required Text
  - Course Requirements
  - Grading Policy
  - Grading Scale
  - Attendance Policy
  - Late Work Policy
  - Academic Integrity
  - Disability Services
  - Course Schedule
  - And more
  - **Preserves Original Formatting**: Extracts text with original capitalization and punctuation intact

- **Excel Export**: Export multiple syllabi to a single Excel file with each course as a row and sections as columns
  - Automatic course code and title extraction
  - Each syllabus gets its own row
  - All selected sections appear as columns
  - Professionally formatted with headers and borders

## Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

### Setup

1. Navigate to the project directory:
```bash
cd "p:\dev\Projects\SM Syllabi Review"
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
venv\Scripts\activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python syllabi_extractor.py
```

2. **Load Files**: Click "Load File(s)" to select one or more syllabus documents
   - Supports .txt, .pdf, and .docx formats
   - Multiple files can be loaded at once
   - Each file is added to the "Loaded Files" list

3. **View Content**: Click on a file in the "Loaded Files" list to preview its content in the text editor
   - You can preview different files, but selection and export work with whichever is currently selected

4. **Select Text**: 
   - **Manual Selection**: Drag to highlight any text in the preview area - it will appear in the "Selected Text" box (only from currently viewed file)
   - **Predefined Sections**: Check boxes for sections you want to extract from ALL loaded files (the app will search for them automatically)

5. **Export**: Click "Export to Excel" to save your selections to an Excel file
   - Each loaded syllabus becomes one row in the Excel file
   - Course code and title are automatically extracted and included
   - All selected sections appear as columns
   - You'll be prompted to choose a location and filename
   - The Excel file includes formatting with headers, borders, and word wrapping

## Dependencies

- **PyQt6**: GUI framework
- **python-docx**: Support for .docx files
- **PyPDF2**: Support for .pdf files
- **openpyxl**: Excel file generation and formatting

## File Structure

```
SM Syllabi Review/
├── syllabi_extractor.py    # Main application file
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Notes

- The app automatically extracts course codes and titles from syllabus headings (e.g., "SM 2200: HISTORY AND CONTEMPORARY ASPECTS OF SPORT")
- Predefined sections are searched case-insensitively within the document
- Selected text appears in real-time in the "Selected Text" display
- Excel exports include formatting with headers, borders, and word wrapping
- Multiple files can be loaded and exported simultaneously
- Each syllabus becomes a separate row in the Excel file
- Frozen header row for easy scrolling through large exports

## Troubleshooting

### PyQt6 Installation Issues
On some systems, you may need to install PyQt6 separately:
```bash
pip install --upgrade PyQt6
```

### PDF Extraction Issues
If PDF text extraction is not working properly, ensure PyPDF2 is correctly installed:
```bash
pip install --upgrade PyPDF2
```

### Word Document Issues
For better .docx compatibility, ensure python-docx is up to date:
```bash
pip install --upgrade python-docx
```

## License

Created for Kennesaw State University - Department of Exercise Science and Sport Management

## Future Enhancements

- Batch text selection from multiple files at once
- Custom section definition and templates
- Find and replace functionality
- Dark mode
- Recent files list
- Undo/Redo functionality
- PDF form filling from extracted data
- Email delivery of exports
- Cloud storage integration (OneDrive, Google Drive)
