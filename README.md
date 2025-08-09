# Smart PDF to Excel Converter

A powerful web application that converts any PDF document into organized Excel format using intelligent content detection and OCR technology.

## ✨ Features

- **Universal PDF Support**: Works with any type of PDF (reports, forms, tables, documents, etc.)
- **Intelligent Content Detection**: Automatically identifies and categorizes different content types
- **Multi-Sheet Excel Output**: Creates separate sheets for tables, lists, headers, and paragraphs
- **Smart Formatting**: Auto-sizes columns, applies styling, and creates summary reports
- **Modern Web Interface**: Beautiful, responsive design with real-time feedback

## 🎯 What It Detects

- **Tables & Data Grids**: Automatically detects and formats tabular data
- **Lists & Bullet Points**: Organizes numbered and bulleted lists
- **Headers & Titles**: Identifies document sections and headings
- **Paragraphs & Text**: Preserves general text content with proper formatting

## 🚀 Setup Instructions

### Prerequisites

1. **Python 3.7+** installed on your system
2. **Tesseract OCR** for text extraction
3. **Poppler** for PDF to image conversion

### Installation

1. **Clone the repository and create a virtual environment**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install Tesseract OCR**
   - **Ubuntu/Debian**: `sudo apt-get install tesseract-ocr`
   - **macOS (Homebrew)**: `brew install tesseract`
   - **Windows**: Download from [UB Mannheim builds](https://github.com/UB-Mannheim/tesseract/wiki)
     - Update the path in `app.py` line 25: `pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"`

4. **Install Poppler** (for PDF to image conversion)
   - **Ubuntu/Debian**: `sudo apt-get install poppler-utils`
   - **macOS (Homebrew)**: `brew install poppler`
   - **Windows**: Download from [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases/)
     - Extract to `C:\poppler\` and update the path in `app.py` line 28: `POPPLER_PATH = r"C:\poppler\poppler-24.08.0\Library\bin"`

5. **Run the application**
   ```bash
   flask run
   ```
   The app will be available at `http://127.0.0.1:5000`

## 📁 Project Structure

```
HBA-PDF-to-EXCEL-/
├── app.py              # Main Flask application
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html     # Web interface
├── static/
│   └── style.css      # Styling
├── uploads/           # Temporary PDF storage
└── output/            # Generated Excel files
```

## 🔧 How It Works

1. **PDF Upload**: User uploads any PDF file through the web interface
2. **Image Conversion**: PDF pages are converted to high-quality images using Poppler
3. **OCR Processing**: Tesseract OCR extracts text from each page
4. **Content Analysis**: AI algorithms detect and categorize different content types:
   - Tables (using separator detection)
   - Lists (numbered/bulleted items)
   - Headers (short, formatted text)
   - Paragraphs (general text content)
5. **Excel Generation**: Creates a multi-sheet Excel file with:
   - Separate sheets for each content type
   - Formatted tables with headers
   - Summary sheet with conversion statistics
   - Auto-sized columns for readability

## 📊 Output Format

The generated Excel file contains:

- **Summary Sheet**: Overview of extracted content types and counts
- **Tables_X**: Formatted tables with headers and styling
- **Lists_X**: Organized lists and bullet points
- **Content_X**: General paragraphs and text content
- **Headers_X**: Document headers and section titles

## 🎨 Features

- **Responsive Design**: Works on desktop, tablet, and mobile devices
- **Real-time Processing**: Shows progress and handles errors gracefully
- **Modern UI**: Beautiful gradient design with smooth animations
- **Error Handling**: Comprehensive error messages and validation
- **File Management**: Automatic cleanup of temporary files

## 🔍 Technical Details

- **OCR Engine**: Tesseract 4.0+ with optimized configuration
- **PDF Processing**: Poppler for high-quality image conversion
- **Excel Generation**: OpenPyXL for advanced formatting and styling
- **Web Framework**: Flask with Bootstrap 5 for responsive design
- **Content Detection**: Custom algorithms for table, list, and header detection

## 🐛 Troubleshooting

- **Tesseract not found**: Update the path in `app.py` line 25
- **Poppler not found**: Update the path in `app.py` line 28
- **Processing errors**: Check that uploaded files are valid PDFs
- **Memory issues**: Large PDFs may require more RAM

## 📝 License

This project is open source and available under the MIT License.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.


## Testing

Unit tests build a tiny PDF from a base64 string to avoid storing binary fixtures in the repository. To update the embedded fixture, encode a PDF file and replace the value in `tests/test_process_pdf.py`:

```bash
python - <<'PY'
import base64, pathlib
pdf_bytes = pathlib.Path('sample.pdf').read_bytes()
print(base64.b64encode(pdf_bytes).decode())
PY
```
