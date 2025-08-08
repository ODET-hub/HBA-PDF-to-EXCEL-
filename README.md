# PDF to Excel Converter

This web application converts multi-page bank statement PDFs into structured Excel files using OCR.

## Features

- Upload PDF bank statements and convert them to Excel (.xlsx)
- Extraction handled page by page with Tesseract OCR
- Basic table parsing for Date, Description, Debit, Credit, and Balance columns
- Download the resulting Excel file

## Setup Instructions

1. **Clone the repository and create a virtual environment**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```
2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
3. **Install Tesseract OCR**
   - **Ubuntu/Debian**: `sudo apt-get install tesseract-ocr`
   - **macOS (Homebrew)**: `brew install tesseract`
   - **Windows**: Download the installer from [UB Mannheim builds](https://github.com/UB-Mannheim/tesseract/wiki) and add the installation directory to your PATH.

4. **Run the application**
   ```bash
   flask run
   ```
   The app will be available at `http://127.0.0.1:5000`.

### Notes

- PDF files are stored temporarily in the `uploads/` folder and converted Excel files are saved in the `output/` folder.
- Ensure Tesseract is correctly installed and available in your system PATH. Adjust `pytesseract.pytesseract.tesseract_cmd` in `app.py` if necessary.
- For larger PDFs (up to ~50 pages), processing time may vary depending on your hardware.


## Testing

Unit tests build a tiny PDF from a base64 string to avoid storing binary fixtures in the repository. To update the embedded fixture, encode a PDF file and replace the value in `tests/test_process_pdf.py`:

```bash
python - <<'PY'
import base64, pathlib
pdf_bytes = pathlib.Path('sample.pdf').read_bytes()
print(base64.b64encode(pdf_bytes).decode())
PY
```
