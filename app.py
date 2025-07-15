"""Flask application to convert PDF bank statements to Excel using OCR."""

import os
import tempfile
import re

from flask import (
    Flask,
    render_template,
    request,
    send_from_directory,
    redirect,
    url_for,
    flash,
)
from werkzeug.utils import secure_filename
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR"
from PIL import Image

# Folder configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

# Only PDF files are accepted
ALLOWED_EXTENSIONS = {'pdf'}

# Initialize the Flask application
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.secret_key = 'supersecretkey'

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""

    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def parse_text(text):
    """Parse OCR text into structured rows using a simple regex."""

    rows = []
    pattern = re.compile(
        r"(?P<Date>\d{2}/\d{2}/\d{4})?\s*(?P<Description>.+?)\s+"
        r"(?P<Debit>-?\d+[\d,]*\.\d{2})?\s+"
        r"(?P<Credit>-?\d+[\d,]*\.\d{2})?\s+"
        r"(?P<Balance>-?\d+[\d,]*\.\d{2})?"
    )
    for line in text.split("\n"):
        match = pattern.search(line)
        if match:
            data = match.groupdict()
            if any(data.values()):
                # Empty values are replaced with empty strings
                rows.append({k: v if v else "" for k, v in data.items()})
    return rows


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    """Handle file upload and trigger PDF processing."""

    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            try:
                excel_filename = process_pdf(filepath)
                return redirect(url_for('download_file', filename=excel_filename))
            except Exception as e:
                flash(f'Processing failed: {e}')
                return redirect(request.url)
    return render_template('index.html')


def process_pdf(path):
    """Convert PDF pages to images, run OCR, and build an Excel file."""

    with tempfile.TemporaryDirectory() as tempdir:
        # Convert each page of the PDF to an image
        images = convert_from_path(path, fmt="jpeg", output_folder=tempdir)

        all_rows = []
        for page_num, img in enumerate(images, start=1):
            # Run OCR on the page image
            text = pytesseract.image_to_string(img)
            rows = parse_text(text)

            # Attach page number to each extracted row
            for r in rows:
                r["Page"] = page_num
            all_rows.extend(rows)

        if not all_rows:
            raise ValueError("No data extracted from PDF.")

        df = pd.DataFrame(all_rows)
        excel_name = os.path.splitext(os.path.basename(path))[0] + ".xlsx"
        excel_path = os.path.join(app.config["OUTPUT_FOLDER"], excel_name)
        df.to_excel(excel_path, index=False)
        return excel_name


@app.route('/output/<filename>')
def download_file(filename):
    """Serve the generated Excel file for download."""

    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    # Run the Flask development server
    app.run(debug=True)
