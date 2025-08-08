import base64
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from pathlib import Path
from PIL import Image
from app import process_pdf, app as flask_app

# Base64-encoded minimal PDF used for tests.
# To regenerate PDF_BASE64, run:
#   python - <<'END'
#   import base64, pathlib
#   pdf_bytes = b"%PDF-1.1\n%\xe2\xe3\xcf\xd3\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Count 0/Kids[]>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF"
#   pathlib.Path('sample.pdf').write_bytes(pdf_bytes)
#   print(base64.b64encode(pdf_bytes).decode())
#   END

PDF_BASE64 = (
    "JVBERi0xLjEKJeLjz9MKMSAwIG9iajw8L1R5cGUvQ2F0YWxvZy9QYWdlcyAyIDAgUj4+ZW5kb2JqCjIgMCBvYmo8PC9UeXBlL1BhZ2VzL0NvdW50IDAvS2lkc1tdPj5lbmRvYmoKdHJhaWxlcjw8L1Jvb3QgMSAwIFI+PgolJUVPRg=="
)


def _write_sample_pdf(tmp_path: Path) -> Path:
    pdf_path = tmp_path / "sample.pdf"
    pdf_path.write_bytes(base64.b64decode(PDF_BASE64))
    return pdf_path


def test_process_pdf(monkeypatch, tmp_path):
    pdf_file = _write_sample_pdf(tmp_path)

    def fake_convert_from_path(path, fmt="jpeg", output_folder=None):
        img = Image.new("RGB", (10, 10), color="white")
        return [img]

    def fake_image_to_string(img):
        return "01/01/2020 Test 1.00 2.00 3.00"

    monkeypatch.setattr("app.convert_from_path", fake_convert_from_path)
    monkeypatch.setattr("app.pytesseract.image_to_string", fake_image_to_string)

    flask_app.config["OUTPUT_FOLDER"] = str(tmp_path)
    excel_name = process_pdf(str(pdf_file))
    assert (tmp_path / excel_name).exists()
