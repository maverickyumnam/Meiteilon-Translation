import sys
from pathlib import Path
import pandas as pd

PDF_PATH = None
if len(sys.argv) > 1:
    PDF_PATH = sys.argv[1]
else:
    print("Usage: pdf_to_xlsx.py <pdf_path>")
    sys.exit(2)

pdf = Path(PDF_PATH)
if not pdf.exists():
    print(f"File not found: {pdf}")
    sys.exit(1)

texts = []

# Try PyPDF2 / pypdf
try:
    try:
        from PyPDF2 import PdfReader
    except Exception:
        # pypdf package
        from pypdf import PdfReader

    reader = PdfReader(str(pdf))
    for i, page in enumerate(reader.pages):
        try:
            t = page.extract_text() or ""
        except Exception:
            t = ""
        texts.append(t)
except Exception as e:
    # Try PyMuPDF (fitz)
    try:
        import fitz
        doc = fitz.open(str(pdf))
        for p in doc:
            texts.append(p.get_text())
    except Exception as e2:
        print("Failed to extract text: need PyPDF2/pypdf or PyMuPDF (fitz)")
        sys.exit(1)

# Create DataFrame: one row per page
rows = [{'english': t} for t in texts]
if not rows:
    print('No text extracted from PDF')
    sys.exit(1)

out_xlsx = pdf.with_suffix('.xlsx')
df = pd.DataFrame(rows)
df.to_excel(out_xlsx, index=False)
print('Wrote', out_xlsx)
