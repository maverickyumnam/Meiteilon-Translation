from PIL import Image, ImageOps
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import pandas as pd
import re
from pathlib import Path
from pdf2image import convert_from_path

MAX_WORDS = 38   

def sentence_aware_chunks(text, max_words=MAX_WORDS):
    text = re.sub(r'\s+', ' ', text).strip()
    if not text:
        return []

    sentences = re.split(r'(?<=[.!?])\s+', text)
    chunks, current, count = [], [], 0

    for s in sentences:
        words = s.split()

        if len(words) > max_words:
            if current:
                chunks.append(" ".join(current))
                current, count = [], 0
            for i in range(0, len(words), max_words):
                chunks.append(" ".join(words[i:i+max_words]))
            continue

        if count + len(words) > max_words:
            chunks.append(" ".join(current))
            current, count = [s], len(words)
        else:
            current.append(s)
            count += len(words)

    if current:
        chunks.append(" ".join(current))

    return chunks

def process_pdf(pdf_path):
    print(f" Reading PDF: {pdf_path.name}")
    # Look for this line in ocr/pdf_excel.py and change it:
    pages = convert_from_path(pdf_path, dpi=300, poppler_path=r'C:\IMALab\manipuri-english-translation-pipeline\poppler-26.02.0\Library\bin')
    all_chunks = []

    for i, page in enumerate(pages, start=1):
        print(f"  └─ OCR page {i}")
        img = page.convert("L")
        img = ImageOps.autocontrast(img)
        text = pytesseract.image_to_string(img, lang="eng")
        all_chunks.extend(sentence_aware_chunks(text))

    return all_chunks

def main():
    pdf_file = input(" Enter PDF file path: ").strip()
    pdf_path = Path(pdf_file)

    if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
        print(" Invalid PDF file.")
        return

    chunks = process_pdf(pdf_path)

    if not chunks:
        print(" No text extracted.")
        return

    output_file = f"{pdf_path.stem}_output.xlsx"
    df = pd.DataFrame({"english": chunks})
    df.to_excel(output_file, index=False)

    print("\n Excel file created successfully")
    print(f" Output file: {output_file}")
    print(f" Total number of chunks: {len(chunks)}")

if __name__ == "__main__":
    main()
