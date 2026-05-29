from PIL import Image, ImageOps
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import pandas as pd
import re
from pathlib import Path
from pdf2image import convert_from_path, pdfinfo_from_path
import tempfile
import gc
import argparse
import os

MAX_WORDS = 38   

def sentence_aware_chunks(text, max_words=MAX_WORDS):
    text = re.sub(r'\s+', ' ', text).strip()
    if not text:
        return []
    # Split into full sentences but avoid splitting after single-letter initials
    # Only split when the following token looks like a sentence start (capital + lowercase)
    sentences = re.split(r"(?<=[.!?])\s+(?=(?:[\"“”']?)[A-Z][a-z])", text)
    return [s.strip() for s in sentences if s and s.strip()]

def process_pdf(pdf_path, dpi: int = 200):
    print(f" Reading PDF: {pdf_path.name}")
    poppler_bin = r'C:\IMALab\manipuri-english-translation-pipeline\poppler-26.02.0\Library\bin'

    try:
        info = pdfinfo_from_path(str(pdf_path), userpw=None, poppler_path=poppler_bin)
        total_pages = int(info.get('Pages', 0))
    except Exception:
        total_pages = None

    all_chunks = []

    # Use a temporary folder so pdftoppm writes images to disk instead of streaming large blobs into memory
    with tempfile.TemporaryDirectory() as tmpdir:
        for i in range(1, (total_pages or 0) + 1):
            print(f"  └─ OCR page {i}/{total_pages}")
            try:
                imgs = convert_from_path(
                    str(pdf_path), dpi=dpi, first_page=i, last_page=i,
                    output_folder=tmpdir, fmt='png', poppler_path=poppler_bin
                )
                if not imgs:
                    continue
                page_img = imgs[0]
                img = page_img.convert("L")
                img = ImageOps.autocontrast(img)
                text = pytesseract.image_to_string(img, lang="eng")
                all_chunks.extend(sentence_aware_chunks(text))

                # explicitly close and delete image to release memory
                try:
                    page_img.close()
                    img.close()
                except Exception:
                    pass
                del page_img, img, imgs
                gc.collect()

            except MemoryError:
                print("MemoryError processing page", i, "— try lowering DPI or increasing system memory.")
                raise
            except Exception as e:
                print(f"Warning: failed to process page {i}: {e}")
                continue

    # filter out noise-like standalone fragments such as 'Chapter 1', 'Page 12', bare numbers, or roman numerals
    def is_noise_fragment(s: str) -> bool:
        if not isinstance(s, str):
            return True
        s_stripped = s.strip()
        if not s_stripped:
            return True

        s_lower = s_stripped.lower()

        chapter_re = re.compile(r'^(chapter|chap|ch)\b\s*[0-9ivx]+$', re.I)
        page_re = re.compile(r'^(page|pg|p)\b\s*\d+(\s*of\s*\d+)?$', re.I)
        digits_re = re.compile(r'^\d{1,4}$')
        roman_re = re.compile(r'^[ivxlcdm]{1,5}$', re.I)

        # topic/section headings like 'Topic 1' or 'Section 2.3'
        topic_re = re.compile(r'^(topic|section|sec)\b\s*\d+(?:[\.\d]*)$', re.I)

        # remove leading numbering prefixes like '1. ', '2.3. ', 'Chapter 1. ', '1) '
        prefix_re = re.compile(r'^\s*(?:chapter|chap|ch|section|sec|topic)?\s*\d+(?:[\.\)\:~-]*\d*)*\s*[:\)\.-]*\s*', re.I)

        # strip numbering prefix for inspection
        s_no_prefix = prefix_re.sub('', s_stripped).strip()
        if not s_no_prefix:
            return True

        s_no_prefix_lower = s_no_prefix.lower()

        if chapter_re.match(s_lower):
            return True
        if topic_re.match(s_lower):
            return True
        if page_re.match(s_lower):
            return True
        if digits_re.match(s_lower):
            return True
        if roman_re.match(s_lower):
            return True

        if chapter_re.match(s_lower):
            return True
        if page_re.match(s_lower):
            return True
        if digits_re.match(s_lower):
            return True
        if roman_re.match(s_lower):
            return True

        return False

    filtered = [c for c in all_chunks if not is_noise_fragment(c)]
    return filtered

def main():
    parser = argparse.ArgumentParser(description="OCR a PDF into an Excel file")
    parser.add_argument("--file", "-f", help="PDF filename (optional). If omitted, lists files in input/")
    args = parser.parse_args()

    INPUT_DIR = Path("input")
    OUTPUT_DIR = Path("output/ocr")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if args.file:
        pdf_path = Path(args.file)
        if not pdf_path.exists():
            alt = INPUT_DIR / args.file
            if alt.exists():
                pdf_path = alt
    else:
        candidates = sorted(INPUT_DIR.glob("*.pdf"))
        if not candidates:
            print("No PDF files found in 'input/'")
            return
        print("Available PDF files in 'input/':")
        for i, p in enumerate(candidates, start=1):
            print(f"  {i}. {p.name}")
        choice = input("Select number: ").strip()
        try:
            idx = int(choice) - 1
            pdf_path = candidates[idx]
        except Exception:
            print("Invalid selection")
            return

    if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
        print(" Invalid PDF file.")
        return

    chunks = process_pdf(pdf_path)

    if not chunks:
        print(" No text extracted.")
        return

    output_file = OUTPUT_DIR / f"{pdf_path.stem}_output.xlsx"
    df = pd.DataFrame({"english": chunks})
    df.to_excel(output_file, index=False)

    print("\n Excel file created successfully")
    print(f" Output file: {output_file}")
    print(f" Total number of chunks: {len(chunks)}")

if __name__ == "__main__":
    main()
