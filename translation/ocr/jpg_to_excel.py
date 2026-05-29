from PIL import Image, ImageOps
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

import pandas as pd
import re
from pathlib import Path

MAX_WORDS = 38   

def sentence_aware_chunks(text, max_words=MAX_WORDS):
    text = re.sub(r'\s+', ' ', text).strip()
    if not text:
        return []

    sentences = re.split(r'(?<=[.!?])\s+', text)
    chunks, current, count = [], [], 0

    for s in sentences:
        words = s.split()

        # If a single sentence is longer than max_words
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

def process_image(image_path):
    img = Image.open(image_path).convert("L")
    img = ImageOps.autocontrast(img)
    text = pytesseract.image_to_string(img, lang="eng")
    return sentence_aware_chunks(text)

def main():
    folder_name = input("Enter the image folder name: ").strip()
    folder = Path(folder_name)

    if not folder.exists() or not folder.is_dir():
        print("Folder does not exist.")
        return

    images = list(folder.glob("*.jpg")) + list(folder.glob("*.jpeg"))
    if not images:
        print("No JPG/JPEG files found in the folder.")
        return

    all_chunks = []
    for img in images:
        print(f"Processing: {img.name}")
        chunks = process_image(img)
        all_chunks.extend(chunks)

    output_file = f"{folder.name}_output.xlsx"
    df = pd.DataFrame({"english": all_chunks})
    df.to_excel(output_file, index=False)

    print("\nExcel file created successfully")
    print(f"Output file: {output_file}")
    print(f"Total number of chunks: {len(all_chunks)}")

if __name__ == "__main__":
    main()
