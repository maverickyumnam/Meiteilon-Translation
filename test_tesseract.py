import os
import pytesseract

tess_path = r'C:\IMALab\manipuri-english-translation-pipeline\Tesseract-OCR\tesseract.exe'

print(f"Checking path: {tess_path}")
if os.path.exists(tess_path):
    print("SUCCESS: File found at the specified path.")
    try:
        pytesseract.pytesseract.tesseract_cmd = tess_path
        print("Tesseract version: " + pytesseract.get_tesseract_version().version)
    except Exception as e:
        print(f"ERROR: Found file, but couldn't execute it: {e}")
else:
    print("FAILURE: File NOT found. Check the folder spelling/path exactly.")