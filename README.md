# Manipuri–English Translation Pipeline

This project provides Python programs to:
- Extract English text from images and PDFs (OCR)
- Split text into sentence-aware chunks
- Translate English text to Manipuri (Meitei)
- Transliterate Manipuri text from Meitei Mayek to Roman script

This document explains **how to run the programs step by step**, starting from creating a virtual environment.

---

## Prerequisites

Before using the programs, make sure you have:

- Python 3.9 or newer installed
- Visual Studio Code (VS Code) installed
- Basic familiarity with opening folders and using the terminal

---

## Step 1: Open the Project in VS Code

1. Open **VS Code**
2. Click **File → Open Folder**
3. Select the project folder:
```bash
manipuri-english-translation-pipeline
```
4. Open the terminal:
```bash
Ctrl + `
```

Make sure the terminal is **PowerShell** (on Windows).

---

## Step 2: Create a Virtual Environment (venv)

A virtual environment keeps project libraries separate from system Python.

From the project folder, run:

```bash
python -m venv venv
```
This creates a folder named venv.

---

## Step 3: Activate the Virtual Environment

Activate the environment before installing packages or running scripts:
```bash
venv\Scripts\activate
```

If activated correctly, the terminal will show:
```bash
(venv)
```

---

## Step 4: Install Required Python Packages

Install all required libraries using:
```bash
pip install -r requirements.txt
```
This command installs everything needed for OCR, translation, Excel handling, and transliteration.

---

## Step 5: Project Structure

The project is organized as follows:
```text
manipuri-english-translation-pipeline/
│
├── ocr/
│   ├── jpg_to_excel.py        # OCR from image files
│   └── pdf_excel.py           # OCR from PDF files
│
├── translation/
│   └── trans_sarv.py          # English → Manipuri translation
│
├── transliteration/
│   └── new_transliterate.py   # Meitei Mayek → Roman transliteration
│
├── requirements.txt
├── README.md
└── venv/
```

---

## Step 6: Using the Programs
Make sure (venv) is active before running any program.

## 1. OCR from Images

This program extracts English text from .jpg or .jpeg images and saves it to an Excel file.

Run:
```bash
python ocr/jpg_to_excel.py
```
You will be asked to enter the name of a folder containing images.

Output:
An Excel file containing English sentence chunks.

## 2. OCR from PDF

This program extracts English text from a PDF file.

Run:
```
python ocr/pdf_excel.py
```
You will be asked to enter the PDF file path.

Output:
An Excel file with extracted English text.

## 3. English → Manipuri Translation

This program translates English text in an Excel file into Manipuri.

Before running, set the API key (only once):
```
$env:SARVAM_API_KEY="Your_SarvamAI_API_key"
```
You have to reset the api key each time you restart the terminal and activate venv.

Run:
```
python translation/trans_sarv.py
```
You will be asked to enter the Excel file name.

Output:
An Excel file containing both English and Manipuri text.

## 4. Meitei Mayek → Roman Transliteration

This program transliterates Manipuri text written in Meitei Mayek into Roman script.

Run:
```
python transliteration/new_transliterate.py
```
You will be asked to enter the Excel file name.

Output:
An Excel file with an additional column containing Roman transliteration.

---

## Notes:

Always activate the virtual environment before running programs

Run one program at a time, in order

Do not rename scripts unless you update commands

Excel files are created automatically in the project folder
