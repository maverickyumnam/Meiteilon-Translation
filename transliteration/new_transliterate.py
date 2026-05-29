import os
import sys
import re
import pandas as pd


# Transliteration table (Meitei character -> Roman transliteration)
MAPPING = {
    # Mapum Mayek
    "ꯀ": "ka",
    "ꯁ": "sa",
    "ꯂ": "la",
    "ꯃ": "ma",
    "ꯄ": "pa",
    "ꯅ": "na",
    "ꯆ": "cha",
    "ꯇ": "ta",
    "ꯈ": "kha",
    "ꯉ": "nga",
    "ꯊ": "tha",
    "ꯋ": "wa",
    "ꯌ": "ya",
    "ꯍ": "ha",
    "ꯎ": "u",
    "ꯏ": "i",
    "ꯐ": "pha",
    "ꯑ": "a",
    "ꯒ": "ga",
    "ꯓ": "jha",
    "ꯔ": "ra",
    "ꯕ": "ba",
    "ꯖ": "ja",
    "ꯗ": "da",
    "ꯘ": "gha",
    "ꯙ": "dha",
    "ꯚ": "bha",

    # Lonsum Mayek
    "ꯛ": "k",
    "ꯜ": "l",
    "ꯝ": "m",
    "ꯞ": "p",
    "ꯟ": "n",
    "ꯠ": "t",
    "ꯡ": "ng",
    "ꯢ": "ai",

    # Cheitap Mayek
    "ꯥ": "ā",
    "ꯦ": "e",
    "ꯣ": "o",
    "ꯧ": "ou",
    "ꯨ": "u",
    "ꯩ": "ei",
    "ꯤ": "i",
    "ꯪ": "ang",

    # Digits
    "꯰": "0",
    "꯱": "1",
    "꯲": "2",
    "꯳": "3",
    "꯴": "4",
    "꯵": "5",
    "꯶": "6",
    "꯷": "7",
    "꯸": "8",
    "꯹": "9",

    # Punctuation
    "।": ".",
    "॥": "..",
    "꧈": ",",
}

MAPUM_MAYEK = {
    "ꯀ", "ꯁ", "ꯂ", "ꯃ", "ꯄ", "ꯅ", "ꯆ", "ꯇ", "ꯈ", "ꯉ",
    "ꯊ", "ꯋ", "ꯌ", "ꯍ", "ꯐ", "ꯑ", "ꯒ", "ꯓ", "ꯔ", "ꯕ",
    "ꯖ", "ꯗ", "ꯘ", "ꯙ", "ꯚ"
}

CHEITAP_MAYEK = {
    "ꯥ", "ꯦ", "ꯣ", "ꯧ", "ꯨ", "ꯩ", "ꯤ", "ꯪ"
}

FALLBACKS = {
    "\uAAE0": "?"
}


# Transliteration function
def transliterate_meitei_to_roman(text: str) -> str:
    if not isinstance(text, str):
        return ""

    out = []
    chars = list(text)
    n = len(chars)
    i = 0

    while i < n:
        ch = chars[i]

        if ch.isspace():
            out.append(ch)
            i += 1
            continue

        if ch in MAPPING:
            roman = MAPPING[ch]

            # If Mapum followed by Cheitap, drop trailing 'a'
            if ch in MAPUM_MAYEK and i + 1 < n:
                nxt = chars[i + 1]
                if nxt in CHEITAP_MAYEK and roman.endswith("a"):
                    roman = roman[:-1]

            # If ꯪ follows ꯏ or ꯎ, remove 'a' from 'ang'
            if ch == "ꯪ" and i > 0:
                prev = chars[i - 1]
                if prev in {"ꯏ", "ꯎ"} and roman == "ang":
                    roman = "ng"

            out.append(roman)
            i += 1
            continue

        if ch in FALLBACKS:
            out.append(FALLBACKS[ch])
            i += 1
            continue

        out.append(ch)
        i += 1

    return "".join(out)

# Excel integration
def detect_meitei_column(df: pd.DataFrame) -> str:
    candidates = ["manipuri", "meitei", "meitei_script", "meitei_mayek", "text"]
    for c in candidates:
        if c in df.columns:
            return c
    return df.columns[0]

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Transliterate Meitei Mayek in an Excel file")
    parser.add_argument("--file", "-f", help="Excel filename (optional). If omitted, lists files in input/")
    args = parser.parse_args()

    INPUT_DIR = "input"

    if args.file:
        filename = args.file
        if not filename.lower().endswith(".xlsx"):
            filename += ".xlsx"
        if not os.path.exists(filename):
            alt = os.path.join(INPUT_DIR, filename)
            if os.path.exists(alt):
                filename = alt
        if not os.path.exists(filename):
            print(f"File not found: {filename}")
            sys.exit(1)
    else:
        files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith('.xlsx')]
        if not files:
            print("No .xlsx files found in 'input/'")
            sys.exit(1)
        print("Select a file to transliterate:")
        for i, f in enumerate(files, start=1):
            print(f"  {i}. {f}")
        choice = input("Enter number: ").strip()
        try:
            idx = int(choice) - 1
            filename = os.path.join(INPUT_DIR, files[idx])
        except Exception:
            print("Invalid selection")
            sys.exit(1)

    out_name = "output_" + os.path.basename(filename)

    df = pd.read_excel(filename, engine="openpyxl")
    meitei_col = detect_meitei_column(df)

    if meitei_col != "meitei_script":
        df = df.rename(columns={meitei_col: "meitei_script"})

    df["roman_standard"] = (
        df["meitei_script"]
        .fillna("")
        .astype(str)
        .apply(transliterate_meitei_to_roman)
    )

    # Split transliterated text into cells on sentence boundaries (full stop),
    # but avoid splitting after common honorifics, abbreviations, and acronyms.
    def split_into_cells(text: str):
        if not isinstance(text, str) or not text:
            return []

        # common abbreviations (lowercase, without trailing dot)
        abbrevs = {
            'mr', 'mrs', 'ms', 'dr', 'prof', 'sr', 'jr', 'miss', 'rev', 'st', 'mt',
            'inc', 'ltd', 'etc', 'eg', 'ie', 'vs', 'fig', 'jan', 'feb', 'mar', 'apr',
            'jun', 'jul', 'aug', 'sep', 'sept', 'oct', 'nov', 'dec'
        }

        tokens = text.split()
        parts = []
        cur = []

        for i, tok in enumerate(tokens):
            cur.append(tok)
            # check if token ends with a period-like sentence terminator
            if tok.endswith('.') or tok.endswith('।') or tok.endswith('॥'):
                stripped = tok.rstrip('.।॥')
                lower = stripped.lower()

                # if token contains multiple periods (e.g., U.S.A.) treat as acronym -> don't split
                if tok.count('.') > 1:
                    continue

                # if stripped token is a single letter (initial) -> don't split
                if len(stripped) == 1 and stripped.isalpha():
                    continue

                # if stripped token is a known abbreviation (mr, mrs, e.g., etc) -> don't split
                if lower in abbrevs:
                    continue

                # otherwise, treat as sentence boundary
                parts.append(' '.join(cur).strip())
                cur = []

        # append remaining
        if cur:
            parts.append(' '.join(cur).strip())

        # normalize: remove empty parts and drop standalone headers/page numbers
        cleaned = []

        # compile filters
        chapter_re = re.compile(r'^(chapter|chap|ch)\b\s*[0-9ivx]+$', re.I)
        page_re = re.compile(r'^(page|pg|p)\b\s*\d+(\s*of\s*\d+)?$', re.I)
        digits_re = re.compile(r'^\d{1,4}$')
        roman_re = re.compile(r'^[ivxlcdm]{1,5}$', re.I)

        for p in parts:
            s = p.strip()
            s_lower = s.lower()

            # skip if exactly 'chapter 1' or 'page 12' or just a page number or roman numeral
            if chapter_re.match(s_lower):
                continue
            if page_re.match(s_lower):
                continue
            if digits_re.match(s_lower):
                continue
            if roman_re.match(s_lower):
                continue

            cleaned.append(p)

        return [p for p in cleaned if p]

    # apply splitting and expand into columns
    split_lists = df["roman_standard"].apply(split_into_cells)
    max_parts = max((len(x) for x in split_lists), default=0)

    for idx in range(max_parts):
        colname = f"roman_part_{idx+1}"
        df[colname] = split_lists.apply(lambda lst, i=idx: lst[i] if i < len(lst) else "")

    df.to_excel(out_name, index=False, engine="openpyxl")
    print(f"Done. Saved transliterated file as: {out_name}")
    print("Column 'meitei_script' -> 'roman_standard' added.")

if __name__ == "__main__":
    main()
