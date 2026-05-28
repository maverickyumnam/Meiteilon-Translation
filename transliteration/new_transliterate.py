import os
import sys
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
    print("Meitei Mayek -> Roman (standard) transliteration")
    filename = input("Enter Excel filename (e.g., input.xlsx): ").strip()

    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    if not os.path.exists(filename):
        print(f"File not found: {filename}")
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

    df.to_excel(out_name, index=False, engine="openpyxl")
    print(f"Done. Saved transliterated file as: {out_name}")
    print("Column 'meitei_script' -> 'roman_standard' added.")

if __name__ == "__main__":
    main()
