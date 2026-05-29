"""Clean Excel file by removing/clearing cells that contain gibberish or excessive numbers.

Keeps rows where the target column contains a probable full sentence.

Usage:
  python translation/clean_excel.py --file input.xlsx --col english --mode drop
  python translation/clean_excel.py --file input.xlsx --col english --mode clear --dry-run
"""
from pathlib import Path
import argparse
import re
import sys
import os
from typing import Tuple

# Reconfigure stdout/stderr to avoid UnicodeEncodeError on Windows terminals
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import pandas as pd


def analyze_text(text: str) -> Tuple[bool, dict]:
    """Return (keep, diagnostics) where keep=True means text looks like a full sentence.

    Diagnostics contains heuristics used.
    """
    if not isinstance(text, str):
        return False, {"reason": "not_string"}

    s = text.strip()
    if not s:
        return False, {"reason": "empty"}

    # attempt to fix common OCR word-splitting errors (hyphenation and internal spaces)
    def fix_broken_words(text: str) -> Tuple[str, bool]:
        corrected = False

        # fix hyphenated splits like 'fo-od' -> 'food' when both sides letters
        def _join_hyphen(m):
            nonlocal corrected
            a, b = m.group(1), m.group(2)
            candidate = a + b
            # require at least one vowel after join
            if re.search(r'[aeiouAEIOU]', candidate):
                corrected = True
                return candidate
            return m.group(0)

        text = re.sub(r"\b([A-Za-z]{2,})-([A-Za-z]{2,})\b", _join_hyphen, text)

        # fix internal splits like 'resis tance' -> 'resistance' when both parts alpha
        def _join_space_splits(t: str) -> Tuple[str, bool]:
            parts = re.split(r"(\s+)", t)
            i = 0
            changed = False
            out = []
            while i < len(parts):
                token = parts[i]
                if i + 2 < len(parts) and re.match(r'^[A-Za-z]{2,}$', token) and re.match(r'^[A-Za-z]{2,}$', parts[i+2]):
                    cand = token + parts[i+2]
                    if re.search(r'[aeiouAEIOU]', cand):
                        out.append(cand)
                        i += 3
                        changed = True
                        continue
                out.append(token)
                i += 1
            return (''.join(out), changed)

        fixed, ch = _join_space_splits(text)
        if ch:
            corrected = True
            text = fixed

        return text, corrected

    # stricter rules: must be a single full sentence (start with capital, end with a full stop)
    # reject if contains bullet-like markers or multiple lines
    # try to correct common broken words first
    s_fixed, was_fixed = fix_broken_words(s)
    if was_fixed:
        s = s_fixed

    if "\n" in s:
        return False, {"reason": "multiline"}

    # reject if leading bullet or any bullet-style line starts
    if re.match(r"^\s*[-\u2022\*\u2023\u25E6\u2043\u2219•–—]", s):
        return False, {"reason": "leading_bullet"}
    if re.search(r"(^|\n)\s*[-\u2022\*\u2023\u25E6\u2043\u2219•–—]", s):
        return False, {"reason": "contains_bullet"}

    # strip common leading quotes or brackets for checking start
    stripped = re.sub(r"^[\s\"'“”‘’\(\[]+", "", s)

    # must start with capital letter
    if not re.match(r'^[A-Z]', stripped):
        return False, {"reason": "no_initial_cap"}

    # must end with a sentence-ending punctuation: ., ?, ! (allow ? and !)
    end_stripped = re.sub(r"[\s\"'\)\]]+$", "", s)
    if not end_stripped.endswith(('.', '?', '!')):
        return False, {"reason": "no_trailing_sentence_punct"}
    words = [w for w in re.split(r"\s+", s) if w]
    word_count = len(words)

    # if it's an exclamation, reject obvious shouting/exclaimers: multiple !, all-caps, or very short
    if end_stripped.endswith('!'):
        if s.count('!') > 1:
            return False, {"reason": "excessive_exclaim"}
        if word_count < 3:
            return False, {"reason": "short_exclaim"}
        # reject if mostly uppercase (shouting)
        letters = [c for c in s if c.isalpha()]
        if letters and sum(1 for c in letters if c.isupper()) / len(letters) > 0.8:
            return False, {"reason": "shouting_exclaim"}

    # drop sentences containing common abbreviations or multi-dot acronyms that won't translate well
    abbrev_re = re.compile(r"\b(?:Mr|Mrs|Ms|Dr|Prof|Sr|Jr|St|vs|etc|e\.g|i\.e|cf|fig|al|approx)\.", re.I)
    if abbrev_re.search(s):
        return False, {"reason": "contains_common_abbrev"}

    # multi-dot acronyms like U.S.A. or A.B.C. or initials sequences
    if re.search(r"\b(?:[A-Za-z]\.){2,}", s):
        return False, {"reason": "multi_dot_acronym"}
    if word_count < 3:
        return False, {"reason": "too_short_words", "word_count": word_count}

    # reject if too many digits
    digits = sum(1 for c in s if c.isdigit())
    if digits / max(1, len(s)) > 0.35:
        return False, {"reason": "too_many_digits"}

    # pass
    return True, {"reason": "ok", "word_count": word_count}


def clean_dataframe(df: pd.DataFrame, col: str, mode: str = "drop") -> Tuple[pd.DataFrame, dict]:
    """Return cleaned DataFrame and summary dict.

    mode: 'drop' to drop rows where col fails, 'clear' to blank the cell but keep rows.
    """
    if col not in df.columns:
        raise ValueError(f"Column '{col}' not found in DataFrame")

    keep_mask = []
    diagnostics = []

    for val in df[col].fillna("").astype(str):
        keep, info = analyze_text(val)
        keep_mask.append(bool(keep))
        diagnostics.append(info)

    summary = {
        "total": len(df),
        "kept": sum(1 for k in keep_mask if k),
        "removed": sum(1 for k in keep_mask if not k),
    }

    df2 = df.copy()
    if mode == "drop":
        df2 = df2[[k for k in keep_mask]]
    elif mode == "clear":
        for i, keep in enumerate(keep_mask):
            if not keep:
                df2.at[i, col] = ""
    else:
        raise ValueError("mode must be 'drop' or 'clear'")

    return df2, summary


def main():
    parser = argparse.ArgumentParser(description="Remove gibberish/number-noise cells from an Excel file")
    parser.add_argument("--file", "-f", help="Input Excel file (.xlsx). If omitted, select from output/ocr/")
    parser.add_argument("--list", "-L", action="store_true", help="List available .xlsx files in output/ocr/ and exit")
    parser.add_argument("--col", "-c", help="Column to check (default: first column)")
    parser.add_argument("--mode", "-m", choices=["drop", "clear"], default="drop", help="Drop rows or clear cell")
    parser.add_argument("--dry-run", action="store_true", help="Don't write output, just show summary")
    parser.add_argument("--out", "-o", help="Output filename (optional). Defaults to cleaned_<input>")
    args = parser.parse_args()

    INPUT_DIR = Path("output/ocr")
    OUTPUT_DIR = Path("output/clean")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # If --list requested, show files and exit
    candidates = sorted(INPUT_DIR.glob("*.xlsx"))
    if args.list:
        if not candidates:
            print("No .xlsx files found in 'output/ocr/'. Run OCR first.")
            sys.exit(0)
        print("Available .xlsx files in output/ocr/:")
        for i, p in enumerate(candidates, start=1):
            print(f"  {i}. {p.name}")
        sys.exit(0)

    # If no --file provided, prompt selection from output/ocr/
    if not args.file:
        if not candidates:
            print("No .xlsx files found in 'output/ocr/'. Run OCR first or provide --file")
            sys.exit(1)
        print("Select a file to clean:")
        for i, p in enumerate(candidates, start=1):
            print(f"  {i}. {p.name}")
        choice = input("Enter number: ").strip()
        try:
            idx = int(choice) - 1
            input_path = candidates[idx]
        except Exception:
            print("Invalid selection")
            sys.exit(1)
    else:
        input_path = Path(args.file)
        # if user provided only a basename, try in input/
        if not input_path.exists():
            alt = INPUT_DIR / args.file
            if alt.exists():
                input_path = alt

    if not input_path.exists() or input_path.suffix.lower() != ".xlsx":
        print("Input file not found or not an .xlsx file")
        sys.exit(1)

    df = pd.read_excel(str(input_path), engine="openpyxl")

    col = args.col or df.columns[0]

    cleaned, summary = clean_dataframe(df, col, mode=args.mode)

    print("Summary:")
    print(f"  total rows: {summary['total']}")
    print(f"  kept rows : {summary['kept']}")
    print(f"  removed   : {summary['removed']}")

    out_file = args.out or str(OUTPUT_DIR / f"cleaned_{input_path.name}")

    if args.dry_run:
        print("Dry run: not writing output. Use --out to set output path or omit --dry-run to save.")
        return

    cleaned.to_excel(out_file, index=False, engine="openpyxl")
    print("Wrote:", out_file)


if __name__ == "__main__":
    main()
