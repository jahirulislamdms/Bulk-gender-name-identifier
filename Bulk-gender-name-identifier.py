#!/usr/bin/env python3
"""
Bulk-gender-name-identifier.py

Terminal-driven tool with tkinter dialogs to map first names to genders using a reference CSV.
Features:
 - Unicode normalization + diacritic removal (NFKD)
 - Optional Unidecode transliteration (if installed) to improve multilingual handling
 - Exact normalized matching + fuzzy matching (rapidfuzz)
 - Terminal progress bars (tqdm)
 - Input: CSV, XLS, XLSX
 - Output: XLSX file next to input named <inputname>_with_gender.xlsx
 - Adds two columns:
     Gender            -> matched gender (Exact or fuzzy)
     Gender_MatchInfo  -> "Exact" or "Fuzzy:<match_name>:<score>"
"""

import os
import sys
import unicodedata
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from rapidfuzz import process, fuzz
from tqdm import tqdm

# Try optional unidecode for better transliteration
try:
    from unidecode import unidecode
except Exception:
    unidecode = None

# ----------------------------
# Helpers
# ----------------------------
def normalize_name(s: str, use_unidecode: bool = True) -> str:
    """
    Normalize a name for matching:
      - None -> ""
      - strip, lower
      - optional transliteration with Unidecode for non-latin scripts (if installed)
      - replace German ß with ss
      - NFKD decomposition and remove combining marks (diacritics)
      - replace hyphens/underscores with space, remove punctuation
      - collapse whitespace
    """
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = s.strip()
    if s == "":
        return ""
    # Unidecode transliteration to ASCII (optional)
    if use_unidecode and unidecode is not None:
        try:
            s = unidecode(s)
        except Exception:
            pass
    # Specific German ß handling, lowercase
    s = s.replace("ß", "ss").replace("ẞ", "ss")
    s = s.lower()
    # NFKD normalization & remove combining marks
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    # Replace hyphens and underscores and remove other punctuation
    s = s.replace("-", " ").replace("_", " ")
    s = re.sub(r"[^\w\s]", "", s)  # allow letters, numbers, underscore (we removed underscores), whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s

def try_read_csv(path, **kwargs):
    try:
        return pd.read_csv(path, **kwargs)
    except Exception:
        return pd.read_csv(path, encoding="latin1", **kwargs)

def read_input_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xls", ".xlsx"]:
        # read first sheet
        # always use openpyxl for xlsx; pandas will fallback for xls
        engine = "openpyxl" if ext == ".xlsx" else None
        return pd.read_excel(path, engine=engine)
    elif ext == ".csv":
        return try_read_csv(path)
    else:
        raise ValueError(f"Unsupported input file extension: {ext}. Supported: .csv, .xls, .xlsx")

def read_gender_source(path):
    df = try_read_csv(path)
    # find columns case-insensitively
    cols_lower = {c.lower(): c for c in df.columns}
    if 'name' not in cols_lower or 'gender' not in cols_lower:
        raise ValueError("Gender source must contain 'name' and 'gender' columns (case-insensitive).")
    name_col = cols_lower['name']
    gender_col = cols_lower['gender']
    df = df[[name_col, gender_col]].dropna(subset=[name_col])
    df.columns = ['name', 'gender']
    # normalize names
    df['name_norm'] = df['name'].astype(str).map(lambda x: normalize_name(x, use_unidecode=True))
    # keep first occurrence for duplicates
    df = df.drop_duplicates(subset=['name_norm'], keep='first')
    mapping = dict(zip(df['name_norm'], df['gender'].astype(str)))
    return mapping

# ----------------------------
# Main interactive flow
# ----------------------------
def main():
    print("=== Gender Identifier Tool (with fuzzy & progress) ===")
    print("This script uses tkinter file dialogs and runs in the terminal.\n")

    root = tk.Tk()
    root.withdraw()

    # select gender source
    print("Select the gender source CSV file (must contain 'name' and 'gender' columns).")
    gender_src_path = filedialog.askopenfilename(
        title="Select gender source CSV",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not gender_src_path:
        print("No gender source selected. Exiting.")
        sys.exit(1)
    print(f"Gender source: {gender_src_path}")

    # select input file
    print("\nSelect the input file to identify genders (CSV, XLS, or XLSX).")
    input_path = filedialog.askopenfilename(
        title="Select input file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not input_path:
        print("No input file selected. Exiting.")
        sys.exit(1)
    print(f"Input file: {input_path}")

    # read gender source
    try:
        gender_map = read_gender_source(gender_src_path)
    except Exception as e:
        print("Error reading gender source:", e)
        sys.exit(1)
    print(f"Loaded {len(gender_map)} normalized names from gender source.")

    # read input
    try:
        df = read_input_file(input_path)
    except Exception as e:
        print("Error reading input file:", e)
        sys.exit(1)

    if df.shape[0] == 0:
        print("Input file is empty. Exiting.")
        sys.exit(1)

    # show columns and prompt user to choose
    print("\nInput file columns:")
    for i, c in enumerate(df.columns):
        print(f"  [{i}] {c}")
    selected = input("\nEnter the column index or exact column name that contains the FIRST name (press Enter for column 0): ").strip()
    if selected == "":
        col_idx = 0
        chosen_col = df.columns[0]
    else:
        chosen_col = None
        try:
            idx = int(selected)
            if idx < 0 or idx >= len(df.columns):
                print("Index out of range.")
                sys.exit(1)
            chosen_col = df.columns[idx]
        except ValueError:
            if selected in df.columns:
                chosen_col = selected
            else:
                # case-insensitive match attempt
                matches = [c for c in df.columns if c.lower() == selected.lower()]
                if len(matches) == 1:
                    chosen_col = matches[0]
                else:
                    print("Column name not found. Exiting.")
                    sys.exit(1)
    print(f"Chosen column: '{chosen_col}'")

    # Prompt for fuzzy threshold
    try:
        thr_input = input("\nEnter fuzzy-match threshold (0-5000, default 85). Higher -> stricter: ").strip()
        fuzzy_threshold = int(thr_input) if thr_input != "" else 5000
        if fuzzy_threshold < 0 or fuzzy_threshold > 5000:
            print("Threshold must be 0-5000. Using default 85.")
            fuzzy_threshold = 85
    except Exception:
        fuzzy_threshold = 85
    print(f"Fuzzy threshold set to: {fuzzy_threshold}")

    # Prepare name tokens: take first token
    names_raw = df[chosen_col].astype(str).fillna("").map(lambda s: s.strip())
    first_tokens = names_raw.map(lambda x: x.split()[0] if x.strip() != "" else "")
    # normalize tokens
    normalized_tokens = first_tokens.map(lambda x: normalize_name(x, use_unidecode=True))

    # Precompute unique normalized tokens to map once
    unique_norm = sorted(set(normalized_tokens.tolist()))
    unique_norm = [u for u in unique_norm if u != ""]  # drop empty token

    # Prepare lookup list for rapidfuzz
    gender_keys = list(gender_map.keys())

    # We'll fill a dict: norm_token -> (gender, info)
    resolved = {}

    # First pass: exact matches
    exact_matches = 0
    fuzzy_matches = 0

    print("\nMapping unique names to genders (progress shown)...")
    for token in tqdm(unique_norm, desc="Mapping", unit="name"):
        # exact
        if token in gender_map:
            resolved[token] = (gender_map[token], "Exact")
            exact_matches += 1
            continue
        # fuzzy match using rapidfuzz
        if len(gender_keys) == 0:
            resolved[token] = ("Unknown", "NoReferenceData")
            continue
        match = process.extractOne(query=token, choices=gender_keys, scorer=fuzz.token_sort_ratio)
        if match is None:
            resolved[token] = ("Unknown", "NoMatch")
            continue
        best_name, score, _ = match  # score 0-100
        if score >= fuzzy_threshold:
            resolved[token] = (gender_map.get(best_name, "Unknown"), f"Fuzzy:{best_name}:{int(score)}")
            fuzzy_matches += 1
        else:
            resolved[token] = ("Unknown", f"NoGoodFuzzy:{best_name}:{int(score)}")

    # Map back to full column
    genders = []
    matchinfos = []
    for norm in normalized_tokens:
        if not norm or norm == "":
            genders.append("Unknown")
            matchinfos.append("Empty")
        else:
            g, info = resolved.get(norm, ("Unknown", "Unresolved"))
            genders.append(g)
            matchinfos.append(info)

    # Add output columns, avoid overwriting existing
    out_col = "Gender"
    info_col = "Gender_MatchInfo"
    if out_col in df.columns:
        i = 1
        while f"{out_col}_{i}" in df.columns:
            i += 1
        out_col = f"{out_col}_{i}"
        print(f"Note: 'Gender' column exists. Writing to '{out_col}' instead.")
    if info_col in df.columns:
        i = 1
        while f"{info_col}_{i}" in df.columns:
            i += 1
        info_col = f"{info_col}_{i}"

    df[out_col] = genders
    df[info_col] = matchinfos

    # Save to Excel in same folder
    input_dir = os.path.dirname(os.path.abspath(input_path))
    input_name = os.path.splitext(os.path.basename(input_path))[0]
    output_name = f"{input_name}_with_gender.xlsx"
    output_path = os.path.join(input_dir, output_name)

    try:
        print("\nWriting output Excel file...")
        df.to_excel(output_path, index=False)
    except Exception as e:
        print("Failed to write Excel output:", e)
        # fallback to CSV
        fallback = os.path.join(input_dir, f"{input_name}_with_gender.csv")
        try:
            df.to_csv(fallback, index=False)
            print(f"Wrote fallback CSV to: {fallback}")
            print("Done (CSV fallback).")
            sys.exit(0)
        except Exception as e2:
            print("Failed to write fallback CSV as well:", e2)
            sys.exit(1)

    total_rows = len(df)
    known_count = (df[out_col] != "Unknown").sum()
    unknown_count = total_rows - known_count

    print(f"\nSuccess! Output written to: {output_path}")
    print("Summary:")
    print(f"  Rows processed: {total_rows}")
    print(f"  Unique names mapped: {len(unique_norm)}")
    print(f"  Exact matches: {exact_matches}")
    print(f"  Fuzzy matches: {fuzzy_matches}")
    print(f"  Gender assigned (found): {known_count}")
    print(f"  Unknown: {unknown_count}")
    print("\nFinished.")

if __name__ == "__main__":
    main()
