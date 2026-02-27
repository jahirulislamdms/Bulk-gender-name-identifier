# 🧠 Smart Gender Name Identifier

A powerful terminal-based Python tool that maps first names to genders using a reference CSV file.

Supports fuzzy matching, Unicode normalization, Excel/CSV input, and automatic Excel output generation.

---

## 🚀 Features

- ✅ Exact normalized name matching
- ✅ Fuzzy matching using RapidFuzz
- ✅ Unicode normalization (NFKD)
- ✅ Diacritic removal
- ✅ Optional Unidecode transliteration support
- ✅ Progress bars (tqdm)
- ✅ Input support: CSV, XLS, XLSX
- ✅ Output: Excel file with gender columns added
- ✅ Handles multilingual names

---

## 📦 Requirements

- Python 3.8+
- pandas
- rapidfuzz
- tqdm
- openpyxl

Optional (recommended for better multilingual support):

- unidecode

Install dependencies:

```bash
pip install pandas rapidfuzz tqdm openpyxl unidecode
```

---

## 📂 Required Gender Source File

The gender reference file must be a CSV containing:

```
name,gender
John,Male
Emily,Female
Alex,Unisex
```

Column names must include:
- `name`
- `gender`

(Case-insensitive)

---

## ▶️ How to Run

```bash
python Bulk-gender-name-identifier.py
```

### Step-by-step:

1. Select the gender reference CSV file.
2. Select the input file (CSV, XLS, or XLSX).
3. Choose the column containing the first names.
4. Set fuzzy match threshold (default recommended: 85).
5. Wait for processing to complete.

---

## 📁 Output

The script generates a new Excel file in the same folder as the input:

```
<inputname>_with_gender.xlsx
```

### Two new columns are added:

- `Gender`
- `Gender_MatchInfo`

---

## 🔎 Matching Logic

1. Normalize names:
   - Lowercase
   - Remove accents/diacritics
   - Remove punctuation
   - Transliterate (if unidecode installed)

2. Try exact match

3. If no exact match → fuzzy match using RapidFuzz

4. If fuzzy score meets threshold → assign gender

5. Otherwise → mark as `Unknown`

---

## 📊 Output Example

| FirstName | Gender | Gender_MatchInfo |
|-----------|--------|------------------|
| John      | Male   | Exact            |
| Emely     | Female | Fuzzy:emily:92   |
| Xyzabc    | Unknown| NoGoodFuzzy:...  |

---

## ⚙️ Fuzzy Threshold

- Range: 0–5000  
- Recommended: 85  
- Higher value = stricter matching

---

## 💡 Use Cases

- CRM data cleaning
- Lead enrichment
- Marketing segmentation
- Customer profiling
- Data preprocessing

---

## 🛡️ Privacy

This tool runs completely offline.  
No data is sent to external servers.

---

## 📜 License

MIT License
