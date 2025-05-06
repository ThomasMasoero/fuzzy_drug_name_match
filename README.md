
# 💊 Drug Name to ATC Code Matcher

This Python program is designed to clean and standardize drug names recorded in clinical case report forms (CRFs) by matching them to their corresponding **ATC (Anatomical Therapeutic Chemical)** codes. It supports both **generic** and **commercial** drug names and uses fuzzy string matching to handle typos and variations.

## 📂 Input Files

Place the following Excel files in the **same folder as the script**:

1. **`ATC_DDD_Index ITA.xlsx`**  
   - Contains generic drug names and their corresponding ATC codes.  
   - Required columns:  
     - `Nome_L5`: Generic drug name  
     - `Codice ATC_L5`: ATC code  

2. **`commercial_name.xlsx`**  
   - Contains commercial (brand) names of drugs linked to ATC codes.  
   - Required columns:  
     - `MARCHIO`: Brand name of the drug  
     - `CODICE_ATC`: ATC code  
     - `PRINCIPIO_ATTIVO`: Active substance (optional for this script)  

3. **`terapie_altro.xlsx`**  
   - Contains clinician-entered drug names from the CRFs.  
   - Required column:  
     - `altro_farmaco`: Free-text drug names (may include typos, combinations, or separators like "+", "/", "vs", etc.)

## 🧠 Key Features

- **Flexible Separators:** Handles multiple drugs in a single field separated by `+`, `/`, `,`, `vs`, `e`, etc.
- **Fuzzy Matching:** Uses `rapidfuzz` to tolerate typos and non-standard spellings.
- **Generic + Commercial Matching:** Matches against both ATC generic names and brand names.
- **Case and Whitespace Normalization:** Cleans and standardizes all text inputs.
- **Portable:** Automatically detects its own directory to load and save files (no hard-coded paths).

## 📤 Output

- **`matched_results.xlsx`**  
  Excel file containing cleaned and matched drug entries.  
  Output columns:
  - `altro_farmaco`: Original entry (after splitting and cleaning)
  - `MatchedDrug`: Closest matched reference name
  - `ATCCode`: Corresponding ATC code
  - `Score`: Matching confidence score (0–100)

## 🛠️ Dependencies

Install required packages with:

```bash
pip install pandas openpyxl rapidfuzz
```

## 🚀 How to Run

1. Place the script and all input files in the same folder.
2. Run the script with Python 3:

```bash
python match_drug_names.py
```
