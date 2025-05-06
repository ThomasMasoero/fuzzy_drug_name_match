import pandas as pd 
import re
from rapidfuzz import process
from pathlib import Path

# 0) Get the folder where this script is located
base_dir = Path(__file__).parent.resolve()

# 1) Configurable list of separators
separators = ['+', ',', 'vs', 'versus', '/', ':', " e "]
sep_pattern = r'\s*(?:' + '|'.join(map(re.escape, separators)) + r')\s*'
sep_regex = re.compile(sep_pattern, flags=re.IGNORECASE)

# 2) Load reference drug list (generic)
ref_df = pd.read_excel(base_dir / 'ATC_DDD_Index ITA.xlsx')   # 'Nome_L5', 'Codice ATC_L5'
ref_df.columns = ref_df.columns.str.strip()
ref_df = ref_df.rename(columns={'Nome_L5': 'DrugName', 'Codice ATC_L5': 'ATC'})

# 3) Load commercial drug names
ref_com_df = pd.read_excel(base_dir / 'commercial_name.xlsx')  # 'MARCHIO', 'CODICE_ATC', 'PRINCIPIO_ATTIVO'
ref_com_df.columns = ref_com_df.columns.str.strip()
ref_com_df = ref_com_df.rename(columns={'MARCHIO': 'DrugName', 'CODICE_ATC': 'ATC'})

# 4) Combine both into a unified reference list
combined_ref = pd.concat([ref_df, ref_com_df[['DrugName', 'ATC']]], ignore_index=True)
combined_ref['DrugName'] = combined_ref['DrugName'].str.lower().str.strip()
combined_ref.drop_duplicates(subset='DrugName', inplace=True)

# 5) Load and clean CRF data
crf_df = pd.read_excel(base_dir / 'terapie_altro - Copia.xlsx')   # 'altro_farmaco'
crf_df.columns = crf_df.columns.str.strip()
crf_df['altro_farmaco'] = (
    crf_df['altro_farmaco']
      .astype(str)
      .str.split(sep_regex)
)
crf_df = crf_df.explode('altro_farmaco').reset_index(drop=True)
crf_df['altro_farmaco'] = crf_df['altro_farmaco'].str.lower().str.strip()

# 6) Prepare name list for fuzzy matching
reference_names = combined_ref['DrugName'].tolist()

# 7) Matching function
def match_drug_name(drug):
    match = process.extractOne(drug, reference_names, score_cutoff=80)
    if match:
        matched_name, score, idx = match
        atc_code = combined_ref.iloc[idx]['ATC']
        return matched_name, atc_code, score
    return None, None, None

# 8) Apply matching
matches = crf_df['altro_farmaco'].apply(match_drug_name)
crf_df[['MatchedDrug', 'ATCCode', 'Score']] = pd.DataFrame(matches.tolist(), index=crf_df.index)

# 9) Save results in the same directory
crf_df.to_excel(base_dir / 'matched_results.xlsx', index=False)
