import re
import pandas as pd
from config import PC_HIS_NAME, PC_FULL_NAME, PARTY_COL_MAP


def _normalise(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())


def _word_set(s: str) -> set:
    return set(re.findall(r"[a-z0-9]+", s.lower()))


_HIS_SUFFIXES = [" - fbd", "-fbd", " - faridabad", " faridabad"]


def _strip_suffixes(name: str) -> str:
    lower = name.lower()
    for suf in _HIS_SUFFIXES:
        if lower.endswith(suf):
            return name[: len(name) - len(suf)].strip()
    return name


def _collapse_spaces(s: str) -> str:
    """Strip ALL whitespace, lowercase. 'Child care Centre' == 'Childcare centre'."""
    return re.sub(r"\s+", "", s).lower()


def match_party(company_name: str, pc_df: pd.DataFrame) -> pd.Series | None:
    """Return Party_Config row matching company_name, or None if no match found.

    5-pass:
      1. exact case-insensitive
      2. suffix-stripped case-insensitive
      3. partial containment
      4. space-collapsed exact (handles "Childcare" vs "Child care")
      5. word-set overlap (≥ 0.75)
    """
    if not company_name:
        return None

    cn_norm = _normalise(company_name)

    # Pass 1 — exact case-insensitive
    for _, row in pc_df.iterrows():
        if _normalise(str(row[PC_HIS_NAME])) == cn_norm:
            return row

    # Pass 2 — strip common HIS suffixes then exact CI
    cn_stripped = _normalise(_strip_suffixes(company_name))
    for _, row in pc_df.iterrows():
        party_stripped = _normalise(_strip_suffixes(str(row[PC_HIS_NAME])))
        if party_stripped == cn_stripped:
            return row

    # Pass 3 — partial containment (one name inside the other)
    for _, row in pc_df.iterrows():
        p_norm = _normalise(str(row[PC_HIS_NAME]))
        if cn_norm in p_norm or p_norm in cn_norm:
            return row

    # Pass 4 — collapse all whitespace and compare (with and without HIS suffix).
    # Catches "Kalra Childcare Centre - FBD" vs "Kalra Child care Centre".
    cn_collapsed_full    = _collapse_spaces(company_name)
    cn_collapsed_stripped = _collapse_spaces(_strip_suffixes(company_name))
    for _, row in pc_df.iterrows():
        raw   = str(row[PC_HIS_NAME])
        c_raw = _collapse_spaces(raw)
        c_str = _collapse_spaces(_strip_suffixes(raw))
        if c_raw in (cn_collapsed_full, cn_collapsed_stripped):
            return row
        if c_str and c_str in (cn_collapsed_full, cn_collapsed_stripped):
            return row
        # Also handle the reverse direction (party name longer than CSV name)
        if cn_collapsed_stripped and cn_collapsed_stripped in c_raw:
            return row

    # Pass 5 — word-set overlap fallback for genuinely-different spellings
    cn_words = _word_set(company_name)
    if not cn_words:
        return None
    best_score = 0
    best_row = None
    for _, row in pc_df.iterrows():
        p_words = _word_set(str(row[PC_HIS_NAME]))
        if not p_words:
            continue
        overlap = len(cn_words & p_words) / max(len(cn_words), len(p_words))
        if overlap > best_score and overlap >= 0.75:
            best_score = overlap
            best_row = row
    return best_row


def resolve_mr_column(his_name: str, master_df: pd.DataFrame) -> str | None:
    """Return the Master Rate Sheet column header for a party's HIS Sheet Name.

    Resolution order:
      1. Direct stripped match
      2. PARTY_COL_MAP explicit entries
      3. Case-insensitive exact match
      4. Strip HIS suffixes (e.g. ' - FBD') then direct + case-insensitive
      5. Word-set overlap >= 0.65 (last resort)
    """
    from config import COL_TEST_CODE, COL_TEST_NAME, COL_MRP
    non_party = {COL_TEST_CODE, COL_TEST_NAME, COL_MRP}
    mr_cols = [c for c in master_df.columns if c not in non_party]

    his_stripped = his_name.strip()

    # Pass 1 — direct match
    if his_stripped in mr_cols:
        return his_stripped

    # Pass 2 — PARTY_COL_MAP explicit entries
    if his_stripped in PARTY_COL_MAP:
        mapped = PARTY_COL_MAP[his_stripped]
        if mapped in mr_cols:
            return mapped

    # Pass 3 — case-insensitive exact
    his_lower = his_stripped.lower()
    for col in mr_cols:
        if col.strip().lower() == his_lower:
            return col

    # Pass 4 — strip HIS suffixes then try direct + case-insensitive
    his_no_suffix = _strip_suffixes(his_stripped)
    if his_no_suffix != his_stripped:
        if his_no_suffix in mr_cols:
            return his_no_suffix
        his_no_suffix_lower = his_no_suffix.lower()
        for col in mr_cols:
            if col.strip().lower() == his_no_suffix_lower:
                return col

    # Pass 5 — space-normalised match ("Childcare" == "Child care", "S.K.G." == "S.K.G")
    his_no_space = re.sub(r"\s+", "", his_no_suffix).lower()
    for col in mr_cols:
        if re.sub(r"\s+", "", col).lower() == his_no_space:
            return col

    # Pass 6 — word-set overlap >= 0.65 (last resort for minor wording differences)
    his_words = _word_set(his_no_suffix)
    best_score = 0.0
    best_col = None
    for col in mr_cols:
        col_words = _word_set(col)
        if not col_words:
            continue
        overlap = len(his_words & col_words) / max(len(his_words), len(col_words))
        if overlap > best_score and overlap >= 0.65:
            best_score = overlap
            best_col = col
    return best_col


def get_all_party_names(pc_df: pd.DataFrame) -> list[str]:
    """Return list of all HIS Sheet Names from Party_Config."""
    return pc_df[PC_HIS_NAME].dropna().astype(str).tolist()
