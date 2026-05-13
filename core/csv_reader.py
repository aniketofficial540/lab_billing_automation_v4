import os
import pandas as pd
from config import CSV_COLS_KEEP, CSV_COL_CPT, COOP_CODE
from core.logger import get_logger

log = get_logger(__name__)

# Try the most common HIS export encodings in priority order. The first one
# that decodes the file cleanly wins and is logged. latin-1 is the universal
# fallback (it never raises UnicodeDecodeError) so it sits last.
_CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "windows-1252", "latin-1")
_EXCEL_EXTS = (".xlsx", ".xls")


def read_his_export(path: str) -> pd.DataFrame:
    """Load HIS export (Excel or CSV), validate required columns, drop COOP0322.

    .attrs keys: coop_removed, billable_rows, total_before, source
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in _EXCEL_EXTS:
        try:
            df = pd.read_excel(path, dtype=str)
        except Exception as e:
            raise ValueError(f"Cannot read Excel file: {e}")
        source = "excel"

        # Strip column names of surrounding whitespace
        df.columns = [c.strip() for c in df.columns]

        # Validate required columns
        missing = [c for c in CSV_COLS_KEEP if c not in df.columns]
        if missing:
            raise ValueError(
                f"HIS export is missing required columns: {', '.join(missing)}. "
                "Please check you uploaded the correct file."
            )

        if df.empty:
            raise ValueError("HIS export is empty — no data rows found.")

        total_before = len(df)

        # Drop COOP0322 rows
        coop_mask = df[CSV_COL_CPT].str.strip() == COOP_CODE
        coop_removed = int(coop_mask.sum())
        df = df[~coop_mask].copy()

        if df.empty:
            raise ValueError(
                f"After removing {coop_removed} COOP0322 rows, no billable rows remain."
            )

        # Keep only the 8 required columns
        df = df[CSV_COLS_KEEP].copy()

        # Strip string values
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()

        # Drop rows with blank/invalid CPT codes (cancelled tests often appear this way)
        cpt_series = df[CSV_COL_CPT].astype(str).str.strip().str.lower()
        blank_cpt_mask = cpt_series.isin(["", "nan", "none", "null"])
        blank_cpt_removed = int(blank_cpt_mask.sum())
        if blank_cpt_removed:
            df = df[~blank_cpt_mask].copy()
            if df.empty:
                raise ValueError(
                    f"After removing {blank_cpt_removed} blank CPT row(s), no billable rows remain."
                )

        billable_rows = len(df)

        df.attrs["coop_removed"] = coop_removed
        df.attrs["billable_rows"] = billable_rows
        df.attrs["blank_cpt_removed"] = blank_cpt_removed
        df.attrs["total_before"] = total_before
        df.attrs["source"] = source

        log.info("HIS export loaded: source=%s, rows=%d, path=%s", source, len(df), path)
        return df

    df, source = _read_csv_with_fallback(path)
    df.attrs["source"] = source
    return df


def read_his_csv(path: str) -> pd.DataFrame:
    """Backward-compatible wrapper for CSV-only use cases."""
    df, source = _read_csv_with_fallback(path)
    df.attrs["source"] = source
    return df


def _read_csv_with_fallback(path: str) -> tuple[pd.DataFrame, str]:
    df = None
    chosen_enc = None
    last_decode_err = None
    for enc in _CSV_ENCODINGS:
        try:
            df = pd.read_csv(path, dtype=str, encoding=enc)
            chosen_enc = enc
            break
        except UnicodeDecodeError as e:
            last_decode_err = e
            continue
        except Exception as e:
            raise ValueError(f"Cannot read CSV file: {e}")

    if df is None:
        raise ValueError(
            f"Cannot read CSV file: tried {_CSV_ENCODINGS} — last error: {last_decode_err}"
        )

    # Strip column names of surrounding whitespace
    df.columns = [c.strip() for c in df.columns]

    # Validate required columns
    missing = [c for c in CSV_COLS_KEEP if c not in df.columns]
    if missing:
        raise ValueError(
            f"HIS CSV is missing required columns: {', '.join(missing)}. "
            "Please check you uploaded the correct file."
        )

    if df.empty:
        raise ValueError("HIS CSV file is empty — no data rows found.")

    total_before = len(df)

    # Drop COOP0322 rows
    coop_mask = df[CSV_COL_CPT].str.strip() == COOP_CODE
    coop_removed = int(coop_mask.sum())
    df = df[~coop_mask].copy()

    if df.empty:
        raise ValueError(
            f"After removing {coop_removed} COOP0322 rows, no billable rows remain."
        )

    # Keep only the 8 required columns
    df = df[CSV_COLS_KEEP].copy()

    # Strip string values
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    # Drop rows with blank/invalid CPT codes (cancelled tests often appear this way)
    cpt_series = df[CSV_COL_CPT].astype(str).str.strip().str.lower()
    blank_cpt_mask = cpt_series.isin(["", "nan", "none", "null"])
    blank_cpt_removed = int(blank_cpt_mask.sum())
    if blank_cpt_removed:
        df = df[~blank_cpt_mask].copy()
        if df.empty:
            raise ValueError(
                f"After removing {blank_cpt_removed} blank CPT row(s), no billable rows remain."
            )

    billable_rows = len(df)

    df.attrs["coop_removed"] = coop_removed
    df.attrs["billable_rows"] = billable_rows
    df.attrs["blank_cpt_removed"] = blank_cpt_removed
    df.attrs["total_before"] = total_before
    df.attrs["encoding_used"] = chosen_enc

    log.info("CSV opened with encoding=%s, rows=%d, path=%s", chosen_enc, len(df), path)
    return df, "csv"
