import pandas as pd
from config import CSV_COLS_KEEP, CSV_COL_CPT, COOP_CODE


def pre_flight(df: pd.DataFrame) -> None:
    """Check DataFrame has exactly 8 cols, no COOP0322, no null CPTs.

    Raises ValueError with plain-English message on failure.
    """
    # No COOP0322 rows should remain
    if CSV_COL_CPT in df.columns:
        coop_remaining = (df[CSV_COL_CPT].str.strip() == COOP_CODE).sum()
        if coop_remaining > 0:
            raise ValueError(
                f"Internal error: {coop_remaining} COOP0322 row(s) were not removed. "
                "Please re-upload the export file."
            )

    # Required columns present
    missing = [c for c in CSV_COLS_KEEP if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required columns after processing: {', '.join(missing)}."
        )

    # No blank/invalid CPT codes
    if CSV_COL_CPT in df.columns:
        cpt_series = df[CSV_COL_CPT].astype(str).str.strip().str.lower()
        blank_cpts = cpt_series.isin(["", "nan", "none", "null"]).sum()
        if blank_cpts > 0:
            raise ValueError(
                f"{int(blank_cpts)} row(s) have a blank CPT code. "
                "Please check the HIS export for data gaps."
            )


def verify_bill_amounts(matched_df: pd.DataFrame) -> list[str]:
    """Verify K = I − J (±0.01 tolerance) for every matched row.

    Returns list of error strings (empty means all pass).
    """
    errors = []
    for _, row in matched_df.iterrows():
        mrp = row.get("MRP")
        discount = row.get("Discount")
        bill = row.get("Bill_Amount")

        if mrp is None or discount is None or bill is None:
            continue

        try:
            expected = float(mrp) - float(discount)
            actual = float(bill)
            if abs(expected - actual) > 0.01:
                cpt = row.get("cpt_code", "?")
                svc = row.get("Service Name", "?")
                errors.append(
                    f"CPT {cpt} ({svc}): MRP {mrp} − Discount {discount} = {expected:.2f} "
                    f"but Bill Amount is {actual:.2f}"
                )
        except (TypeError, ValueError):
            cpt = row.get("cpt_code", "?")
            errors.append(f"CPT {cpt}: non-numeric value in MRP/Discount/Bill Amount")

    return errors


def validate_unmatched_rates(confirmed_rates: list[dict]) -> list[str]:
    """Validate confirmed rates: MRP and Bill Amount are numeric, non-negative, bill <= mrp.

    confirmed_rates: list of dicts with keys cpt_code, mrp, discount, bill_amount
    (discount is auto-calculated as mrp - bill_amount by the UI)
    Returns list of error strings (empty means all valid).
    """
    errors = []
    for r in confirmed_rates:
        cpt = r.get("cpt_code", "?")

        for field in ("mrp", "bill_amount"):
            val = r.get(field)
            if val is None or str(val).strip() == "":
                errors.append(f"CPT {cpt}: '{field}' is blank — both MRP and Bill Amount are mandatory.")
                break

        try:
            mrp = float(r["mrp"])
            bill = float(r["bill_amount"])
        except (TypeError, ValueError, KeyError):
            errors.append(f"CPT {cpt}: MRP and Bill Amount must be numbers.")
            continue

        if mrp < 0 or bill < 0:
            errors.append(f"CPT {cpt}: amounts cannot be negative.")
            continue

        if bill > mrp:
            errors.append(f"CPT {cpt}: Bill Amount ({bill}) cannot exceed MRP ({mrp}).")

    return errors
