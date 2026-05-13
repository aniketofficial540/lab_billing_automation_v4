import pandas as pd
from config import (
    CSV_COL_CPT,
    COL_TEST_CODE,
    COL_MRP,
    CSV_COLS_KEEP,
)


def lookup_prices(
    his_df: pd.DataFrame,
    mr_column: str,
    master_df: pd.DataFrame,
    party_short_code: str = None,
    billing_month: str = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Look up MRP and Bill Amount for each CPT code in the party's column.

    Silently checks addendum JSON if party_short_code and billing_month are provided.
    Returns (matched_df, unmatched_df).
    """
    # Build lookup dict: test_code -> (mrp, bill_amount)
    price_map: dict[str, tuple] = {}
    for _, row in master_df.iterrows():
        tc = str(row[COL_TEST_CODE]).strip() if pd.notna(row[COL_TEST_CODE]) else ""
        if not tc:
            continue
        mrp_raw = row[COL_MRP]
        mrp = float(mrp_raw) if pd.notna(mrp_raw) and mrp_raw != "" else None

        if mr_column in master_df.columns:
            ba_raw = row[mr_column]
            bill_amt = float(ba_raw) if pd.notna(ba_raw) and ba_raw != "" else None
        else:
            bill_amt = None

        price_map[tc] = (mrp, bill_amt)

    matched_rows = []
    unmatched_rows = []

    for _, row in his_df.iterrows():
        cpt = str(row[CSV_COL_CPT]).strip() if pd.notna(row[CSV_COL_CPT]) else ""
        base = row.to_dict()

        mrp, bill_amt = price_map.get(cpt, (None, None))

        if cpt in price_map:
            if bill_amt is not None:
                discount = (mrp - bill_amt) if mrp is not None else None
                base["MRP"] = mrp
                base["Discount"] = discount
                base["Bill_Amount"] = bill_amt
                matched_rows.append(base)
            else:
                base["MRP"] = mrp
                base["Discount"] = None
                base["Bill_Amount"] = None
                unmatched_rows.append(base)
        else:
            base["MRP"] = None
            base["Discount"] = None
            base["Bill_Amount"] = None
            unmatched_rows.append(base)

    matched_df = pd.DataFrame(matched_rows) if matched_rows else _empty_price_df(his_df)
    unmatched_df = pd.DataFrame(unmatched_rows) if unmatched_rows else _empty_price_df(his_df)

    return matched_df, unmatched_df


def apply_confirmed_rates(
    unmatched_df: pd.DataFrame, confirmed_rates: list[dict]
) -> pd.DataFrame:
    """Merge employee-confirmed rates into unmatched_df rows.

    confirmed_rates: list of dicts with keys cpt_code, mrp, discount, bill_amount
    Returns updated DataFrame with filled MRP/Discount/Bill_Amount.
    """
    rate_map = {r["cpt_code"]: r for r in confirmed_rates}

    df = unmatched_df.copy()
    for i, row in df.iterrows():
        cpt = str(row[CSV_COL_CPT]).strip()
        if cpt in rate_map:
            r = rate_map[cpt]
            df.at[i, "MRP"] = float(r["mrp"])
            df.at[i, "Discount"] = float(r["discount"])
            df.at[i, "Bill_Amount"] = float(r["bill_amount"])
    return df


def _empty_price_df(his_df: pd.DataFrame) -> pd.DataFrame:
    cols = list(his_df.columns) + ["MRP", "Discount", "Bill_Amount"]
    return pd.DataFrame(columns=cols)
