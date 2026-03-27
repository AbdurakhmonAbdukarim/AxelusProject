import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
import config


def load_source(path: str) -> pd.DataFrame:
    # Read ALL columns as string first — preserves leading zeros like "0407267"
    # We manually convert numeric columns (qty etc.) ourselves after
    df = pd.read_excel(path, dtype=str)

    # Drop columns with no header (NaN or "Unnamed:..." column names)
    df = df.loc[:, df.columns.notna()]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed:")]

    # Convert all column names to string and strip whitespace
    df.columns = [str(c).strip() for c in df.columns]

    df = _normalise_columns(df)
    _check_required_columns(df)
    df = _filter_rows(df)
    print(f"  {len(df)} qualifying rows across {df[config.COL_MANUFACTURER].nunique()} manufacturer(s).")
    return df


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    variants = {
        config.COL_ITEM_NAME:    ["itemname", "item_name", "name"],
        config.COL_SKU:          ["sku", "itemcode", "item_code", "articleno"],
        config.COL_MANUFACTURER: ["manufacturer", "manufact", "brand", "vendor"],
        config.COL_ORDER_QTY:    ["zakazsoni", "zakaz_soni", "zakazqty", "orderqty", "qty", "quantity"],
    }
    rename_map = {}
    for canonical, aliases in variants.items():
        for col in df.columns:
            # col is guaranteed to be a string now
            if col.lower().replace(" ", "") in aliases:
                rename_map[col] = canonical
                break
    df.rename(columns=rename_map, inplace=True)
    return df


def _check_required_columns(df: pd.DataFrame):
    required = {config.COL_ITEM_NAME, config.COL_SKU, config.COL_MANUFACTURER, config.COL_ORDER_QTY}
    missing = required - set(df.columns)
    if missing:
        sys.exit(
            f"[ERROR] Missing columns: {missing}\n"
            f"Found columns: {list(df.columns)}"
        )


def _pad_sku(value: str) -> str:
    """
    Pad numeric SKUs to 7 digits to restore leading zeros lost by Excel.
    e.g. "407267" → "0407267"  (6 digits → pad to 7)
         "0407267" → "0407267" (already 7, untouched)
         "SKU-ABC" → "SKU-ABC" (not numeric, untouched)
    """
    s = str(value).strip()
    if s.lower() in ("nan", "", "none"):
        return s
    if s.isdigit():
        return s.zfill(7)   # pad with zeros up to 7 digits
    return s


def _filter_rows(df: pd.DataFrame) -> pd.DataFrame:
    df[config.COL_ORDER_QTY] = pd.to_numeric(df[config.COL_ORDER_QTY], errors="coerce")
    df = df[df[config.COL_ORDER_QTY].notna() & (df[config.COL_ORDER_QTY] > 0)].copy()
    # Restore leading zeros stripped by Excel's General format
    if config.COL_SKU in df.columns:
        df[config.COL_SKU] = df[config.COL_SKU].apply(_pad_sku)
    return df