import pandas as pd
from datetime import datetime

# ---------------- Normalization ----------------

def normalize(x):
    if pd.isna(x):
        return ""
    return str(x).strip().replace("\u200b", "").replace("\n", "").replace("\t", "")

def norm_lower(x):
    return normalize(x).lower()

def normalize_col_name(c):
    return (
        str(c)
        .strip()
        .lower()
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
        .replace("  ", " ")
    )

def find_column(df, target_name):
    target = normalize_col_name(target_name)
    for c in df.columns:
        if normalize_col_name(c) == target:
            return c
    return None

# ---------------- Validation helpers ----------------

def is_valid_date(val, min_date="2024-01-01", max_date=None):
    if max_date is None:
        max_date = datetime.today().strftime("%Y-%m-%d")
    d = pd.to_datetime(val, errors="coerce", dayfirst=True)
    if pd.isna(d):
        return False
    return pd.Timestamp(min_date) <= d <= pd.Timestamp(max_date)

def is_numeric(val, min_v=0, max_v=100):
    try:
        v = float(val)
        return min_v <= v <= max_v
    except Exception:
        return False

# ---------------- Date cleaner ----------------

def clean_dates(df):
    if df is None or df.empty:
        return df

    for col in df.columns:
        s = df[col]
        if pd.api.types.is_numeric_dtype(s):
            continue
        if pd.api.types.is_datetime64_any_dtype(s) or "date" in col.lower():
            dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
            df[col] = dt.dt.strftime("%Y-%m-%d").fillna("")
    return df
