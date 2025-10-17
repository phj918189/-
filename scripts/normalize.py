
from locale import normalize
from nt import rename
import pandas as pd

COLMAP = {
    "시료번호": "sample_no",
    "현장명": "site_name",
    "채취일시": "collected_at",
    "종류": "kind",
    "측정항목": "item",
    "상태": "status",
}

def normalize(df: pd.DataFrame) -> pd.DataFrame:
    rename = {k: v for k, v in COLMAP.items() if k in df.columns}
    df = df.rename(columns=rename)
    keep = [c for c in ["sample_no", "site_name", "collected_at", "kind", "item", "status"] if c in df.columns]
    df = df[keep].copy()
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].str.strip()
    df["uniq_key"] = df["sample_no"].astype(str) + "_" + df["item"].astype(str)
    return df