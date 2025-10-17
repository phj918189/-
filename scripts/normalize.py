import pandas as pd
from typing import List


HEADER_ITEM_CANDIDATES: List[str] = ["측정항목", "항목", "분석항목"]

# 한글 컬럼명 -> 표준 컬럼명 매핑(동의어 포함)
SYNONYM_COLMAP = {
    # 식별자 후보들
    "측정번호": "sample_no",
    "시료번호": "sample_no",
    "의뢰번호": "sample_no",
    "접수번호": "sample_no",
    "계약번호": "sample_no",
    "사업장관리번호": "sample_no",

    # 그 외 정보
    "현장명": "site_name",
    "사업장": "site_name",
    "채취일시": "collected_at",
    "채취일자": "collected_at",
    "종류": "kind",
    "상태": "status",
    "측정항목": "item",
    "항목": "item",
}


def _clean_headers(headers: List[str]) -> List[str]:
    cleaned: List[str] = []
    for h in headers:
        s = str(h).strip().replace(" ", "")
        cleaned.append(s)
    return cleaned


def _promote_header_row(df: pd.DataFrame) -> pd.DataFrame:
    # 상단에 설명 행이 포함된 경우, '측정항목' 등의 키워드가 있는 행을 찾아 헤더로 승격
    for i in range(min(10, len(df))):
        row_values = [str(v).strip() for v in list(df.iloc[i].values)]
        if any(any(cand in v for cand in HEADER_ITEM_CANDIDATES) for v in row_values):
            df = df.iloc[i + 1 :].copy()
            df.columns = _clean_headers(row_values)
            return df
    # 기본: 현재 헤더 정리만 적용
    df = df.copy()
    df.columns = _clean_headers(list(df.columns))
    return df


def _apply_synonym_mapping(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for col in df.columns:
        std = SYNONYM_COLMAP.get(col)
        if std:
            rename_map[col] = std
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = _promote_header_row(df)
    df = _apply_synonym_mapping(df)

    # 유지 컬럼 구성: item은 필수, 나머지는 존재하는 것만 유지
    desired = ["sample_no", "site_name", "collected_at", "kind", "item", "status"]
    keep = [c for c in desired if c in df.columns]
    if "item" not in keep:
        raise ValueError("필수 컬럼 'item(측정항목)'을 찾을 수 없습니다.")

    df = df[keep].copy()

    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip()

    # uniq_key 생성: sample_no 있으면 sample_no+item, 없으면 site_name/collected_at 조합
    if "sample_no" in df.columns:
        df["uniq_key"] = df["sample_no"].astype(str) + "_" + df["item"].astype(str)
    else:
        parts: List[str] = []
        if "site_name" in df.columns:
            parts.append(df["site_name"].astype(str))
        if "collected_at" in df.columns:
            parts.append(df["collected_at"].astype(str))
        parts.append(df["item"].astype(str))
        if parts:
            combined = parts[0]
            for p in parts[1:]:
                combined = combined + "_" + p
            df["uniq_key"] = combined
        else:
            df["uniq_key"] = df.index.astype(str)

    # 공백/결측 item 제거
    df = df[df["item"].astype(str).str.strip() != ""].reset_index(drop=True)
    return df