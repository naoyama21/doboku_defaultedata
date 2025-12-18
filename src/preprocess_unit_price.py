import os
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd


def _to_halfwidth(text: str) -> str:
    try:
        import unicodedata
    except Exception:
        return text
    result_chars: List[str] = []
    for ch in text:
        try:
            name = unicodedata.name(ch)
            if "FULLWIDTH" in name:
                ascii_char = unicodedata.normalize("NFKC", ch)
                result_chars.append(ascii_char)
            else:
                result_chars.append(ch)
        except Exception:
            result_chars.append(ch)
    return "".join(result_chars)


def clean_cell(text: Optional[str]) -> str:
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    # Normalize line breaks and spaces
    text = text.replace("\r", " ").replace("\n", " ").strip()
    # Collapse multiple spaces
    text = re.sub(r"\s+", " ", text)
    # Normalize quotes
    text = text.strip('\'"')
    # Normalize fullwidth forms
    text = _to_halfwidth(text)
    return text


def normalize_unit(unit: str) -> str:
    unit = clean_cell(unit)
    # Drop leading '空' only when followed by m-units (m, m2, m³, etc.)
    unit = re.sub(r"^\s*空\s*(?=(m²|m2|m³|m3)\b)", "", unit)
    # Standardize common engineering units
    unit = unit.replace("m2", "m²").replace("m^2", "m²")
    unit = unit.replace("m3", "m³").replace("m^3", "m³")
    unit = unit.replace("㎡", "m²").replace("㎥", "m³")
    return unit


def split_category(cat_text: str) -> Tuple[str, str]:
    """
    Split category cell into (大分類名, 工種名).

    Rules:
      1) If the very first token (leftmost non-space sequence) ends with '工' and
         is followed by at least one space, treat that token as 大分類名 and the rest as 工種名.
         e.g. '共通工 構造物補修工  断面修復工 (左官工法)'
              → ('共通工', '構造物補修工  断面修復工 (左官工法)')
      2) Else, if there are 2+ consecutive spaces somewhere, split on the first such gap:
         e.g. '土工  安定処理工(自走式土質改良工)'
              → ('土工', '安定処理工(自走式土質改良工)')
      3) Else, return ('', whole_string)
    """
    # DO NOT collapse spaces here; raw_category passed in already preserves runs of spaces
    s = cat_text if isinstance(cat_text, str) else str(cat_text)
    s = s.replace("\r", " ").replace("\n", " ").strip().strip('\'"')

    # Rule 1: first token (leftmost non-space) + at least one following space → (大分類, 残り)
    m_lead = re.match(r"^(\S+)\s+(.*)$", s)
    if m_lead:
        first, rest = m_lead.group(1), m_lead.group(2).strip()
        # Avoid duplicated leading big-category in remainder like '共通工 共通工 ...'
        if rest.startswith(first + " "):
            rest = rest[len(first) + 1 :].strip()
        return first, rest

    # Rule 2: split on first run of 2+ spaces
    m_two = re.search(r"\s{2,}", s)
    if m_two:
        idx = m_two.start()
        left = s[:idx].strip()
        right = s[m_two.end():].strip()
        return left, right

    # Rule 3: fallback
    return "", s


def extract_table_meta(table_text: str) -> Tuple[str, str, str]:
    """
    From '自走式土質改良機設置(撤去) 1台1回当り単価表'
    return (細別名, 作業単位_数量, 作業単位_単位)
    """
    s = clean_cell(table_text)
    # Remove trailing '単価表' word for 細別名 base
    base = s
    base = re.sub(r"\s*単価表\s*$", "", base)
    # helper: convert ASCII parentheses to full-width
    def to_fullwidth_parens(text: str) -> str:
        return text.replace("(", "（").replace(")", "）")

    # Default name initializes to base (we will slice out matched unit parts)
    name = base

    # Extract angle-bracket note: <...> or ＜…＞ and normalize commas, then remove from base
    angle_note = ""
    m_angle = re.search(r"[<＜]([^>＞]+)[>＞]", base)
    if m_angle:
        angle_note = m_angle.group(1)
        # Normalize comma runs to Japanese comma
        angle_note = re.sub(r"\s*,\s*", "、", angle_note)
        angle_note = re.sub(r"、{2,}", "、", angle_note).strip("、 ").strip()
        # Remove the bracketed segment from base to not interfere unit extraction
        base = (base[:m_angle.start()] + base[m_angle.end():]).strip()
        name = base
    unit_qty = ""
    unit_unit = ""
    # Pattern A: combined 'N <unitA> N2 回 当り/当たり' (e.g., '1基1回当り', '1台1回当り')
    m_combo = re.search(r"([0-9０-９,〇○]+)\s*(台|基|本|枚|個|ケーブル|ブロック)\s*([0-9０-９,〇○]+)\s*(回)\s*(当り|当たり)", base)
    if m_combo:
        unit_qty = m_combo.group(1)
        unitA = m_combo.group(2)
        num2 = clean_cell(m_combo.group(3)).replace(",", "")
        unit_unit = f"{unitA}{num2}回"
        # remove matched segment from name
        span = m_combo.span()
        name = (base[:span[0]] + base[span[1]:]).strip()
    else:
        # Pattern B: 'N <unit>(optional_annotation) 当り/当たり'
        unit_core = r"(?:空?\s*(?:掛?\s*(?:m²|m2)|m³|m3)|km|m|本|基|構造物|箇所|袋|t|台|日|h|時間|車|式|箇月|月|工事|径間|組|ケーブル|枚|個|穴|孔|橋|トンネル|ブロック)"
        m_simple = re.search(
            r"([0-9０-９,〇○]+)\s*(" + unit_core + r")(\s*\([^)]*\))?\s*(当り|当たり)",
            base,
            flags=re.IGNORECASE,
        )
        if m_simple:
            unit_qty = m_simple.group(1)
            unit_core_val = m_simple.group(2)
            unit_annotation = m_simple.group(3) or ""
            unit_unit = normalize_unit(unit_core_val)
            # remove matched segment from name
            span = m_simple.span()
            name = (base[:span[0]] + base[span[1]:]).strip()
            # append annotation to name (as full-width parentheses), not to unit
            if unit_annotation.strip():
                name = (name + to_fullwidth_parens(unit_annotation)).strip()
        else:
            # Pattern C: explicit fallback '1<unitA>1回当り' without spaces
            m_fallback = re.search(r"1\s*(台|基|本|枚|個|ケーブル)\s*1\s*回\s*(当り|当たり)", base)
            if m_fallback:
                unit_qty = "1"
                unit_unit = f"{m_fallback.group(1)}1回"
                span = m_fallback.span()
                name = (base[:span[0]] + base[span[1]:]).strip()
            else:
                # Pattern D: 'N 回 当り/当たり' alone (e.g., '1回当り')
                m_only_times = re.search(r"([0-9０-９,〇○]+)\s*回\s*(当り|当たり)", base)
                if m_only_times:
                    unit_qty = m_only_times.group(1)
                    unit_unit = "回"
                    span = m_only_times.span()
                    name = (base[:span[0]] + base[span[1]:]).strip()

    name = name.strip()
    if angle_note:
        name = f"{name}（{angle_note}）".strip()
    unit_qty = clean_cell(unit_qty).replace(",", "")
    unit_unit = normalize_unit(unit_unit)
    return name, unit_qty, unit_unit


def read_unit_price_csv(csv_path: str) -> pd.DataFrame:
    """
    Read unit_price_table_data.csv without header.
    Expected columns:
      0: 大分類/工種を含むテキスト
      1: 細別＋単価表テキスト
      2..: '名称','規格','単位','数量','摘要' またはデータ
    """
    # Many lines contain commas inside quoted fields (e.g., 1,000mm).
    # Use csv.reader to respect quotes; still fall back to join tail into c7 when field count exceeds 7.
    import csv

    rows: List[List[str]] = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=",", quotechar='"')
        for parts in reader:
            if not parts or not any((p or "").strip() for p in parts):
                continue
            if len(parts) < 7:
                parts = parts + [""] * (7 - len(parts))
            elif len(parts) > 7:
                head = parts[:6]
                tail = ",".join(parts[6:])
                parts = head + [tail]
            rows.append(parts[:7])

    df = pd.DataFrame(rows, columns=["raw_category", "raw_table", "c3", "c4", "c5", "c6", "c7"]).fillna("")

    # Preserve double spaces in raw_category to enable stable split by \s{2,}
    def clean_preserve_spaces(text: Optional[str]) -> str:
        if text is None:
            return ""
        if not isinstance(text, str):
            text = str(text)
        text = text.replace("\r", " ").replace("\n", " ").strip()
        text = text.strip('\'"')
        return _to_halfwidth(text)

    df["raw_category"] = df["raw_category"].apply(clean_preserve_spaces)
    for col in ["raw_table", "c3", "c4", "c5", "c6", "c7"]:
        df[col] = df[col].apply(clean_cell)
    return df


def normalize_unit_price_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filter out header/total and machine-rate tables; map to normalized columns.
    Output columns:
      大分類名, 工種名, 細別名, 基本歩掛名,
      所要日数作業単位_数量, 所要日数作業単位_単位,
      歩掛作業単位_数量, 歩掛作業単位_単位,
      名称, 規格, 単位, 数量, 摘要
    """
    if df.empty:
        return df

    # Skip table headers (名称, 規格, 単位, 数量, 摘要)
    is_header = (df["c3"] == "名称") & (df["c4"] == "規格")
    # Skip totals '計'
    is_total = df["c3"] == "計"
    # Skip machine rate '機械運転' block (structure differs)
    is_machine_block = df["raw_table"].str.contains("機械運転", na=False)

    data_mask = (~is_header) & (~is_total) & (~is_machine_block) & (df["c3"] != "")
    d = df.loc[data_mask].copy()

    # Extract category
    d[["大分類名", "工種名"]] = d["raw_category"].apply(lambda s: pd.Series(split_category(s)))
    # Extract table meta
    tbl_meta = d["raw_table"].apply(lambda s: pd.Series(extract_table_meta(s)))
    tbl_meta.columns = ["細別名", "_作業単位_数量", "_作業単位_単位"]
    d = pd.concat([d, tbl_meta], axis=1)

    d["_作業単位_単位"] = d["_作業単位_単位"].apply(normalize_unit)

    # Map core columns
    d = d.rename(columns={"c3": "名称", "c4": "規格", "c5": "単位", "c6": "数量", "c7": "摘要"})

    # Fill 歩掛作業単位_* with extracted units; 所要日数作業単位_* は空欄（要望）
    d["所要日数作業単位_数量"] = ""
    d["所要日数作業単位_単位"] = ""
    d["歩掛作業単位_数量"] = d["_作業単位_数量"]
    d["歩掛作業単位_単位"] = d["_作業単位_単位"]

    # 基本歩掛名は空欄（要望）
    d["基本歩掛名"] = ""

    # Reorder/select columns
    cols = [
        "大分類名",
        "工種名",
        "細別名",
        "基本歩掛名",
        "所要日数作業単位_数量",
        "所要日数作業単位_単位",
        "歩掛作業単位_数量",
        "歩掛作業単位_単位",
        "名称",
        "規格",
        "単位",
        "数量",
        "摘要",
    ]
    d = d[cols]

    # Remove '単価表' token anywhere in text columns; preserve numbers around it
    text_cols = ["大分類名", "工種名", "細別名", "名称", "規格", "摘要"]
    for col in text_cols:
        if col in d.columns:
            d[col] = (
                d[col]
                .astype(str)
                .str.replace(r"\s*単価表\s*", " ", regex=True)
                .str.replace(r"\s+", " ", regex=True)
                .str.strip()
            )
    return d


def load_and_normalize_unit_price(csv_path: str) -> pd.DataFrame:
    df_raw = read_unit_price_csv(csv_path)
    return normalize_unit_price_rows(df_raw)


def normalize_table_data_for_aux(csv_path: str) -> pd.DataFrame:
    """
    Minimal normalization for table_data.csv to use as auxiliary source (notes/complements).
    Strategy: The source has inconsistent CSV structure with many commas.
    For auxiliary purposes, we keep each raw line as a single string column.
    """
    lines: List[str] = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        for raw_line in f:
            line = raw_line.rstrip("\r\n")
            if not line:
                continue
            lines.append(clean_cell(line))
    return pd.DataFrame({"raw": lines})


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.normpath(os.path.join(base_dir, "..", "data"))
    up_csv = os.path.join(data_dir, "unit_price_table_data.csv")
    td_csv = os.path.join(data_dir, "table_data.csv")

    norm_df = load_and_normalize_unit_price(up_csv)
    out_dir = os.path.join(data_dir, "normalized")
    os.makedirs(out_dir, exist_ok=True)
    norm_df.to_csv(os.path.join(out_dir, "unit_price_normalized.csv"), index=False, encoding="utf-8")

    td_df = normalize_table_data_for_aux(td_csv)
    td_df.to_csv(os.path.join(out_dir, "table_data_aux.csv"), index=False, encoding="utf-8")


