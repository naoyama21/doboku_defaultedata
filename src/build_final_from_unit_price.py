import os
import re
from typing import Dict, List, Optional

import pandas as pd

from preprocess_unit_price import (
    load_and_normalize_unit_price,
    normalize_table_data_for_aux,
    clean_cell,
    normalize_unit,
)


def load_category_dict(dict_csv_path: str) -> List[Dict[str, str]]:
    if not os.path.exists(dict_csv_path):
        return []
    df = pd.read_csv(dict_csv_path, dtype=str, encoding="utf-8").fillna("")
    rows: List[Dict[str, str]] = []
    for _, r in df.iterrows():
        rows.append({"type": r.get("type", ""), "pattern": r.get("pattern", ""), "value": r.get("value", "")})
    return rows


def derive_category(name: str, dict_rows: List[Dict[str, str]]) -> str:
    n = clean_cell(name)
    # 1) explicit category dictionary (substring/regex-like check)
    for row in dict_rows:
        if row["type"] != "category":
            continue
        pat = row["pattern"]
        if not pat:
            continue
        try:
            if re.search(pat, n):
                return row["value"]
        except re.error:
            if pat in n:
                return row["value"]
    # 2) heuristic fallback
    if n.endswith("運転") or any(k in n for k in ["フィニッシャ", "カッタ", "クレーン", "スプレッダ", "レベラ", "ジャンボ", "ショベル", "バックホウ", "ローラ"]):
        return "機械"
    if any(k in n for k in ["コンクリート", "接着剤", "鉄筋", "鉄網", "目地材", "砕石", "砂利", "スペーサー", "アスファルト"]):
        return "資材"
    if any(k in n for k in ["作業員", "世話役", "工"]):
        return "労務"
    return ""


def build_final_df(unit_price_csv: str, table_data_csv: str, dict_csv: str) -> pd.DataFrame:
    # Load normalized unit price data
    up_df = load_and_normalize_unit_price(unit_price_csv)
    # Optional auxiliary data (currently unused but kept for future merger/validation)
    _aux = normalize_table_data_for_aux(table_data_csv)

    # Load category dictionary
    dict_rows = load_category_dict(dict_csv)

    # Map to target schema
    final_rows: List[Dict[str, str]] = []
    for _, r in up_df.iterrows():
        name = r.get("名称", "")
        unit_ = normalize_unit(r.get("単位", ""))
        qty = clean_cell(r.get("数量", ""))
        category = derive_category(name, dict_rows)
        # 所要日数作業単位_* は空欄（要望）、歩掛作業単位_* はUP側の値を使用
        work_qty = clean_cell(r.get("歩掛作業単位_数量", ""))
        work_unit = normalize_unit(r.get("歩掛作業単位_単位", ""))

        final_rows.append(
            {
                "大分類名": r.get("大分類名", ""),
                "工種名": r.get("工種名", ""),
                "細別名": r.get("細別名", ""),
                "所要日数作業単位_数量": "",
                "所要日数作業単位_単位": "",
                "基本所要日数名": "",
                "基本所要日数": "",
                "歩掛作業単位_数量": work_qty,
                "歩掛作業単位_単位": work_unit,
                "基本歩掛名": "",
                "歩掛カテゴリ": category,
                "項目名": name,
                "歩掛数量": qty,
                "歩掛単位": unit_,
            }
        )

    final_df = pd.DataFrame(final_rows)
    # Column order as specified
    col_order = [
        "大分類名",
        "工種名",
        "細別名",
        "所要日数作業単位_数量",
        "所要日数作業単位_単位",
        "基本所要日数名",
        "基本所要日数",
        "歩掛作業単位_数量",
        "歩掛作業単位_単位",
        "基本歩掛名",
        "歩掛カテゴリ",
        "項目名",
        "歩掛数量",
        "歩掛単位",
    ]
    final_df = final_df[col_order]
    return final_df


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.normpath(os.path.join(base_dir, "..", "data"))
    output_dir = os.path.join(data_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    unit_price_csv = os.path.join(data_dir, "unit_price_table_data.csv")
    table_data_csv = os.path.join(data_dir, "table_data.csv")
    dict_csv = os.path.join(base_dir, "keyword_map.csv")  # placed under src

    final_df = build_final_df(unit_price_csv, table_data_csv, dict_csv)
    out_csv = os.path.join(output_dir, "final_mapping.csv")
    final_df.to_csv(out_csv, index=False, encoding="utf-8")
    print(f"Wrote: {out_csv} (rows={len(final_df)})")


if __name__ == "__main__":
    main()


