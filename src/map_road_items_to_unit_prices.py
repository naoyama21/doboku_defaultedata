import argparse
import re
import unicodedata
from pathlib import Path
from typing import List, Tuple, Set

import pandas as pd
from rapidfuzz import fuzz


# 文字列を比較用に正規化（NFKC化 → 前後空白除去 → 連続空白の圧縮）
def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.strip()
    # 空白類を単一スペースに統一
    text = re.sub(r"\s+", " ", text)
    return text


def best_score(item_norm: str, row_norm_shobetsu: str, row_norm_meishou: str) -> Tuple[int, str]:
    """
    アイテム名（正規化済み）と「細別名」「名称」（正規化済み）の双方を
    WRatio で照合し、より高いスコアと照合対象名を返す。
    """
    s1 = fuzz.WRatio(item_norm, row_norm_shobetsu or "")
    s2 = fuzz.WRatio(item_norm, row_norm_meishou or "")
    if s1 >= s2:
        return int(s1), "細別名"
    return int(s2), "名称"


def forward_fill_categories(df: pd.DataFrame) -> pd.DataFrame:
    # 「カテゴリ名」「サブカテゴリ名」の見出しセルを前方埋め（表形式の段組想定）
    cols = ["カテゴリ名", "サブカテゴリ名"]
    present = [c for c in cols if c in df.columns]
    if present:
        df[present] = df[present].ffill()
    return df


def build_defaults_from_script() -> Tuple[Path, Path, Path]:
    # スクリプトの位置からデフォルト入出力パスを解決
    script_path = Path(__file__).resolve()
    # ルートは2階層上を想定（.../bugakari/src/... → ルート → bugakari）
    repo_root = script_path.parents[2]
    road_csv = repo_root / "bugakari" / "data" / "道路工事.xlsx - Sheet1.csv"
    unit_csv = repo_root / "bugakari" / "data" / "normalized" / "unit_price_normalized.csv"
    out_dir = repo_root / "bugakari" / "data" / "mappings"
    return road_csv, unit_csv, out_dir


def main():
    default_road, default_unit, default_outdir = build_defaults_from_script()

    # コマンドライン引数（説明は日本語で記載）
    parser = argparse.ArgumentParser(description="道路工事アイテムを単価データへファジー照合し、候補一覧を作成します。")
    parser.add_argument("--road", type=Path, default=default_road, help="道路工事.xlsx - Sheet1.csv のパス")
    parser.add_argument("--unit", type=Path, default=default_unit, help="unit_price_normalized.csv のパス")
    parser.add_argument("--outdir", type=Path, default=default_outdir, help="出力ディレクトリ（候補/未一致CSV）")
    parser.add_argument("--threshold", type=int, default=85, help="ファジー一致のしきい値（0-100）")
    parser.add_argument(
        "--cat_filter",
        type=str,
        default="both",
        choices=["both", "either", "borkind"],
        help="Filter rows by category rule: both=工種名にカテゴリ/サブカテゴリの両方を含む, either=どちらか一方を含む, borkind=大分類名にカテゴリ or 工種名にサブカテゴリを含む",
    )

    args = parser.parse_args()

    outdir: Path = args.outdir
    outdir.mkdir(parents=True, exist_ok=True)

    # データ読み込み
    road_df = pd.read_csv(args.road, encoding="utf-8")
    unit_df = pd.read_csv(args.unit, encoding="utf-8")

    # 道路側：カテゴリ見出しの前方埋め → アイテム名の空行を除外
    road_df = forward_fill_categories(road_df)
    if "アイテム名" in road_df.columns:
        road_df = road_df[~road_df["アイテム名"].isna() & (road_df["アイテム名"].astype(str).str.strip() != "")]
    else:
        raise ValueError("入力CSVに 'アイテム名' 列が見つかりません: " + str(args.road))

    # 比較用の正規化カラムを追加
    road_df["norm_カテゴリ名"] = road_df["カテゴリ名"].map(normalize_text)
    road_df["norm_サブカテゴリ名"] = road_df["サブカテゴリ名"].map(normalize_text)
    road_df["norm_アイテム名"] = road_df["アイテム名"].map(normalize_text)

    # 単価側に必要なカラムが存在するか検証
    for col in ["大分類名", "工種名", "細別名", "名称"]:
        if col not in unit_df.columns:
            raise ValueError(f"unit_price_normalized.csv に必要な列が見つかりません: {col}")

    # 単価側の比較用正規化カラム
    unit_df["norm_大分類名"] = unit_df["大分類名"].map(normalize_text)
    unit_df["norm_工種名"] = unit_df["工種名"].map(normalize_text)
    unit_df["norm_細別名"] = unit_df["細別名"].map(normalize_text)
    unit_df["norm_名称"] = unit_df["名称"].map(normalize_text)

    candidates_records: List[dict] = []

    # 道路側アイテムごとに候補検索
    for _, item_row in road_df.iterrows():
        cat = item_row.get("norm_カテゴリ名", "")
        sub = item_row.get("norm_サブカテゴリ名", "")
        item = item_row.get("norm_アイテム名", "")

        if not item:
            continue

        # カテゴリフィルタ条件に応じて単価側を絞り込む
        if args.cat_filter == "both":
            # 工種名が「カテゴリ名」かつ「サブカテゴリ名」を両方含む
            if not cat or not sub:
                filtered = unit_df.iloc[0:0]
            else:
                mask = unit_df["norm_工種名"].str.contains(cat, na=False) & unit_df["norm_工種名"].str.contains(sub, na=False)
                filtered = unit_df[mask]
        elif args.cat_filter == "either":
            # 工種名が（カテゴリ or サブカテゴリ）のいずれかを含む
            terms = [t for t in [cat, sub] if t]
            if not terms:
                filtered = unit_df.iloc[0:0]
            else:
                mask = False
                for t in terms:
                    mask = mask | unit_df["norm_工種名"].str.contains(t, na=False)
                filtered = unit_df[mask]
        else:  # borkind: 大分類名がカテゴリを含む OR 工種名がサブカテゴリを含む
            mask = False
            if cat:
                mask = mask | unit_df["norm_大分類名"].str.contains(cat, na=False)
            if sub:
                mask = mask | unit_df["norm_工種名"].str.contains(sub, na=False)
            filtered = unit_df[mask]

        # ファジー一致スコアを計算し、しきい値以上のみを候補として採用
        for _, urow in filtered.iterrows():
            score, match_on = best_score(item, urow.get("norm_細別名", ""), urow.get("norm_名称", ""))
            if score >= args.threshold:
                rec = {
                    "カテゴリ名": item_row.get("カテゴリ名", ""),
                    "サブカテゴリ名": item_row.get("サブカテゴリ名", ""),
                    "アイテム名": item_row.get("アイテム名", ""),
                    "大分類名": urow.get("大分類名", ""),
                    "工種名": urow.get("工種名", ""),
                    "細別名": urow.get("細別名", ""),
                    "名称": urow.get("名称", ""),
                    "規格": urow.get("規格", ""),
                    "単位": urow.get("単位", ""),
                    "数量": urow.get("数量", ""),
                    "摘要": urow.get("摘要", ""),
                    "match_on": match_on,
                    "match_score": score,
                }
                candidates_records.append(rec)

    # 未一致リスト：候補に一度も現れなかった（カテゴリ, サブカテゴリ, アイテム名）の組を集合差で抽出
    base_unmatched = road_df[["カテゴリ名", "サブカテゴリ名", "アイテム名"]].drop_duplicates()
    if candidates_records:
        matched_keys: Set[Tuple[str, str, str]] = set((r["カテゴリ名"], r["サブカテゴリ名"], r["アイテム名"]) for r in candidates_records)
        mask_unmatched = ~base_unmatched.apply(lambda r: (r["カテゴリ名"], r["サブカテゴリ名"], r["アイテム名"]) in matched_keys, axis=1)
        unmatched_df = base_unmatched[mask_unmatched]
    else:
        unmatched_df = base_unmatched

    # 出力先のパス
    out_candidates = args.outdir / "道路工事_unit_price_candidates.csv"
    out_unmatched = args.outdir / "道路工事_unmatched.csv"

    # 候補CSVの書き出し（候補が無ければヘッダのみ）
    if candidates_records:
        pd.DataFrame.from_records(candidates_records).to_csv(out_candidates, index=False, encoding="utf-8")
    else:
        pd.DataFrame(columns=[
            "カテゴリ名", "サブカテゴリ名", "アイテム名",
            "大分類名", "工種名", "細別名",
            "名称", "規格", "単位", "数量", "摘要",
            "match_on", "match_score",
        ]).to_csv(out_candidates, index=False, encoding="utf-8")

    # 未一致CSVの書き出し
    unmatched_df.to_csv(out_unmatched, index=False, encoding="utf-8")

    print(f"Wrote candidates: {out_candidates}")
    print(f"Wrote unmatched:  {out_unmatched}")


if __name__ == "__main__":
    main()
