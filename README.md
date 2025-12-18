## 概要
道路工事のアイテム一覧（カテゴリ/サブカテゴリ/アイテム名）を、標準歩掛の単価表データに正規化・マッピングするためのスクリプト群です。

- 単価表の正規化: `src/preprocess_unit_price.py`
- 正規化データから最終行構造を生成: `src/build_final_from_unit_price.py`
- 道路工事アイテムとのマッチング: `bugakari/scripts/map_road_items_to_unit_prices.py`

## 前提
- Python 3.10+ 推奨
- 依存: `pandas`, `rapidfuzz`

```powershell
python -m pip install -r ".\requirements.txt"
```

## データ配置（主要）
- 原データ
  - `bugakari/data/unit_price_table_data.csv`（単価表・生CSV）
  - `bugakari/data/table_data.csv`（補助用・任意）
  - `bugakari/data/道路工事.xlsx - Sheet1.csv`（道路工事アイテム）
- 正規化出力
  - `bugakari/data/normalized/unit_price_normalized.csv`
  - `bugakari/data/normalized/table_data_aux.csv`
- マッピング出力
  - `bugakari/data/mappings/道路工事_unit_price_candidates.csv`
  - `bugakari/data/mappings/道路工事_unmatched.csv`

## 単価表の正規化
```powershell
python ".\bugakari\src\preprocess_unit_price.py"
```
出力: `bugakari/data/normalized/unit_price_normalized.csv`

正規化の主な仕様:
- カラム: `大分類名, 工種名, 細別名, 基本歩掛名, 所要日数作業単位_数量, 所要日数作業単位_単位, 歩掛作業単位_数量, 歩掛作業単位_単位, 名称, 規格, 単位, 数量, 摘要`
- 区切りの厳密化: CSVのクォートを尊重（`csv.reader`）し、`1,000`などのカンマを誤分割しない
- カテゴリ分割: 左端トークン＋スペースで（先頭トークン=大分類名、残り=工種名）。重複開始（例: `共通工 共通工 ...`）は右側を整理
- 「単価表」の語は全テキスト列から除去（数字は保持）
- `二重管工法 1,000mm以上2,000mm以下` などの細別名と、`1 本`などの作業単位（数量/単位）を抽出

## 最終行構造の生成（任意）
```powershell
python ".\bugakari\src\build_final_from_unit_price.py"
```
出力: `bugakari/data/output/final_mapping.csv`

## 道路工事アイテムとのマッチング
```powershell
python ".\bugakari\scripts\map_road_items_to_unit_prices.py" ^
  --road ".\bugakari\data\道路工事.xlsx - Sheet1.csv" ^
  --unit ".\bugakari\data\normalized\unit_price_normalized.csv" ^
  --outdir ".\bugakari\data\mappings" ^
  --threshold 85 ^
  --cat_filter borkind
```
PowerShell 1行版:
```powershell
python ".\bugakari\scripts\map_road_items_to_unit_prices.py" --road ".\bugakari\data\道路工事.xlsx - Sheet1.csv" --unit ".\bugakari\data\normalized\unit_price_normalized.csv" --outdir ".\bugakari\data\mappings" --threshold 85 --cat_filter borkind
```

出力:
- 候補: `道路工事_unit_price_candidates.csv`
  - 列: 道路側（`カテゴリ名, サブカテゴリ名, アイテム名`）＋単価側（`大分類名, 工種名, 細別名, 名称, 規格, 単位, 数量, 摘要`）＋`match_on, match_score`
- 未ヒット: `道路工事_unmatched.csv`

### フィルタモード（候補の絞り込み）
- `--cat_filter both`: 単価側の`工種名`に「カテゴリ名」と「サブカテゴリ名」の両方を含む行のみ
- `--cat_filter either`: 単価側の`工種名`にどちらか一方を含む行
- `--cat_filter borkind`（推奨）: 単価側の`大分類名`に「カテゴリ名」を含む OR 単価側の`工種名`に「サブカテゴリ名」を含む行

### 類似度判定
- 「アイテム名」 vs 「細別名」「名称」を `rapidfuzz.fuzz.WRatio` で比較し、高い方を採用
- `--threshold`（既定85）以上の候補のみ出力。候補を増やしたい場合は80や75に下げる

### 上位候補の抽出（任意）
```python
import pandas as pd
df = pd.read_csv(r".\bugakari\data\mappings\道路工事_unit_price_candidates.csv", dtype=str)
df["match_score"] = df["match_score"].astype(int)
top1 = (df.sort_values(["カテゴリ名","サブカテゴリ名","アイテム名","match_score"], ascending=[True,True,True,False])
          .drop_duplicates(["カテゴリ名","サブカテゴリ名","アイテム名"]))
top1.to_csv(r".\bugakari\data\mappings\道路工事_unit_price_top1.csv", index=False, encoding="utf-8")
```

## よくある質問
- 候補が少ない:
  - `--threshold` を下げる（85→80/75）
  - `--cat_filter borkind` を使う（ヒット範囲を広げる）
- カテゴリ語尾のゆらぎ（例: `舗装工` vs `舗装`）で漏れる:
  - `borkind` を使うか、語尾（「工/工事/工法」）除去の前処理を追加する

## 補助スクリプトの概要
- `extract_tables_fromW.py`: Word(DOCX)から表抽出→CSV
- `modify_csv.py`: スペース/改行の整理、簡易パターンの整形
- `merge_tables.py`: 取得済み表の自動マージ（所定パターンをGroup化して結合）

## ライセンス
本リポジトリ内のスクリプトは、特記なき場合MIT相当を想定（要調整）。業務データ（PDF/CSV等）の取扱いは社内規約に従ってください。
