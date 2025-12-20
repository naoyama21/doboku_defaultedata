## README

### 概要
- 歩掛・単価表の生データを整形→正規化→道路工事アイテムとファジー照合→最終CSV生成までを実行。
- 流れ: PDF分割 → Geminiで表抽出 → 抽出CSV結合 → 分類 → クリーニング → 正規化（手直し可） → 照合 → 最終集計。

### 前提・セットアップ
- Python 3.10+ 推奨
- 依存関係:
```powershell
$ROOT = "c:\Users\Shimoyama Naoya\OneDrive\ドキュメント\kencopa\default data\bugakari"
python -m pip install -r "$ROOT\requirements.txt"
python -m pip install PyPDF2
```

### クイックスタート
1) PDF分割（50～100ページ単位推奨）
```powershell
python "$ROOT\src\split_pdf.py"
```

2) Geminiで表抽出（CSV）
- 出力要件（重要）:
  - UTF-8 / カンマ区切り、各セルはダブルクォートで囲む
  - 各レコードは2列以上、2列目に必ず表の見出し/タイトル（「単価表」を含むなら必ず残す）
  - カンマ/改行はセル内クォートで保持
- 例: `data\tmp\gemini_tables_chunk_0001.csv` などに保存

3) 抽出CSVの結合 → 分類入力ファイル作成（中身はCSV形式）
```powershell
$TMP = "$ROOT\data\tmp"
$OUT = "$ROOT\data\第２編土木工事標準歩掛.txt"
Get-ChildItem $TMP\gemini_tables_chunk_*.csv | Get-Content | Set-Content -Encoding UTF8 $OUT
```

4) 表/単価表の分類（2列目に「単価表」を含むかで判定）
```powershell
python "$ROOT\src\classify_data_from_file.py"
```
- 出力: `data/table_data_raw.csv`, `data/unit_price_table_data_raw.csv`

5) クリーニング（番号/丸数字/枝番/単価表(1) 等の除去・7列揃え）
```powershell
python "$ROOT\src\prepare_unit_price_from_raw.py"
Copy-Item "$ROOT\data\unit_price_table_data_raw_cleaned.csv" "$ROOT\data\unit_price_table_data.csv" -Force
```

6) 正規化（照合用データ生成）
```powershell
python "$ROOT\src\preprocess_unit_price.py"
```
- 出力: `data/normalized/unit_price_normalized.csv`
- ここで一度、人手でおかしな箇所があれば修正（例: 大分類/工種の分割、細別名、単価表の取り残し、ヘッダ/計/機械運転の混入）

7) 照合（候補/未一致の作成）
```powershell
# either: 工種名が（カテゴリ or サブカテゴリ）のいずれかを含む
python "$ROOT\src\map_road_items_to_unit_prices.py" --cat_filter either --threshold 85

# borkind: 大分類名がカテゴリ or 工種名がサブカテゴリを含む
python "$ROOT\src\map_road_items_to_unit_prices.py" --cat_filter borkind --threshold 80
```
- 出力: `data/mappings/道路工事_unit_price_candidates.csv`, `data/mappings/道路工事_unmatched.csv`

8) 最終集計（任意）
```powershell
python "$ROOT\src\build_final_from_unit_price.py"
```
- 出力: `data/output/final_mapping.csv`

### 最終CSVの列（最新仕様）
- カテゴリ名, サブカテゴリ名, アイテム名
- 所要日数作業単位_数量, 所要日数作業単位_単位
- 基本所要日数名, 基本所要日数
- 歩掛作業単位_数量, 歩掛作業単位_単位
- 基本歩掛名, 歩掛カテゴリ
- 項目名, 歩掛数量, 歩掛単位, 説明（= 摘要）

### トラブルシューティング
- パスは必ず二重引用符で囲む（空白/日本語対策）
- 7列に揃っていない → クリーニング後に必ず `unit_price_table_data.csv` へコピー
- 候補が少ない/多い → `--threshold` 調整、`--cat_filter` を `either`/`borkind` に変更
- 「単価表」が残る → Gemini出力の「2列目＝見出し」要件と正規化時の置換を再確認

### 9) 未照合の手当て（手作業）
- 対象: `data/mappings/道路工事_unmatched.csv`
- 方針:
  - 単価側に該当がある場合: 正規化/抽出/フィルタ、`keyword_map.csv` を調整して 5→6→7 を再実行
  - 単価側に該当が無い場合: 対応する歩掛を手作成し、最終CSVに反映
    - 簡易対応: `data/output/final_mapping.csv` に同じ列構成で追記
    - 推奨: `data/manual_overrides.csv`（最終列と同じカラム）を用意し、後で `build_final_from_unit_price.py` に取り込み処理を追加
