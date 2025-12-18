import csv
import os
import re
from typing import List

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.normpath(os.path.join(BASE_DIR, "..", "data"))

RAW_IN = os.path.join(DATA_DIR, "unit_price_table_data_raw.csv")
OUT_CSV = os.path.join(DATA_DIR, "unit_price_table_data_raw_cleaned.csv")

# ①〜⑳
CIRCLED = r"[\u2460-\u2473]"
# 各種ダッシュ/長音
DASH = r"[‐‑–—ー-]"


def clean_preserve_spaces(text: str) -> str:
    """
    文字列の軽量クリーニング（内部の空白は潰さない）
    - 改行/復帰を空白に置換
    - 前後の空白と外側のクォートを除去
    """
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    text = text.replace("\r", " ").replace("\n", " ")
    text = text.strip().strip('\'"')
    return text


def strip_leading_numbering(s: str) -> str:
    """
    先頭にある見出し番号だけを除去する
    - 例: '2章 ' / '(2)' / '（2）' / '①' / '①-2' / '3-6' / '2.' / '2)'
    """
    if s is None:
        return ""
    t, prev = s, None
    while t != prev:
        prev = t
        t = re.sub(r"^\s*\d+\s*章\s*", " ", t)
        t = re.sub(r"^\s*[（(]\s*\d+\s*[)）]\s*", " ", t)
        t = re.sub(rf"^\s*{CIRCLED}(?:\s*{DASH}\s*\d+)?\s*", " ", t)
        t = re.sub(rf"^\s*\d+\s*{DASH}\s*\d+\s*", " ", t)
        t = re.sub(r"^\s*\d+\s*[.)．)]\s*", " ", t)
        # 余分な空白整形
        t = re.sub(r"\s+", " ", t).strip()
    return t


def strip_heading_tokens_anywhere(s: str) -> str:
    """
    行中の見出し番号トークンを除去する（語と語の区切りに現れるもの）
    - 例: (2) / （2） / ① / ①-2 / 3-6 / 2. / 2)
    - '100m³' のような内容の数値は対象外
    - トークン境界は空白/行端で判定（'排水材設置工-2' のような語中は基本対象外）
    """
    if s is None:
        return ""
    t, prev = s, None
    # 厳密な空白境界（直前が行頭/空白、直後が空白/行末の場合に一致）
    bp = r"(?<!\S)"
    bs = r"(?!\S)"
    pattern = rf"{bp}(?:[（(]\s*\d+\s*[)）]|{CIRCLED}(?:\s*{DASH}\s*\d+)?|\d+\s*{DASH}\s*\d+|{DASH}\s*\d+|\d+\s*[.)．)]){bs}"
    while t != prev:
        prev = t
        t = re.sub(pattern, " ", t)
        t = re.sub(rf"{bp}\d+\s*章{bs}", " ", t)
        # 語に隣接した丸数字/丸数字+枝番、括弧付き番号も空白に置換
        t = re.sub(rf"{CIRCLED}\s*{DASH}\s*\d+", " ", t)  # ⑤-2 など（語に隣接）
        t = re.sub(rf"{CIRCLED}", " ", t)                 # ⑤ 単体（語に隣接）
        t = re.sub(r"[（(]\s*\d+\s*[)）]", " ", t)        # (2) / （2） （語に隣接）
        t = t.strip()
    return t


def normalize_row(parts: List[str]) -> List[str]:
    """
    1行を7列に正規化する（不足はパディング、超過は末尾に結合）
    0列目（raw_category）/1列目（raw_table）には見出し番号の除去等を適用
    """
    def dedupe_trailing_dash_number_when_repeated(text: str) -> str:
        """
        「-数字」の削除ルール（語の重複を伴う場合のみ削除）
        - 直前または直後のトークンが同一語（括弧内注記は無視）なら、後者の語末「-数字」を削除
        - 例: '道路維持修繕 道路維持修繕-1 道路除雪工'
              → '道路維持修繕 道路維持修繕 道路除雪工'
        - 単独の '排水材設置工-2' は維持（重複がないため）
        """
        if not text:
            return text
        toks = text.split()
        out: List[str] = []

        def norm_base(s: str) -> str:
            # 末尾の -数字 を除去し、後続の括弧以降を比較用に除去
            s1 = re.sub(rf"{DASH}\s*\d+$", "", s)
            s1 = re.sub(r"\s*[（(].*$", "", s1)
            return s1

        for i, tok in enumerate(toks):
            base = re.sub(rf"{DASH}\s*\d+$", "", tok)
            if re.search(rf"{DASH}\s*\d+$", tok):
                prev_same = False
                next_same = False
                if i > 0:
                    prev_same = norm_base(base) == norm_base(toks[i - 1])
                if i + 1 < len(toks):
                    next_same = norm_base(base) == norm_base(toks[i + 1])
                if prev_same or next_same:
                    tok = base
            out.append(tok)
        return " ".join(out)

    if not parts or not any((p or "").strip() for p in parts):
        return []
    # Fix to 7 fields: pad or join tail into last field
    if len(parts) < 7:
        parts = parts + [""] * (7 - len(parts))
    elif len(parts) > 7:
        head, tail = parts[:6], parts[6:]
        parts = head + [",".join(tail)]

    # Column 0/1: heading/number tokens removal (leading + anywhere)
    c0 = strip_heading_tokens_anywhere(strip_leading_numbering(clean_preserve_spaces(parts[0])))
    c0 = dedupe_trailing_dash_number_when_repeated(c0)
    # 文中の「工-数字」を削除（例: 仮囲い設置・撤去工-2 → 仮囲い設置・撤去工）
    c0 = re.sub(rf"(?<=工)\s*{DASH}\s*\d+\b", " ", c0)
    c0 = re.sub(r"\s+", " ", c0).strip()
    c1 = strip_heading_tokens_anywhere(strip_leading_numbering(clean_preserve_spaces(parts[1])))
    # raw_table 特有: 「単価表(1)」などの括弧付き番号を除去し「単価表」に統一
    c1 = re.sub(r"(単価表)\s*[（(]\s*\d+\s*[)）]", r"\1", c1)
    # raw_table 特有: 先頭に「単価表 」があり、その後方にも「単価表」がある場合は先頭側を削除
    # 例: 「単価表 防水工100m²当り単価表」→「防水工100m²当り単価表」
    c1 = re.sub(r"^\s*単価表\s+(?=.*単価表)", " ", c1)
    c1 = re.sub(r"\s+", " ", c1).strip()
    c1 = dedupe_trailing_dash_number_when_repeated(c1)
    # raw_table 内の「工-数字」も念のため除去
    c1 = re.sub(rf"(?<=工)\s*{DASH}\s*\d+\b", " ", c1)
    c1 = re.sub(r"\s+", " ", c1).strip()

    # 他列は軽量クリーニングのみ（内部空白は保持）
    c2 = clean_preserve_spaces(parts[2])
    c3 = clean_preserve_spaces(parts[3])
    c4 = clean_preserve_spaces(parts[4])
    c5 = clean_preserve_spaces(parts[5])
    c6 = clean_preserve_spaces(parts[6])
    return [c0, c1, c2, c3, c4, c5, c6]


def main() -> None:
    rows: List[List[str]] = []
    with open(RAW_IN, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=",", quotechar='"')
        for parts in reader:
            row = normalize_row(parts)
            if row:
                rows.append(row)
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(OUT_CSV, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL)
        w.writerows(rows)
    print(f"Wrote: {OUT_CSV} (rows={len(rows)})")


if __name__ == "__main__":
    main()

