import sys
import os

try:
    from pdf2docx import Converter
    import fitz  # PyMuPDF, a dependency of pdf2docx
except ImportError:
    print("エラー: pdf2docx またはその依存ライブラリ(PyMuPDF)が見つかりません。")
    print("次のコマンドでインストールしてください: pip install pdf2docx")
    sys.exit()

def convert_pdf_pages_to_docx(pdf_path, output_dir, start_page=None, end_page=None):
    """
    指定されたPDFの各ページを、指定範囲で個別のDOCXファイルに変換します。

    Args:
        pdf_path (str): 入力PDFファイルのパス。
        output_dir (str): 出力DOCXファイルを保存するディレクトリ。
        start_page (int, optional): 変換を開始するページ番号 (1から)。デフォルトは1。
        end_page (int, optional): 変換を終了するページ番号。デフォルトは最後のページ。
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"ディレクトリ '{output_dir}' を作成しました。")

    try:
        # PyMuPDF(fitz)を使ってページ数を取得
        with fitz.open(pdf_path) as doc:
            num_pages = doc.page_count
    except Exception as e:
        print(f"❌ PDFファイル '{pdf_path}' を開いてページ数を取得する際にエラーが発生しました: {e}")
        return

    print(f"'{pdf_path}' の総ページ数: {num_pages}")

    # ページの範囲を決定 (1-based)
    _start = start_page if start_page is not None else 1
    _end = end_page if end_page is not None else num_pages

    # ページ範囲の妥当性をチェック
    if not (1 <= _start <= _end <= num_pages):
        print(f"❌ エラー: 指定されたページ範囲 ({_start}-{_end}) は無効です。総ページ数: {num_pages}")
        return

    print(f"⚙️ ページ {_start} から {_end} までの変換を開始します...")

    try:
        cv = Converter(pdf_path)
    except Exception as e:
        print(f"❌ pdf2docxコンバーターの初期化中にエラーが発生しました: {e}")
        return
        
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    for i in range(_start, _end + 1):
        page_index = i - 1  # 0-based index for the library
        docx_path = os.path.join(output_dir, f"{base_name}_page_{i}.docx")
        
        print(f"   - ページ {i} を '{docx_path}' に変換中...")
        try:
            # 1ページだけを変換 (startはinclusive, endはexclusive)
            cv.convert(docx_path, start=page_index, end=i)
            print(f"     ✅ 完了")
        except Exception as e:
            print(f"     ❌ エラー: ページ {i} の変換中に問題が発生しました: {e}")

    cv.close()
    print("\n✅ 変換処理がすべて完了しました。")

if __name__ == "__main__":
    # スクリプト自身の場所を基準にファイルの絶対パスを構築
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 📌 変換したいPDFファイルのパス
    pdf_file = os.path.join(script_dir, "..", "data", "令和7年度版 国土交通省土木工事積算基準 000p.pdf")
    # 📌 出力先ディレクトリ
    output_folder = os.path.join(script_dir, "..", "data", "word_doboku")

    # 📄 変換したいページ範囲を設定 (例: 1ページ目から5ページ目まで)
    start_page_num = 376
    end_page_num = 999

    # ページを指定して変換を実行
    convert_pdf_pages_to_docx(pdf_file, output_folder, start_page=start_page_num, end_page=end_page_num)

    # 全ページを変換する場合は、start_pageとend_pageを省略
    # convert_pdf_pages_to_docx(pdf_file, output_folder)
