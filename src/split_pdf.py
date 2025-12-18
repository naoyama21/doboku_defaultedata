import os
from PyPDF2 import PdfReader, PdfWriter

def extract_pages_to_single_pdf(input_pdf_path, output_pdf_path, start_page=None, end_page=None):
    """
    PDFファイルから指定されたページの範囲を抽出し、新しい単一のPDFファイルとして保存します。

    Args:
        input_pdf_path (str): 抽出元のPDFファイルのパス。
        output_pdf_path (str): 抽出したページを保存する新しいPDFファイルのパス。
        start_page (int, optional): 抽出を開始するページ番号 (1から)。指定しない場合は最初のページから。
        end_page (int, optional): 抽出を終了するページ番号。指定しない場合は最後のページまで。
    """
    # 出力ディレクトリが存在しない場合は作成
    output_directory = os.path.dirname(output_pdf_path)
    if output_directory and not os.path.exists(output_directory):
        os.makedirs(output_directory)
        print(f"ディレクトリ '{output_directory}' を作成しました。")

    try:
        reader = PdfReader(input_pdf_path)
        num_pages = len(reader.pages)

        _start_page = start_page if start_page is not None else 1
        _end_page = end_page if end_page is not None else num_pages

        if not (1 <= _start_page and _start_page <= _end_page and _end_page <= num_pages):
            print(f"エラー: 指定されたページ範囲 ({_start_page}-{_end_page}) は無効です。総ページ数: {num_pages}")
            return

        writer = PdfWriter()

        for i in range(_start_page - 1, _end_page):
            writer.add_page(reader.pages[i])

        with open(output_pdf_path, "wb") as output_file:
            writer.write(output_file)
        
        print(f"ページ {_start_page} から {_end_page} までを '{output_pdf_path}' に保存しました。")

    except PermissionError:
        print(f"\nエラー: ファイル '{os.path.abspath(output_pdf_path)}' への書き込みが拒否されました (Permission denied)。")
        print("考えられる原因:")
        print("  - ファイルが他のプログラム (PDFビューア、OneDriveの同期プロセス等) で開かれている、またはロックされている。")
        print("  - 出力先フォルダに対する書き込み権限がない。")
        print("  - ファイルが読み取り専用に設定されている。")
        print("ヒント: 問題のファイルを閉じる、OneDriveの同期が完了するのを待つ、または別のフォルダに出力してみてください。")
        raise
    except FileNotFoundError:
        raise
    except Exception as e:
        print(f"PDFの処理中に予期せぬエラーが発生しました: {e}")
        raise

def split_pdf_in_chunks(input_pdf_path, output_directory, chunk_size=100):
    """
    PDFファイルを指定されたページ数ごとに分割します。

    Args:
        input_pdf_path (str): 分割元のPDFファイルのパス。
        output_directory (str): 分割したPDFを保存するディレクトリ。
        chunk_size (int, optional): 1ファイルあたりのページ数。デフォルトは100。
    """
    try:
        reader = PdfReader(input_pdf_path)
        num_pages = len(reader.pages)
        base_name = os.path.splitext(os.path.basename(input_pdf_path))[0]

        print(f"'{input_pdf_path}' (総ページ数: {num_pages}) を {chunk_size} ページごとに分割します。")

        for start_page in range(1, num_pages + 1, chunk_size):
            end_page = min(start_page + chunk_size - 1, num_pages)
            
            output_filename = f"{base_name}_pages_{start_page}.pdf"
            output_filepath = os.path.join(output_directory, output_filename)
            
            extract_pages_to_single_pdf(input_pdf_path, output_filepath, start_page, end_page)

        print("\nPDFの分割が完了しました。")

    except FileNotFoundError:
        print(f"エラー: 入力ファイル '{input_pdf_path}' が見つかりません。")
    except Exception:
        print(f"処理を中止しました。")

if __name__ == "__main__":
    # スクリプトの場所を基準にファイルパスを解決
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_file = os.path.join(script_dir, "..", "data", "第２編土木工事標準歩掛_OCR結合済み.pdf")

    # 出力先フォルダを指定 (スクリプトの2階層上の 'pdf' フォルダ)
    output_folder = os.path.join(script_dir, "..", "data", "50_pdf_doboku")
    
    # PDFを100ページごとに分割
    split_pdf_in_chunks(input_file, output_folder, chunk_size=50)