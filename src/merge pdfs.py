from PyPDF2 import PdfMerger
import os

def merge_pdfs(input_pdf_paths, output_pdf_path):
    """
    複数のPDFファイルを結合して、一つの新しいPDFファイルを作成します。

    :param input_pdf_paths: 結合するPDFファイルのパスのリスト (list of str)
    :param output_pdf_path: 出力する結合済みPDFファイルのパス (str)
    """
    
    # PdfMergerオブジェクトを作成
    merger = PdfMerger()
    
    print("--- 結合処理を開始します ---")
    
    all_files_exist = True
    
    # 入力ファイルをmergerに追加
    for pdf_path in input_pdf_paths:
        if os.path.exists(pdf_path):
            print(f"✅ ファイルを追加: {pdf_path}")
            # PDFファイルをmergerに追加
            merger.append(pdf_path)
        else:
            print(f"❌ エラー: ファイル '{pdf_path}' が見つかりませんでした。スキップします。")
            all_files_exist = False

    if not all_files_exist and not input_pdf_paths:
         print("エラー: 結合するための有効な入力ファイルが指定されていません。")
         merger.close()
         return
    
    try:
        # 結合した結果を新しいファイルとして書き出し
        with open(output_pdf_path, "wb") as output_file:
            merger.write(output_file)
            
        print(f"🎉 結合が完了しました。出力ファイル: '{output_pdf_path}'")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        
    finally:
        # リソースを解放
        merger.close()


# --- 使用例 ---
# ⚠️ 注意: 以下のファイルパスを、ご自身の環境に合わせて変更してください。
# 結合したい順にファイルパスをリストで指定します。

# 1. 結合したいPDFファイルのリスト
pdf_files_to_merge = [
    "../data/pdf_3_page_chunks/第２編土木工事標準歩掛_pages_1-314.pdf",
    "../data/pdf_3_page_chunks/第２編土木工事標準歩掛_pages_315-628.pdf",
    "../data/pdf_3_page_chunks/第２編土木工事標準歩掛_pages_629-941.pdf" 
]

# 2. 出力される結合済みPDFファイルのパス
output_file_name = "第２編土木工事標準歩掛_OCR結合済み.pdf"
output_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", output_file_name) 

# 3. 関数を実行
if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # チェックのために、入力ファイルのフルパスを作成
    files_to_check = [os.path.join(script_dir, path) for path in pdf_files_to_merge]
    
    missing_files = [path for path in files_to_check if not os.path.exists(path)]

    if missing_files:
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("‼️ エラー: 処理に必要なファイルがありません。")
        print("まず、下のコマンドを実行して、PDFファイルを分割してください。")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        
        prepare_script_path = os.path.join(script_dir, "prepare_chunks.py")
        print("\n👉 以下のコマンドをコピーして実行してください:\n")
        print(f'python -u "{prepare_script_path}"\n')
    else:
        # merge_pdfsには解決済みの絶対パスリストと、出力ファイルパスを渡す
        # これでCWDに依存せずに動作する
        merge_pdfs(files_to_check, output_file_path)