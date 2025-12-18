from pdf2docx import Converter
import os
import glob
import time

def convert_pdf_to_docx(pdf_file_path, docx_file_path):
    """
    単一のPDFファイルをWordファイル(.docx)に変換します。

    :param pdf_file_path: 入力PDFファイルのパス (str)
    :param docx_file_path: 出力Wordファイルのパス (str)
    """
    
    # 入力ファイルが存在するか確認
    if not os.path.exists(pdf_file_path):
        print(f"❌ エラー: 入力ファイル '{pdf_file_path}' が見つかりません。")
        return

    # 変換対象のファイル名を取得
    file_name = os.path.basename(pdf_file_path)
    
    print(f"\n--- 変換開始: {file_name} ---")
    
    start_time = time.time()
    
    try:
        # Converterオブジェクトを作成
        cv = Converter(pdf_file_path)
        
        # 変換を実行
        # start=0, end=None は全ページを意味します
        cv.convert(docx_file_path, start=0, end=None) 
        
        # リソースを解放
        cv.close()
        
        end_time = time.time()
        conversion_time = end_time - start_time
        
        print(f"🎉 成功: '{file_name}' -> '{os.path.basename(docx_file_path)}'")
        print(f"   所要時間: {conversion_time:.2f}秒")

    except Exception as e:
        print(f"⚠️ 変換失敗: {file_name}")
        print(f"   エラー内容: {e}")


def process_directory_conversion():
    """
    指定された入力フォルダ内の全PDFファイルを、対応するWordファイルに変換します。
    """
    
    # スクリプトの場所を基準にパスを解決
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # --- 1. パスの設定 ---
    # 既存のパス設定を元に、入力フォルダと出力フォルダを定義します。
    
    # 変換元PDFファイルが格納されているフォルダ
    # 例: .../data/pdf_doboku_split/
    input_base_dir = os.path.join(script_dir, "..", "data")
    input_sub_folder = "pdf_doboku_split"
    input_pdf_folder = os.path.join(input_base_dir, input_sub_folder)
    
    # 変換後のWordファイルを格納するフォルダ
    # 例: .../data/output_docx/
    output_base_dir = os.path.join(script_dir, "..", "data")
    output_docx_folder = os.path.join(output_base_dir, "output_docx")

    print(f"🔍 入力フォルダ: {input_pdf_folder}")
    print(f"📂 出力フォルダ: {output_docx_folder}")
    
    
    # --- 2. 出力フォルダの作成 ---
    if not os.path.exists(output_docx_folder):
        os.makedirs(output_docx_folder)
        print(f"✨ 出力フォルダを作成しました: {output_docx_folder}")

    # --- 3. 処理対象ファイルのリストアップ ---
    # input_pdf_folder内のすべての.pdfファイルを検索します
    # globモジュールを使って、フォルダ内の全ての.pdfファイルをリストアップ
    pdf_files = glob.glob(os.path.join(input_pdf_folder, "*.pdf"))
    
    if not pdf_files:
        print("💡 PDFファイルが見つかりませんでした。パスを確認してください。")
        return

    print(f"\n合計 {len(pdf_files)} 個のPDFファイルを処理します。")

    # --- 4. ファイルを一つずつ変換 ---
    total_start_time = time.time()
    processed_count = 0
    
    for input_pdf_path in pdf_files:
        # PDFファイル名から拡張子(.pdf)を除いた部分を取得
        base_name = os.path.splitext(os.path.basename(input_pdf_path))[0]
        
        # 出力Wordファイル(.docx)のパスを生成
        output_docx_path = os.path.join(output_docx_folder, f"{base_name}.docx")
        
        # 変換関数を呼び出し
        convert_pdf_to_docx(input_pdf_path, output_docx_path)
        processed_count += 1
        
    total_end_time = time.time()
    total_elapsed_time = total_end_time - total_start_time
    
    print("\n==================================")
    print("✅ 全てのPDFファイルの変換処理が完了しました。")
    print(f"処理ファイル数: {processed_count} 個")
    print(f"全体の合計所要時間: {total_elapsed_time:.2f}秒")
    print("==================================")


if __name__ == "__main__":
    process_directory_conversion()