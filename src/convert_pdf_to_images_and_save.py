from pdf2image import convert_from_path
import os

def convert_pdf_to_images_and_save(pdf_path, output_dir, start_page=1, end_page=None, poppler_path=None):
    """
    PDFを指定されたページ範囲で画像ファイルとして保存する関数。

    Args:
        pdf_path (str): 入力PDFファイルのパス。
        output_dir (str): 画像ファイルを保存するディレクトリのパス。
        start_page (int): 変換を開始するページ番号 (1から始まる)。
        end_page (int, optional): 変換を終了するページ番号。Noneの場合、最後まで変換する。
        poppler_path (str, optional): Popplerのbinディレクトリのパス（Windowsなどで必要）。
    """
    
    # 1. 保存先ディレクトリの作成
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"作成されたディレクトリ: {output_dir}")

    # 2. PDFを画像オブジェクトのリストに変換
    try:
        print(f"ページ {start_page} から {end_page if end_page is not None else '最後まで'} の範囲で画像に変換します...")
        
        # start_page と end_page を指定して変換
        images = convert_from_path(
            pdf_path, 
            poppler_path=poppler_path,
            first_page=start_page,
            last_page=end_page
        )
    except Exception as e:
        print(f"PDF画像化エラー: {e}")
        print("Popplerが正しくインストールされているか、Windowsの場合はpoppler_pathが設定されているか確認してください。")
        return

    # 3. 各画像をファイルとして保存
    if images:
        print(f"合計 {len(images)} ページを画像として保存します...")
        
        pdf_base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        
        # 変換された画像リストのインデックス(0から始まる)を、元のページ番号(start_pageから始まる)に変換
        for i, image in enumerate(images):
            current_page_num = start_page + i
            # ファイル名を "{PDFファイル名}_page_{ページ番号}.png" の形式で作成
            image_filename = f"{pdf_base_name}_page_{current_page_num}.png"
            output_path = os.path.join(output_dir, image_filename)
            
            # 画像をPNG形式で保存
            image.save(output_path, 'PNG')
            print(f" -> 保存完了: {image_filename}")
        
        print("--- 指定範囲の画像保存が完了しました。---")
    else:
        print("画像に変換されたページはありませんでした。")

# --- 使用例 ---
# 以下のパスとファイル名を環境に合わせて変更してください

if __name__ == "__main__":
    # --- 使用例 ---
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 以下のパスとファイル名を環境に合わせて変更してください
    INPUT_PDF = os.path.join(script_dir, "..", "data", "令和7年度版 国土交通省土木工事積算基準-1000p.pdf")  # 処理したいPDFファイル名
    OUTPUT_FOLDER = os.path.join(script_dir, "..", "data", "pdf_images")        # 画像を保存するディレクトリ名
    
    # WindowsでPopplerをPATHに設定していない場合、下の行にPopplerのbinディレクトリへのパスを記述してください。
    # 例: POPPLER_BIN_PATH = r"C:\Users\YourUser\poppler-24.02.0\bin"
    POPPLER_BIN_PATH = r"C:\Users\Shimoyama Naoya\Downloads\Release-25.11.0-0\poppler-25.11.0\Library\bin"
    
    # ★★★★ ページ範囲の指定 ★★★★
    # 例 1: 1ページ目から3ページ目までを変換する場合
    START_PAGE = 376
    END_PAGE = 400
    
    # 例 2: 5ページ目以降すべてを変換する場合
    # START_PAGE = 5
    # END_PAGE = None 
    
    # 例 3: 2ページ目だけを変換する場合
    # START_PAGE = 2
    # END_PAGE = 2


    if not os.path.exists(INPUT_PDF):
        print(f"\nエラー: 入力ファイル '{INPUT_PDF}' が見つかりません。")
    # Windowsでpoppler_pathが指定されていない、または無効な場合の警告
    elif os.name == 'nt' and not os.path.isdir(POPPLER_BIN_PATH):
        print(f"\n警告: WindowsではPopplerのパス設定が必要です。")
        print(f"指定されたパス '{POPPLER_BIN_PATH}' は有効なディレクトリではありません。")
        print("スクリプト内の 'POPPLER_BIN_PATH' 変数を、展開したPopplerの'bin'フォルダへの正しいパスに更新してください。")
    else:
        convert_pdf_to_images_and_save(
            INPUT_PDF, 
            OUTPUT_FOLDER, 
            start_page=START_PAGE, 
            end_page=END_PAGE, 
            poppler_path=POPPLER_BIN_PATH
        )