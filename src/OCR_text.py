import pytesseract
import pandas as pd
import numpy as np
import re
import os
from pytesseract import Output
from PIL import Image

# 定数と設定項目は省略（元のコードを参照）
TABLE_TYPE_HYO = "hyo"
TABLE_TYPE_BETSU_HYO = "betsu_hyo"
TABLE_TYPE_NO_SPECIFIC_HEADER = "no_specific_header_tables" 

# --- 設定項目 (Tesseractのパス設定は省略) ---
TESSERACT_CMD_PATH = None 
if TESSERACT_CMD_PATH:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD_PATH
# ----------------------------------------

# --- ヘルパー関数群 ---

def run_ocr_with_coordinates(image):
    """
    画像に対してOCRを実行し、座標情報を含むDataFrameを返す。
    config='--psm 7' を使用し、文字の細かな分離を抑制する。
    """
    # PSM 7 (単一行モード) を使用
    data = pytesseract.image_to_data(image, lang='jpn', output_type=Output.DATAFRAME, config='--psm 7')
    
    data = data.dropna(subset=['text'])
    data = data[data['conf'] > 30]
    data = data[data['text'].str.strip() != '']
    return data

# --- 新しい保存関数: テキストファイルとして保存 ---

def save_raw_text(ocr_data_df, output_dir, file_base_name):
    """
    OCR結果DataFrameの'text'カラムを抽出し、一つのテキストファイルとして保存する。
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created base output directory: {output_dir}")

    # 'text'カラムの値を改行で結合
    raw_text = "\n".join(ocr_data_df['text'].tolist())
    
    txt_file_name = f"{file_base_name}.txt"
    txt_file_path = os.path.join(output_dir, txt_file_name)
    
    try:
        with open(txt_file_path, 'w', encoding='utf-8') as f:
            f.write(raw_text)
        print(f"✅ Raw text saved to {txt_file_path}")
        return True
    except Exception as e:
        print(f"テキストファイル保存エラー: {e}")
        return False


# --- メイン実行部分 ---

def process_images_in_folder(input_folder_path, output_folder_path):
    """入力フォルダ内のすべての画像ファイルからOCRテキストを抽出し、テキストファイルとして保存する"""
    
    if not os.path.exists(input_folder_path):
        print(f"エラー: 入力フォルダ '{input_folder_path}' が見つかりません。")
        return
        
    supported_extensions = ('.png', '.jpg', '.jpeg', '.tiff', '.bmp')
    image_files = [f for f in os.listdir(input_folder_path) if f.lower().endswith(supported_extensions)]
    
    if not image_files:
        print(f"'{input_folder_path}' 内にサポートされている画像ファイルが見つかりませんでした。")
        return
        
    total_processed_files = 0
    
    print(f"合計 {len(image_files)} 個の画像ファイルを処理します...")
    
    for image_file_name in sorted(image_files):
        image_path = os.path.join(input_folder_path, image_file_name)
        file_base_name = os.path.splitext(image_file_name)[0]
        
        print(f"\nProcessing '{image_file_name}'...")
        
        try:
            img = Image.open(image_path)
            
            # 1. OCRを実行し、座標情報を取得
            ocr_data_df = run_ocr_with_coordinates(img)
            
            if not ocr_data_df.empty:
                # 2. 表構造化ロジックをスキップし、生のテキストデータを保存
                if save_raw_text(ocr_data_df, output_folder_path, file_base_name):
                    total_processed_files += 1
            else:
                print(f"ファイル {image_file_name} には信頼度の高いテキストが見つかりませんでした。")
                
        except Exception as e:
            print(f"ファイル {image_file_name} の処理中にエラーが発生しました: {e}")
            
    print(f"\n✅ 合計 {total_processed_files} 個の画像からテキストが抽出され、テキストファイルとして保存されました。")


if __name__ == "__main__":
    # --- 実行時の設定 ---
    
    # 処理対象の画像ファイルがあるフォルダ
    INPUT_IMAGE_FOLDER = os.path.abspath(os.path.join(os.path.dirname(__file__), "../data/pdf_images"))
    # テキスト出力のベースディレクトリ (CSVフォルダとは異なるフォルダを推奨)
    OUTPUT_BASE_FOLDER = os.path.abspath(os.path.join(os.path.dirname(__file__), "../data/ocr_extracted_raw_text"))
    
    process_images_in_folder(INPUT_IMAGE_FOLDER, OUTPUT_BASE_FOLDER)