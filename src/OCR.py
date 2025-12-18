import pytesseract
import pandas as pd
import numpy as np
import re
import os
from pytesseract import Output
from PIL import Image

# 定数としてテーブルタイプを定義 (前のコードから流用)
TABLE_TYPE_HYO = "hyo"
TABLE_TYPE_BETSU_HYO = "betsu_hyo"
TABLE_TYPE_NO_SPECIFIC_HEADER = "no_specific_header_tables" 

# --- 設定項目 (実行環境に合わせて変更) ---
# Windowsの場合、Tesseractの実行ファイルのパスを設定 (適切なパスを設定済みの前提)
TESSERACT_CMD_PATH = None 
if TESSERACT_CMD_PATH:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD_PATH

# -------------------------------------------------------------------
# --- ヘルパー関数群 ---
# -------------------------------------------------------------------

def run_ocr_with_coordinates(image):
    """
    画像に対してOCRを実行し、座標情報を含むDataFrameを返す。
    PSM (Page Segmentation Mode) 7 を使用し、単行モードでの認識を試みることで、
    文字の細かな分離を抑制する。
    """
    # ★ 修正点: config='--psm 7' を追加 ★
    # psm 7: Treat the image as a single text line. (表の行認識に有効な場合がある)
    data = pytesseract.image_to_data(image, lang='jpn', output_type=Output.DATAFRAME, config='--psm 7')
    
    data = data.dropna(subset=['text'])
    data = data[data['conf'] > 30]
    data = data[data['text'].str.strip() != '']
    return data

def create_dataframe_from_ocr_rows(rows_data, hyo_text, sagyome_text, note_text):
    """再構築された行データとメタデータを結合し、整形されたDataFrameを作成する。"""
    if not rows_data:
        return None

    num_cols = len(rows_data[0])
    headers = [f"Col_{i+1}" for i in range(num_cols)]

    final_headers = []
    
    if hyo_text:
        final_headers.append('表')
    if sagyome_text:
        final_headers.append('作業名')
    
    final_headers.extend(headers)
    final_headers.append('注')
    
    processed_rows_data = []
    for row in rows_data:
        current_row_processed = []
        if hyo_text:
            current_row_processed.append(hyo_text)
        if sagyome_text:
            current_row_processed.append(sagyome_text)
        
        current_row_processed.extend(row)
        current_row_processed.append(note_text)
        processed_rows_data.append(current_row_processed)
            
    max_cols = max(len(h) for h in final_headers)
    if processed_rows_data:
        max_cols = max(max_cols, max(len(r) for r in processed_rows_data))
    
    final_headers = final_headers + [''] * (max_cols - len(final_headers))

    adjusted_rows_data = []
    for row_data_entry in processed_rows_data:
        adjusted_rows_data.append(row_data_entry + [''] * (max_cols - len(row_data_entry)))

    df = pd.DataFrame(adjusted_rows_data, columns=final_headers)
    return df

def cluster_and_reconstruct_table(ocr_data_df, y_tolerance=10, x_tolerance=80): # ★ 修正点: x_toleranceを80に変更 ★
    """
    TesseractのOCR座標情報DataFrameを基に、表を再構築し、メタデータ（表/注記など）を抽出する。
    """
    if ocr_data_df.empty:
        return []

    # 1. Y座標 (top) を基に行をクラスタリング
    df_sorted = ocr_data_df.sort_values(by=['top', 'left']).reset_index(drop=True)
    line_groups = [0]
    current_group_index = 0
    for i in range(1, len(df_sorted)):
        if df_sorted.loc[i, 'top'] - df_sorted.loc[i , 'top'] > y_tolerance:
            current_group_index += 1
        line_groups.append(current_group_index)
    df_sorted['line_group'] = line_groups
    
    # 2. X座標 (left) を基に列の境界線 (Xクラスター) を検出
    x_coords = df_sorted['left'].values
    column_x_clusters = []
    x_coords_sorted = np.sort(x_coords)
    if x_coords_sorted.size > 0:
        column_x_clusters.append(x_coords_sorted[0])
        for x in x_coords_sorted[1:]:
            if x - column_x_clusters[ ] > x_tolerance:
                column_x_clusters.append(x)
            
    # 3. 単語を検出した列に割り当てる
    def assign_column(x_coord):
        distances = np.abs(np.array(column_x_clusters) - x_coord)
        return np.argmin(distances)
    df_sorted['column_index'] = df_sorted['left'].apply(assign_column)

    # 4. グループ化とメタデータ/サブテーブルの検出 (元のロジックを維持)
    extracted_dfs_info = []
    current_table_rows = []
    current_hyo_text = ""
    current_sagyome_text = ""
    current_note_text = ""
    is_note_section = False
    
    for group_id, group_df in df_sorted.groupby('line_group'):
        row_cells = {}
        for col_idx, text_list in group_df.groupby('column_index')['text']:
            row_cells[col_idx] = ' '.join(text_list)
            
        row_data = [row_cells.get(col_idx, '') for col_idx in range(len(column_x_clusters))]
        combined_row_text = "".join(row_data).strip()
        first_cell_text = row_data[0]
        
        # --- メタデータ検出ロジック ---
        
        if is_note_section:
            if combined_row_text:
                current_note_text += "\n" + combined_row_text.replace("(注)", "").strip()
            continue
            
        if "(注)" in combined_row_text:
            if current_table_rows:
                df = create_dataframe_from_ocr_rows(current_table_rows, current_hyo_text, current_sagyome_text, current_note_text)
                if df is not None:
                     table_type = TABLE_TYPE_HYO if re.match(r'^表', current_hyo_text) else (TABLE_TYPE_BETSU_HYO if re.match(r'^別表', current_hyo_text) else TABLE_TYPE_NO_SPECIFIC_HEADER)
                     extracted_dfs_info.append({'df': df, 'table_type': table_type})
                
                current_table_rows = []
                current_note_text = ""
                
            is_note_section = True
            current_note_text = combined_row_text.replace("(注)", "").strip()
            continue
            
        if re.match(r'^(表|別表)', first_cell_text):
            if current_table_rows:
                df = create_dataframe_from_ocr_rows(current_table_rows, current_hyo_text, current_sagyome_text, current_note_text)
                if df is not None:
                     table_type = TABLE_TYPE_HYO if re.match(r'^表', current_hyo_text) else (TABLE_TYPE_BETSU_HYO if re.match(r'^別表', current_hyo_text) else TABLE_TYPE_NO_SPECIFIC_HEADER)
                     extracted_dfs_info.append({'df': df, 'table_type': table_type})
                
                current_table_rows = []
            
            current_hyo_text = first_cell_text
            current_sagyome_text = "".join(row_data[1:]).strip() 
            current_note_text = ""
            continue
            
        if combined_row_text:
            current_table_rows.append(row_data)

    if current_table_rows:
        df = create_dataframe_from_ocr_rows(current_table_rows, current_hyo_text, current_sagyome_text, current_note_text)
        if df is not None:
             table_type = TABLE_TYPE_HYO if re.match(r'^表', current_hyo_text) else (TABLE_TYPE_BETSU_HYO if re.match(r'^別表', current_hyo_text) else TABLE_TYPE_NO_SPECIFIC_HEADER)
             extracted_dfs_info.append({'df': df, 'table_type': table_type})
             
    return extracted_dfs_info

def save_dfs_to_csv(dfs_info, output_dir, file_base_name):
    """
    DataFrameのリストを、表のタイプに基づいたサブディレクトリに個別のCSVファイルとして保存する。
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created base output directory: {output_dir}")

    for i, df_info in enumerate(dfs_info):
        df = df_info['df']
        table_type = df_info['table_type'] 

        specific_output_dir = os.path.join(output_dir, table_type)

        if not os.path.exists(specific_output_dir):
            os.makedirs(specific_output_dir)
            print(f"Created output directory: {specific_output_dir}")

        # ファイル名は 「元ファイル名ベース-表の連番.csv」
        csv_file_name = f"{file_base_name}-{i+1}.csv"
        csv_file_path = os.path.join(specific_output_dir, csv_file_name)
        
        df.to_csv(csv_file_path, index=False, header=True, encoding='utf-8')
        print(f"Table from file {file_base_name}, sub-table {i+1} saved to {csv_file_path} in '{table_type}' folder.")


# -------------------------------------------------------------------
# --- メイン実行部分 ---
# -------------------------------------------------------------------

def process_images_in_folder(input_folder_path, output_folder_path):
    """入力フォルダ内のすべての画像ファイルから表を抽出し、CSVとして保存する"""
    
    if not os.path.exists(input_folder_path):
        print(f"エラー: 入力フォルダ '{input_folder_path}' が見つかりません。")
        return
        
    supported_extensions = ('.png', '.jpg', '.jpeg', '.tiff', '.bmp')
    image_files = [f for f in os.listdir(input_folder_path) if f.lower().endswith(supported_extensions)]
    
    if not image_files:
        print(f"'{input_folder_path}' 内にサポートされている画像ファイルが見つかりませんでした。")
        return
        
    total_extracted_tables = 0
    
    print(f"合計 {len(image_files)} 個の画像ファイルを処理します...")
    
    for image_file_name in sorted(image_files):
        image_path = os.path.join(input_folder_path, image_file_name)
        file_base_name = os.path.splitext(image_file_name)[0]
        
        print(f"\nProcessing '{image_file_name}'...")
        
        try:
            # 1. 画像ファイルを読み込む
            img = Image.open(image_path)
            
            # 2. OCRを実行し、座標情報を取得
            ocr_data_df = run_ocr_with_coordinates(img)
            
            if not ocr_data_df.empty:
                # 3. 座標情報から表を再構築し、メタデータを抽出
                extracted_tables_with_types = cluster_and_reconstruct_table(ocr_data_df)
                
                if extracted_tables_with_types:
                    # 4. CSVとして保存 (ファイル名をベースとして使用)
                    save_dfs_to_csv(extracted_tables_with_types, output_folder_path, file_base_name)
                    total_extracted_tables += len(extracted_tables_with_types)
                else:
                    print(f"ファイル {image_file_name} から表構造は見つかりませんでした。")
            else:
                print(f"ファイル {image_file_name} には信頼度の高いテキストが見つかりませんでした。")
                
        except Exception as e:
            # Tesseractが見つからないエラーもここでキャッチされます
            print(f"ファイル {image_file_name} の処理中にエラーが発生しました: {e}")
            
    if total_extracted_tables > 0:
        print(f"\n✅ 合計 {total_extracted_tables} 個の表がCSVファイルとして保存されました。")
    else:
        print("\nすべての画像ファイルから表は見つかりませんでした。")


if __name__ == "__main__":
    # --- 実行時の設定 ---
    
    # 処理対象の画像ファイルがあるフォルダ
    # Windowsのユーザーフォルダ内のパスは、Pythonの実行時には 'C:\Users\Naoya\OneDrive\...' のようにフルパスで指定した方が安全です。
    INPUT_IMAGE_FOLDER = os.path.abspath(os.path.join(os.path.dirname(__file__), "../data/pdf_images"))
    # CSV出力のベースディレクトリ
    OUTPUT_BASE_FOLDER = os.path.abspath(os.path.join(os.path.dirname(__file__), "../data/ocr_extracted_tables_from_images"))
    
    # INPUT_IMAGE_FOLDER に画像ファイルを配置してから実行してください
    process_images_in_folder(INPUT_IMAGE_FOLDER, OUTPUT_BASE_FOLDER)