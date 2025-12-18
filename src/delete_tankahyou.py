import csv
import os

def process_csv_row_conditional(row_data):
    """
    CSVの1行データ（リスト）を受け取り、「単価表」が2つある場合に、
    行全体で見て最初に出現する「単価表」のみを削除して返す関数。
    """
    # 1. 行全体を一時的に一つの文字列として結合（CSVはカンマ区切りで結合）
    # セルに文字列以外のデータが含まれる可能性があるため、str()でキャスト
    line_str = ','.join([str(item) for item in row_data])
    
    # 2. 行全体での「単価表」の出現回数をカウント
    count = line_str.count('単価表')
    
    modified_line_str = line_str
    
    # 3. 条件分岐による削除処理
    if count == 2:
        # 出現回数がちょうど2回の場合、最初に出現するものを削除（置換回数1）
        modified_line_str = line_str.replace('単価表', '', 1)
        
    # 4. 変更された（または変更されなかった）文字列を再びCSVの行（リスト）に戻す
    # 注意: この処理により、元のセルの区切りが正確に再現されない場合があります。
    # 厳密なCSV構造を維持するには複雑な処理が必要ですが、ここではシンプルな方法として
    # 変更後の文字列をカンマで再分割します。
    # 元のリストの要素数と一致しない場合があるため、処理後のデータ形式に注意してください。
    
    # 今回は元のCSVのセルの数や内容が正確に保たれることを重視し、
    # セル単位の処理で「単価表」が存在するセルを特定し、そのセルの内容を修正する
    # 別のアプローチを取ります。
    
    # --- セル単位でのロジックの再構築 ---
    # 元の形式を保ちつつ、行全体で2回出現する場合のみ最初の一つを削除するロジック
    
    # 全セルの文字列を結合し、合計の出現回数を計算
    temp_line_str = ','.join([str(item) for item in row_data])
    total_count = temp_line_str.count('単価表')
    
    if total_count != 2:
        # 2回でなければそのまま元の行を返す
        return row_data
        
    # 2回の場合: 行全体で最初に出現する「単価表」を削除する
    
    # どのセルに「単価表」が含まれているかを記憶する
    cell_index_to_modify = -1
    
    current_index = 0
    
    # 「単価表」が見つかるまで、各セルとその区切り（カンマ）の長さを加算していく
    for i, item in enumerate(row_data):
        item_str = str(item)
        
        # セル内に「単価表」があるか
        if '単価表' in item_str:
            # セル内の出現位置を検出
            first_occurrence_in_cell = item_str.find('単価表')
            
            # 行全体の最初に出現する「単価表」がこのセル内にあると確定
            # (total_count == 2 の前提で、最初に見つけたものが削除対象)
            
            # そのセルを修正し、処理完了
            row_data[i] = item_str.replace('単価表', '', 1)
            return row_data
            
        # 次のセルへ（カンマ区切りを含む）
        # ただし、行全体を結合せず、リスト操作だけで処理を完結させるのが最も安全です。
    
    # 基本的に total_count == 2 であれば、必ず上記の for ループ内で return されるため、
    # ここに到達することはありませんが、万が一のために元の行を返します。
    return row_data

# --- ファイル操作部分 ---

# ファイル名設定
# スクリプトの場所を基準に絶対パスを構築
script_dir = os.path.dirname(os.path.abspath(__file__))
# ユーザー提供のファイルパスを使用
input_file = os.path.join(script_dir, '../data/unit_price_table_data.csv')
output_file = os.path.join(script_dir, '../data/unit_price_table_data_processed.csv')

# CSVファイルの処理実行
try:
    with open(input_file, 'r', newline='', encoding='utf-8') as infile, \
         open(output_file, 'w', newline='', encoding='utf-8') as outfile:
        
        # readerとwriterの作成
        reader = csv.reader(infile)
        writer = csv.writer(outfile)
        
        # 1行ずつ処理
        for row in reader:
            # 修正された処理関数を呼び出し
            processed_row = process_csv_row_conditional(row)
            # 処理後の行を新しいファイルに書き込む
            writer.writerow(processed_row)
            
    print(f"✅ 処理が完了しました。'{input_file}' から読み込み、'{output_file}' に書き出しました。")

except FileNotFoundError:
    print(f"⚠️ エラー：'{input_file}' が見つかりません。ファイルが存在するか確認してください。")
except Exception as e:
    print(f"🛑 予期せぬエラーが発生しました: {e}")