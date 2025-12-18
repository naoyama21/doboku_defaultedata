import csv
import io
import sys

def classify_data_from_file(file_path):
    """
    テキストファイルからデータを読み込み、「表」と「単価表」の行に分類する。
    """
    classified_data = {
        "表": [],
        "単価表": []
    }
    
    try:
        # ファイルを開いて読み込む
        with open(file_path, 'r', newline='', encoding='utf-8') as f:
            # csv.readerを使用して、カンマ区切り（CSV形式）として行を読み込む
            # 既存のデータ形式に合わせてクォート文字を指定
            reader = csv.reader(f, delimiter=',', quotechar='"')
            
            for row in reader: # rowはファイル全体ではなく、行ごとの処理が必要
                # データの行ごとに処理
                if len(row) > 1:
                    # 2番目の要素（インデックス1）が表の名前や単価表のタイトル
                    header = row[1]
                    
                    if "単価表" in header:
                        classified_data["単価表"].append(row)
                    elif "表" in header:
                        classified_data["表"].append(row)
        
        return classified_data
        
    except FileNotFoundError:
        return {"error": f"エラー: ファイルが見つかりません。ファイルパスを確認してください: {file_path}"}
    except Exception as e:
        return {"error": f"処理中にエラーが発生しました: {e}"}

def write_to_csv(file_name, data):
    """
    データを指定されたファイル名でCSV形式で書き出す。
    元の形式（ダブルクォーテーションで全て囲む）を維持する。
    """
    try:
        # 新しいファイルを作成して書き込む
        with open(file_name, 'w', newline='', encoding='utf-8') as f:
            # csv.writerを設定
            writer = csv.writer(f, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
            
            # データを書き出し
            writer.writerows(data)
        
        print(f"✅ ファイル出力完了: {file_name} に {len(data)} 行のデータを出力しました。")
        return True
    except Exception as e:
        print(f"❌ ファイル書き込みエラー ({file_name}): {e}")
        return False

# 実行
import os
# スクリプトの場所を基準に絶対パスを構築
script_dir = os.path.dirname(os.path.abspath(__file__))
file_name = os.path.join(script_dir, '../data/第２編土木工事標準歩掛.txt')
result = classify_data_from_file(file_name)

if "error" in result:
    print(result["error"])
else:
    # 1. 「表」のデータをCSVに出力
    table_file = os.path.join(script_dir, "../data/table_data.csv")
    write_to_csv(table_file, result["表"])
    
    # 2. 「単価表」のデータをCSVに出力
    unit_price_file = os.path.join(script_dir, "../data/unit_price_table_data.csv")
    write_to_csv(unit_price_file, result["単価表"])