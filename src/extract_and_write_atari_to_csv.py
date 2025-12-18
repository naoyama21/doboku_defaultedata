import csv
import os

def extract_and_write_atari_to_csv(input_filename, output_filename):
    """
    指定されたファイルからデータを読み込み、「当り」を含む行を抽出して、
    新しいCSVファイルに書き出します。

    Args:
        input_filename (str): 入力ファイル名。
        output_filename (str): 出力ファイル名。
    """
    
    # 処理するCSVデータのフィールド定義 (ヘッダー行を想定)
    # 抽出ロジックはヘッダー行の有無にかかわらず機能しますが、出力にはヘッダーも追加します。
    FIELD_NAMES = [
        "章・工種", "表名", "項目1", "項目2", "項目3", "項目4", 
        "規格", "単位", "数量", "摘要", "数量2", "数量3", 
        "数量4", "数量5"
    ]

    try:
        # 1. 入力ファイルを読み込み
        with open(input_filename, 'r', newline='', encoding='utf-8') as infile:
            reader = csv.reader(infile)
            
            # ヘッダー行をスキップしつつ格納
            header = next(reader, None)
            
            # 抽出した行を格納するリスト
            result_rows = []
            
            # ヘッダー行に「当り」が含まれるかチェック（通常はないが念のため）
            if header and ("当り" in header[0] or "当り" in header[1]):
                 result_rows.append(header)
            
            # 2. データを1行ずつ読み込み、抽出条件をチェック
            for row in reader:
                if not row:
                    continue
                
                # インデックス0: 章・工種, インデックス1: 表名
                # フィールド数が2未満の場合はスキップ（エラー回避）
                if len(row) < 2:
                    continue
                    
                chapter_col = row[0]
                table_name_col = row[1]
                
                # '章・工種'または'表名'に「当り」が含まれているかチェック
                if "単価表" in chapter_col or "単価表" in table_name_col:
                    result_rows.append(row)

        if not result_rows:
            print("▶︎ 抽出条件（「当り」を含む行）に合致するデータは見つかりませんでした。")
            return
            
        # 3. 抽出した行を出力ファイルに書き出し
        with open(output_filename, 'w', newline='', encoding='utf-8') as outfile:
            writer = csv.writer(outfile)
            
            # ヘッダーを書き込む（元のヘッダーがない場合は定義したヘッダーを使用）
            writer.writerow(header or FIELD_NAMES)

            # 抽出された行を書き込む
            writer.writerows(result_rows)
            
        print(f"✅ 処理が完了しました。")
        print(f"   入力ファイル: {input_filename}")
        print(f"   抽出行数: {len(result_rows)}")
        print(f"   出力ファイル: {output_filename}")


    except FileNotFoundError:
        print(f"❌ エラー: 入力ファイル '{input_filename}' が見つかりません。")
        print("   ファイルが存在することを確認してください。")
    except Exception as e:
        print(f"❌ 処理中に予期せぬエラーが発生しました: {e}")

# --- 実行部分 ---
if __name__ == "__main__":
    # スクリプトの絶対パスを取得
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 入力ファイルと出力ファイルのパスを構築
    input_file = os.path.join(script_dir, '..', 'data', '第２編土木工事標準歩掛＿道路関連.txt')
    output_file = os.path.join(script_dir, '..', 'data', 'doboku_csv', 'output_result.csv')
    
    # メイン関数を実行して、実際のファイルを処理
    extract_and_write_atari_to_csv(input_file, output_file)