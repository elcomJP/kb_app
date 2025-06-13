import os
import pandas as pd
from PyQt5.QtCore import QDate

class DataHandler:
    @staticmethod
    def parse_date(date_str):
        """日付文字列(YYMMDD)をQDateに変換"""
        if not date_str or not str(date_str).isdigit() or len(str(date_str)) != 6:
            return None
        
        date_str = str(date_str)
        year = 2000 + int(date_str[0:2])
        month = int(date_str[2:4])
        day = int(date_str[4:6])
        
        return QDate(year, month, day)
    
    @staticmethod
    def date_to_string(date):
        """QDateをYYMMDD形式の文字列に変換"""
        return f"{date.year() - 2000:02d}{date.month():02d}{date.day():02d}"
    
    @staticmethod
    def load_csv_data(folder_path):
        """フォルダから複数のCSVファイルを読み込み結合する"""
        all_data = []
        csv_count = 0
        
        try:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith('.csv') and 'Count' in file and 'Sale' not in file:
                        file_path = os.path.join(root, file)
                        try:
                            df = pd.read_csv(file_path, encoding='shift-jis', dtype=str)
                            all_data.append(df)
                            csv_count += 1
                            print(f"読み込み成功: {file_path}")
                        except Exception as e:
                            print(f"ファイル読み込みエラー: {file_path}, エラー: {e}")
            
            if not all_data:
                print("有効なCSVファイルが見つかりませんでした")
                return None
            
            combined_data = pd.concat(all_data, ignore_index=True)
            print(f"合計 {csv_count} ファイルを読み込みました。合計 {len(combined_data)} 行のデータ。")
            return combined_data
        
        except Exception as e:
            print(f"フォルダ処理エラー: {e}")
            return None
    
    @staticmethod
    def filter_data_by_date(data, start_date_str, end_date_str, date_column_name):
        """日付範囲でデータをフィルタリング"""
        data[date_column_name] = data[date_column_name].astype(str)
        
        filtered_data = data[
            (data[date_column_name] >= start_date_str) & 
            (data[date_column_name] <= end_date_str)
        ]
        
        print(f"フィルター後のデータ行数: {len(filtered_data)}")
        return filtered_data
    
    @staticmethod
    def create_summary(filtered_data, column_indices):
        """商品別およびグループ別の集計を作成"""
        product_code_index = column_indices["product_code_idx"]
        product_name_index = column_indices["product_name_idx"]
        count_index = column_indices["count_idx"]
        amount_index = column_indices["amount_idx"]
        group_num_index = column_indices["group_num_idx"]
        group_name_index = column_indices["group_name_idx"]
        
        product_code_col = filtered_data.columns[product_code_index]
        product_name_col = filtered_data.columns[product_name_index]
        count_col = filtered_data.columns[count_index]
        amount_col = filtered_data.columns[amount_index]
        group_num_col = filtered_data.columns[group_num_index]
        group_name_col = filtered_data.columns[group_name_index]
        
        amount_sign_col = filtered_data.columns[amount_index - 1]
        card_amount_col = filtered_data.columns[12]
        
        print(f"使用する列: 商品コード={product_code_col}, 商品名称={product_name_col}, " +
            f"枚数={count_col}, 金額={amount_col}, 集計 G 番号={group_num_col}, 集計 G 名称={group_name_col}, " +
            f"金額符号={amount_sign_col}, カード減算額={card_amount_col}")
        
        filtered_data = filtered_data.copy()
        if not pd.api.types.is_numeric_dtype(filtered_data[count_col]):
            filtered_data[count_col] = pd.to_numeric(filtered_data[count_col], errors='coerce')
        if not pd.api.types.is_numeric_dtype(filtered_data[amount_col]):
            filtered_data[amount_col] = pd.to_numeric(filtered_data[amount_col], errors='coerce')
        if not pd.api.types.is_numeric_dtype(filtered_data[card_amount_col]):
            filtered_data[card_amount_col] = pd.to_numeric(filtered_data[card_amount_col], errors='coerce')
        
        valid_data = filtered_data[filtered_data[amount_sign_col] != '1'].copy()
        
        if not valid_data.empty:
            product_summary = valid_data.groupby([product_code_col, product_name_col], as_index=False).agg({
                count_col: 'sum',
                amount_col: 'sum'
            })
        else:
            product_summary = pd.DataFrame(columns=[product_code_col, product_name_col, count_col, amount_col])
        
        if not valid_data.empty:
            group_summary = valid_data.groupby([group_num_col, group_name_col], as_index=False).agg({
                count_col: 'sum',
                amount_col: 'sum'
            })
        else:
            group_summary = pd.DataFrame(columns=[group_num_col, group_name_col, count_col, amount_col])
        
        total_count = valid_data[count_col].sum() if not valid_data.empty else 0
        total_amount = valid_data[amount_col].sum() if not valid_data.empty else 0
        
        cashless_count = valid_data[valid_data[card_amount_col] > 0][count_col].sum() if not valid_data.empty else 0
        cashless_amount = valid_data[valid_data[card_amount_col] > 0][card_amount_col].sum() if not valid_data.empty else 0
        
        print(f"商品別集計結果行数: {len(product_summary)}")
        print(f"グループ別集計結果行数: {len(group_summary)}")
        
        return {
            "product_summary": product_summary,
            "group_summary": group_summary,
            "total_count": total_count,
            "total_amount": total_amount,
            "cashless_count": cashless_count,
            "cashless_amount": cashless_amount
        }