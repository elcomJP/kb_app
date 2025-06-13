import pandas as pd
import os
import re

class DataProcessor:
    """売上データの読み込みと前処理を行うクラス"""
    
    def __init__(self):
        pass
    
    def load_data(self, file_path):
        """ファイルからデータを読み込む
        
        Args:
            file_path (str): 読み込むファイルのパス
            
        Returns:
            DataFrame: 読み込んだデータ
        """
        # ファイル拡張子によって処理を分岐
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.csv':
            # CSVファイルの読み込み
            try:
                # エンコーディングと区切り文字を自動推定
                with open(file_path, 'rb') as f:
                    data = f.read()
                    
                # 日本語CSVではShift-JISが多い
                encodings = ['shift_jis', 'utf-8', 'cp932']
                
                for encoding in encodings:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding)
                        return df
                    except UnicodeDecodeError:
                        continue
                    except Exception as e:
                        print(f"CSVファイル読み込みエラー: {e}")
                        
                # すべてのエンコーディングで失敗した場合
                print("CSVファイルのエンコーディングを特定できませんでした。")
                return None
                
            except Exception as e:
                print(f"CSVファイル読み込みエラー: {e}")
                return None
                
        elif file_ext == '.xlsx' or file_ext == '.xls':
            # Excelファイルの読み込み
            try:
                df = pd.read_excel(file_path)
                return df
            except Exception as e:
                print(f"Excelファイル読み込みエラー: {e}")
                return None
        else:
            print(f"未対応のファイル形式です: {file_ext}")
            return None
    
    def preprocess_data(self, df):
        """データの前処理を行う
        
        Args:
            df (DataFrame): 前処理対象のデータフレーム
            
        Returns:
            DataFrame: 前処理後のデータフレーム
        """
        if df is None or df.empty:
            return df
            
        # データのコピーを作成
        processed_df = df.copy()
        
        # 日付列の処理（取引日付の列を想定）
        date_column_index = 13  # app.pyと合わせる
        if len(processed_df.columns) > date_column_index:
            date_column = processed_df.columns[date_column_index]
            
            # 文字列型に変換
            processed_df[date_column] = processed_df[date_column].astype(str)
            
            # 日付形式を統一（YYYY/MM/DD）
            def format_date(date_str):
                if pd.isna(date_str) or date_str == '':
                    return ''
                    
                # 年月日を抽出
                patterns = [
                    r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',  # YYYY-MM-DD or YYYY/MM/DD
                    r'(\d{2})[/-](\d{1,2})[/-](\d{1,2})'    # YY-MM-DD or YY/MM/DD
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, date_str)
                    if match:
                        year, month, day = match.groups()
                        
                        # 2桁の年の場合は2000年代と仮定
                        if len(year) == 2:
                            year = f"20{year}"
                            
                        # ゼロ埋め
                        month = month.zfill(2)
                        day = day.zfill(2)
                        
                        return f"{year}/{month}/{day}"
                
                return date_str
                
            processed_df[date_column] = processed_df[date_column].apply(format_date)
        
        # 数値列の処理
        # 金額と数量の列を想定
        numeric_columns = [9, 11]  # 枚数/数量と金額の列インデックス
        
        for idx in numeric_columns:
            if len(processed_df.columns) > idx:
                col = processed_df.columns[idx]
                
                # NaNや文字列を0に置換
                processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce').fillna(0)
                
                # 整数型に変換
                processed_df[col] = processed_df[col].astype(int)
        
        return processed_df