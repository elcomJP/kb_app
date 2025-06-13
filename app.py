import os
import pandas as pd
import matplotlib.pyplot as plt

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,QGroupBox, 
                            QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, 
                            QFileDialog, QDateEdit, QTabWidget, QScrollArea, QHeaderView, QComboBox, QMessageBox,QProgressDialog)
from PyQt5.QtCore import Qt, QDate, QSettings, QTimer
from PyQt5.QtGui import QIcon, QPixmap, QColor
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

from widgets import NumericTableWidgetItem
from data_handler import DataHandler
from utils import DateUtils
from export_handler import ExportHandler


class SalesAnalysisApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("KB Series 売上集計")
        self.setGeometry(100, 100, 1400, 900)
        
        # 設定ファイルの初期化
        self.settings = QSettings("KBSeries", "SalesAnalysis")

        # 設定ファイルから以前のフォルダパスを取得
        self.last_folder_path = self.settings.value("last_folder_path", "")
        
        # メインウィジェットとレイアウト
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.main_layout = QVBoxLayout(self.main_widget)
        
        # スクロールエリアの設定
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        self.main_layout.addWidget(self.scroll_area)
        
        # データ保存用変数
        self.csv_data = None
        self.last_summary = None 
        
        # 初期ソート設定
        self.sort_column = 0  # 商品コード列を初期ソート
        self.sort_order = Qt.AscendingOrder  # 昇順
        self.sort_column_g = 0  # グループ番号列を初期ソート
        self.sort_order_g = Qt.AscendingOrder  # 昇順
        self.sort_column_r = 0  # 伝票別テーブル用初期ソート（商品コード列）
        self.sort_order_r = Qt.AscendingOrder  # 昇順
        
        self.column_indices = {
            "date_column_index": 13,  # 取引日付の列インデックス
            "group_num_idx": 18,      # 集計G番号
            "group_name_idx": 19,     # 集計G名称
            "product_code_idx": 16,   # 商品コード/メニュー番号
            "product_name_idx": 17,   # 論理口座名称/メニュー名
            "count_idx": 9,           # 枚数/数量
            "amount_idx": 11,         # 金額
            "amount_sign_idx": 10,    # 金額符号（金額・枚数）
            "card_deduction_idx": 12  # カード減算額
        }
        
        # エクスポートハンドラの初期化
        self.export_handler = ExportHandler() 
        
        # UIの初期化
        self.init_ui()
    
    def init_ui(self):
        # タイトル
        title_label = QLabel("券売機システム")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #2c3e50; margin: 10px;")
        title_label.setAlignment(Qt.AlignCenter)
        self.scroll_layout.addWidget(title_label)
        
        # 上部操作パネル
        self.create_control_panel()

        # タブウィジェット
        self.create_tab_widget()

        # 下部統計パネル
        self.create_statistics_panel()
        
    
    def create_control_panel(self):
        """操作パネルを作成"""
        control_group = QGroupBox("操作パネル")
        control_layout = QVBoxLayout()

        # 店舗名入力エリア
        shop_layout = QHBoxLayout()
        shop_label = QLabel("店舗名:")
        self.shop_input = QLineEdit()
        saved_shop = self.settings.value("shop_name", "")
        self.shop_input.setText(saved_shop)
        self.shop_input.textChanged.connect(self.save_shop_name)
        shop_layout.addWidget(shop_label)
        shop_layout.addWidget(self.shop_input)
        shop_layout.addStretch()
        control_layout.addLayout(shop_layout)
        
        # フォルダ選択エリア
        folder_layout = QHBoxLayout()
        folder_label = QLabel("CSVフォルダ:")
        self.folder_path = QLineEdit()
        self.folder_path.setReadOnly(True)

        # 前回のフォルダパスがあれば表示
        if self.last_folder_path:
            self.folder_path.setText(self.last_folder_path)
        browse_button = QPushButton("参照...")
        browse_button.clicked.connect(self.browse_folder)
        folder_layout.addWidget(folder_label)
        folder_layout.addWidget(self.folder_path)
        folder_layout.addWidget(browse_button)
        control_layout.addLayout(folder_layout)
        
        # 日付検索エリア
        date_layout = QHBoxLayout()
        date_label = QLabel("日付範囲:")
        self.start_date = QDateEdit(QDate.currentDate())
        self.start_date.setDisplayFormat("yyyy/MM/dd")
        self.start_date.setCalendarPopup(True)
        self.start_date.setFixedWidth(120)  

        date_to_label = QLabel("～")

        self.end_date = QDateEdit(QDate.currentDate())
        self.end_date.setDisplayFormat("yyyy/MM/dd")
        self.end_date.setCalendarPopup(True) 
        self.end_date.setFixedWidth(120) 

        # 日付設定ボタン
        today_button = QPushButton("当日")
        today_button.clicked.connect(self.set_today)
        month_button = QPushButton("今月")
        month_button.clicked.connect(self.set_this_month)
        year_button = QPushButton("年間")
        year_button.clicked.connect(self.set_this_year)
        search_button = QPushButton("検索")
        search_button.clicked.connect(self.search_data)
        search_button.setStyleSheet("QPushButton { background-color: #4a86e8; color: white; font-weight: bold; }")

        # キーボード入力を有効化
        self.start_date.setKeyboardTracking(True)
        self.end_date.setKeyboardTracking(True)
            
        date_layout.addWidget(date_label)
        date_layout.addWidget(self.start_date)
        date_layout.addWidget(date_to_label)
        date_layout.addWidget(self.end_date)
        date_layout.addWidget(today_button)
        date_layout.addWidget(month_button)
        date_layout.addWidget(year_button) 
        date_layout.addStretch()  
        date_layout.addWidget(search_button)
        
        control_layout.addLayout(date_layout)
        
        # エクスポート機能
        export_layout = QHBoxLayout()
        export_label = QLabel("データエクスポート:")
        self.export_type = QComboBox()
        self.export_type.addItems(["Excel", "PDF"])
        export_button = QPushButton("エクスポート")
        export_button.clicked.connect(self.export_data)   
        export_layout.addWidget(export_label)
        export_layout.addWidget(self.export_type)
        export_layout.addWidget(export_button)
        export_layout.addStretch()
        
        control_layout.addLayout(export_layout)
        
        control_group.setLayout(control_layout)
        self.scroll_layout.addWidget(control_group)
        
    def create_statistics_panel(self):
        # 統計情報グループボックス
        statistics_group = QGroupBox("統計情報")
        statistics_layout = QVBoxLayout()
        
        # 合計表示エリア
        total_layout = QHBoxLayout()
        self.total_count_label = QLabel("合計枚数: 0")
        self.total_amount_label = QLabel("合計売上: 0円")
        self.cashless_count_label = QLabel("【 キャッシュレス枚数: 0")
        self.cashless_amount_label = QLabel("キャッシュレス売上: 0円 】")

        total_layout.addWidget(self.total_count_label)
        total_layout.addWidget(self.total_amount_label)
        total_layout.addWidget(self.cashless_count_label)
        total_layout.addWidget(self.cashless_amount_label)
        total_layout.addStretch()
        
        statistics_layout.addLayout(total_layout)
        statistics_group.setLayout(statistics_layout)
        
        self.scroll_layout.addWidget(statistics_group)
        
        label_style = "font-size: 14px; font-weight: bold; color: #2c3e50; margin: 5px;"
        self.total_count_label.setStyleSheet(label_style)
        self.total_amount_label.setStyleSheet(label_style)
        self.cashless_count_label.setStyleSheet(label_style)
        self.cashless_amount_label.setStyleSheet(label_style)
        
    def create_tab_widget(self):
        # 表示切り替えタブ
        self.tab_widget = QTabWidget()

        # タブのスタイル設定
        self.tab_style = """
        QTabBar::tab {
            background-color: #f0f0f0;
            border: 1px solid #c0c0c0;
            padding: 8px 16px;
            margin-right: 2px;
        }
        QTabBar::tab:selected {
            background-color: #4a86e8;
            color: white;
            border: 1px solid #3a76d8;
        }
        """
        self.tab_widget.setStyleSheet(self.tab_style)
            
        # 商品別タブ
        self.product_tab = QWidget()
        product_layout = QVBoxLayout(self.product_tab)
        
        # 商品別テーブルを作成
        self.product_table = QTableWidget()
        self.product_table.setColumnCount(4)
        self.product_table.setHorizontalHeaderLabels(["商品コード", "商品名称", "枚数", "金額"])
        self.product_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.product_table.horizontalHeader().sectionClicked.connect(self.sort_product_table)
        product_layout.addWidget(self.product_table)

        # 集計グループ別タブ
        self.group_tab = QWidget()
        group_layout = QVBoxLayout(self.group_tab)
        
        # グループ別テーブルを作成
        self.group_table = QTableWidget()
        self.group_table.setColumnCount(4)
        self.group_table.setHorizontalHeaderLabels(["グループ番号", "グループ名称", "枚数", "金額"])
        self.group_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.group_table.horizontalHeader().sectionClicked.connect(self.sort_group_table)
        group_layout.addWidget(self.group_table)

        # 伝票別タブ（修正）
        self.receipt_tab = QWidget()
        receipt_layout = QVBoxLayout(self.receipt_tab)

        # 伝票種別選択と合計表示のレイアウト
        receipt_control_layout = QHBoxLayout()
        receipt_type_label = QLabel("伝票種別:")
        self.receipt_type_combo = QComboBox()
        self.receipt_type_combo.addItems(["現金売上", "キャッシュレス決済", "赤伝処理"])
        self.receipt_type_combo.currentTextChanged.connect(self.update_receipt_detail)
        
        # 伝票別合計表示ラベル（新規追加）
        self.receipt_total_count_label = QLabel("合計枚数: 0")
        self.receipt_total_amount_label = QLabel("合計金額: 0円")
        
        # ラベルのスタイル設定
        receipt_label_style = "font-size: 12px; font-weight: bold; color: #2c3e50; margin-left: 20px;"
        self.receipt_total_count_label.setStyleSheet(receipt_label_style)
        self.receipt_total_amount_label.setStyleSheet(receipt_label_style)
        
        receipt_control_layout.addWidget(receipt_type_label)
        receipt_control_layout.addWidget(self.receipt_type_combo)
        receipt_control_layout.addWidget(self.receipt_total_count_label)  # 新規追加
        receipt_control_layout.addWidget(self.receipt_total_amount_label)  # 新規追加
        receipt_control_layout.addStretch()
        receipt_layout.addLayout(receipt_control_layout)

        # 伝票別詳細テーブル（ソート機能追加）
        self.receipt_detail_table = QTableWidget()
        self.receipt_detail_table.setColumnCount(4)
        self.receipt_detail_table.setHorizontalHeaderLabels(["商品コード", "商品名称", "枚数", "金額"])
        self.receipt_detail_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.receipt_detail_table.horizontalHeader().sectionClicked.connect(self.sort_receipt_table)  # 新規追加
        receipt_layout.addWidget(self.receipt_detail_table)

        # タブに追加
        self.tab_widget.addTab(self.product_tab, "商品別")
        self.tab_widget.addTab(self.group_tab, "グループ別")
        self.tab_widget.addTab(self.receipt_tab, "伝票別")

        self.scroll_layout.addWidget(self.tab_widget)
    
    def export_data(self):
        """選択したフォーマットでデータをエクスポート"""
        if self.csv_data is None or self.csv_data.empty:
            QMessageBox.information(self, "エクスポート", "エクスポートするデータがありません")
            return
        
        export_type = self.export_type.currentText()
        
        # プログレスダイアログを作成
        progress = QProgressDialog(f"{export_type}形式でエクスポート中...", "キャンセル", 0, 100, self)
        progress.setWindowTitle("データエクスポート")
        progress.setWindowModality(Qt.WindowModal)
        progress.setValue(0)
        progress.show()
        
        try:
            progress.setLabelText("CSVファイルを読み込んでいます...")
            progress.setValue(20)
            
            shop_name = self.shop_input.text()
            if not shop_name:
                shop_name = "KB Series"
                
            export_type_lower = export_type.lower()
            
            progress.setLabelText("データをフィルタリング中...")
            progress.setValue(40)
            
            # フィルタリングされたデータを取得
            start_date_str = DataHandler.date_to_string(self.start_date.date())
            end_date_str = DataHandler.date_to_string(self.end_date.date())
            
            # 列名を取得して日付列のインデックスを確認
            column_names = self.csv_data.columns.tolist()
            date_column_index = 13  # 取引日付の列インデックス
            date_column_name = self.csv_data.columns[date_column_index]
            
            # データ型を確認して強制的に文字列に変換
            date_data = self.csv_data[date_column_name].astype(str)
            
            progress.setLabelText("表示データを準備中...")
            progress.setValue(60)
            
            # 日付フィルタリング
            filtered_data = self.csv_data[
                (date_data >= start_date_str) & 
                (date_data <= end_date_str)
            ]
            
            if filtered_data.empty:
                progress.close()
                QMessageBox.information(self, "エクスポート", "エクスポートするデータがありません")
                return
            
            progress.setLabelText("エクスポート変換中...")
            progress.setValue(80)
            
            # エクスポート実行
            success = self.export_handler.export_data(
                filtered_data, 
                shop_name, 
                self.start_date.date(),  # 開始日
                self.end_date.date(),    # 終了日
                export_type_lower
            )
            
            progress.setValue(100)
            progress.close()
            
            if success:
                QMessageBox.information(self, "エクスポート完了", f"{export_type} 形式でエクスポートが完了しました")
            else:
                QMessageBox.warning(self, "エクスポートエラー", "エクスポート中にエラーが発生しました")
                
        except Exception as e:
            progress.close()
            QMessageBox.warning(self, "エクスポートエラー", f"エクスポート中にエラーが発生しました:\n{str(e)}")

    def save_shop_name(self):
        """店舗名を設定ファイルに保存"""
        self.settings.setValue("shop_name", self.shop_input.text())
    
    def browse_folder(self):
        """CSVフォルダを選択"""
        folder = QFileDialog.getExistingDirectory(self, "CSVフォルダを選択")
        if folder:
            self.folder_path.setText(folder)
            # フォルダパスを設定ファイルに保存
            self.settings.setValue("last_folder_path", folder)
            self.last_folder_path = folder
        
    def set_today(self):
        """当日の日付を設定"""
        DateUtils.set_today(self.start_date, self.end_date)
    
    def set_this_month(self):
        """今月の日付範囲を設定"""
        DateUtils.set_this_month(self.start_date, self.end_date)

    def set_this_year(self):
        """年間の日付範囲を設定"""
        # 現在の年の1月1日を設定
        today = QDate.currentDate()
        start_of_year = QDate(today.year(), 1, 1)
        
        # 1月1日から当日までの範囲を設定
        self.start_date.setDate(start_of_year)
        self.end_date.setDate(today)

    def search_data(self):
        """日付範囲でデータを検索して表示"""
        folder_path = self.folder_path.text()
        if not folder_path:
            return
        
        # プログレスダイアログを作成
        progress = QProgressDialog("CSVデータを読み込み中...", "キャンセル", 0, 100, self)
        progress.setWindowTitle("データ読み込み")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)  # すぐに表示
        progress.setCancelButton(None)  # キャンセルボタンを無効化
        progress.setValue(0)  # この行を追加
        progress.show()

        self.repaint()
        QTimer.singleShot(100, lambda: self._perform_data_loading(progress))

    def _perform_data_loading(self, progress):
        try:
            progress.setLabelText("CSVファイルを読み込んでいます...")
            progress.setValue(20)
            self.repaint()

            # CSVデータ読み込み
            all_data = DataHandler.load_csv_data(self.folder_path.text())
            if all_data is None:
                progress.close()
                return
            
            progress.setLabelText("データをフィルタリング中...")
            progress.setValue(40)
            self.repaint()
            
            # CSVデータを保存
            self.csv_data = all_data
            
            # 日付範囲での絞り込み
            start_date_str = DataHandler.date_to_string(self.start_date.date())
            end_date_str = DataHandler.date_to_string(self.end_date.date())
            
            print(f"検索日付範囲: {start_date_str} から {end_date_str}")
            
            progress.setLabelText("表示データを準備中...")
            progress.setValue(60)
            self.repaint()
            
            # 日付列名を取得
            date_column_index = self.column_indices["date_column_index"]
            
            if len(all_data.columns) <= date_column_index:
                print(f"警告: 予想される取引日付の列が存在しません。列数: {len(all_data.columns)}")
                progress.close()
                return
            
            date_column_name = all_data.columns[date_column_index]
            print(f"取引日付の列: {date_column_name}")
            
            # 日付データをフィルタリング
            filtered_data = DataHandler.filter_data_by_date(all_data, start_date_str, end_date_str, date_column_name)
            
            progress.setLabelText("テーブルデータを準備中...")
            progress.setValue(80)
            self.repaint()
            
            if filtered_data.empty:
                progress.close()
                QMessageBox.information(self, "検索結果", "フィルター条件に合致するデータがありません")
                # テーブルをクリア
                self.product_table.setRowCount(0)
                self.group_table.setRowCount(0)
                self.receipt_detail_table.setRowCount(0)  # 追加
                self.total_count_label.setText("合計枚数: 0")
                self.total_amount_label.setText("合計金額: 0円")
                return
            
            # 集計データ作成
            summary_data = DataHandler.create_summary(filtered_data, self.column_indices)
            
            # 集計データを保存（詳細表示のために必要）
            self.last_summary = summary_data
            
            # 商品別テーブルに表示
            self.display_product_table(summary_data["product_summary"])
            
            # グループ別テーブルに表示
            self.display_group_table(summary_data["group_summary"])

            # 伝票別詳細テーブルに表示（修正）
            self.update_receipt_detail()
            
            progress.setValue(100)
            
            # 合計表示
            self.total_count_label.setText(f"合計枚数: {summary_data['total_count']}")
            self.total_amount_label.setText(f"合計金額: {summary_data['total_amount']:,}円")
            self.cashless_count_label.setText(f"【 キャッシュレス枚数: {summary_data['cashless_count']}")
            self.cashless_amount_label.setText(f"キャッシュレス金額: {summary_data['cashless_amount']:,}円 】")
            print(f"集計完了: 合計枚数={summary_data['total_count']}, 合計金額={summary_data['total_amount']}, " +
                f"キャッシュレス枚数={summary_data['cashless_count']}, キャッシュレス金額={summary_data['cashless_amount']}")
            
            progress.close()
            
        except Exception as e:
            progress.close()
            import traceback
            print(f"データ処理エラー: {e}")
            print(traceback.format_exc())
            QMessageBox.critical(self, "エラー", f"データ処理中にエラーが発生しました:\n{str(e)}")

        # TimeSeriesTabに日付範囲を設定
        if hasattr(self, 'time_series_tab'):
            self.time_series_tab.set_date_filter(start_date_str, end_date_str)
            
    def display_product_table(self, product_summary):
        """商品別テーブルにデータを表示"""
        self.product_table.setRowCount(0)
        
        for _, row in product_summary.iterrows():
            row_position = self.product_table.rowCount()
            self.product_table.insertRow(row_position)
            
            # 商品コード
            try:
                product_code = str(row.iloc[0])
                if product_code.isdigit():
                    self.product_table.setItem(row_position, 0, NumericTableWidgetItem(int(product_code)))
                else:
                    self.product_table.setItem(row_position, 0, QTableWidgetItem(product_code))
            except:
                self.product_table.setItem(row_position, 0, QTableWidgetItem(str(row.iloc[0])))
                
            # 商品名称
            self.product_table.setItem(row_position, 1, QTableWidgetItem(str(row.iloc[1])))
            
            # 枚数
            count_value = int(row.iloc[2])
            count_item = NumericTableWidgetItem(count_value, str(count_value))
            count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.product_table.setItem(row_position, 2, count_item)
            
            # 金額
            amount_value = int(row.iloc[3])
            amount_item = NumericTableWidgetItem(amount_value, f"{amount_value:,}")
            amount_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.product_table.setItem(row_position, 3, amount_item)
        
        # ソートがある場合は適用
        if self.sort_column is not None:
            self.product_table.sortItems(self.sort_column, self.sort_order)
            # ヘッダーにソート方向表示を更新
            self._update_header_sort_indicators(self.product_table, self.sort_column)
    
    def display_group_table(self, group_summary):
        """集計グループ別テーブルにデータを表示"""
        self.group_table.setRowCount(0)
        
        for _, row in group_summary.iterrows():
            row_position = self.group_table.rowCount()
            self.group_table.insertRow(row_position)
            
            # グループ番号
            try:
                group_num = str(row.iloc[0])
                if group_num.isdigit():
                    self.group_table.setItem(row_position, 0, NumericTableWidgetItem(int(group_num)))
                else:
                    self.group_table.setItem(row_position, 0, QTableWidgetItem(group_num))
            except:
                self.group_table.setItem(row_position, 0, QTableWidgetItem(str(row.iloc[0])))
                
            # グループ名称
            self.group_table.setItem(row_position, 1, QTableWidgetItem(str(row.iloc[1])))
            
            # 枚数
            count_value = int(row.iloc[2])
            count_item = NumericTableWidgetItem(count_value, str(count_value))
            count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.group_table.setItem(row_position, 2, count_item)
            
            # 金額
            amount_value = int(row.iloc[3])
            amount_item = NumericTableWidgetItem(amount_value, f"{amount_value:,}")
            amount_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.group_table.setItem(row_position, 3, amount_item)
        
        # ソートがある場合は適用
        if self.sort_column_g is not None:
            self.group_table.sortItems(self.sort_column_g, self.sort_order_g)
            # ヘッダーにソート方向表示を更新
            self._update_header_sort_indicators(self.group_table, self.sort_column_g, is_group=True)

    
    def update_receipt_detail(self):
        """選択された伝票種別の詳細を表示"""
        if self.csv_data is None or self.csv_data.empty:
            return
            
        self.receipt_detail_table.setRowCount(0)
        
        # 日付でフィルタリングされたデータを取得
        start_date_str = DataHandler.date_to_string(self.start_date.date())
        end_date_str = DataHandler.date_to_string(self.end_date.date())
        
        # 日付列名を取得
        date_column_index = self.column_indices["date_column_index"]
        date_column_name = self.csv_data.columns[date_column_index]
        
        # データ型を確認して強制的に文字列に変換
        date_data = self.csv_data[date_column_name].astype(str)
        
        # 日付フィルタリング
        filtered_data = self.csv_data[
            (date_data >= start_date_str) & 
            (date_data <= end_date_str)
        ]
        
        if filtered_data.empty:
            return
        
        # 列インデックスを取得
        amount_idx = self.column_indices["amount_idx"]
        count_idx = self.column_indices["count_idx"]
        amount_sign_idx = self.column_indices["amount_sign_idx"]
        card_deduction_idx = self.column_indices["card_deduction_idx"]
        product_code_idx = self.column_indices["product_code_idx"]
        product_name_idx = self.column_indices["product_name_idx"]
        
        # 列名を取得
        amount_col = filtered_data.columns[amount_idx]
        count_col = filtered_data.columns[count_idx]
        amount_sign_col = filtered_data.columns[amount_sign_idx]
        card_deduction_col = filtered_data.columns[card_deduction_idx]
        product_code_col = filtered_data.columns[product_code_idx]
        product_name_col = filtered_data.columns[product_name_idx]
        
        # データ型を数値に変換（エラーハンドリング付き）
        try:
            filtered_data[amount_col] = pd.to_numeric(filtered_data[amount_col], errors='coerce').fillna(0)
            filtered_data[count_col] = pd.to_numeric(filtered_data[count_col], errors='coerce').fillna(0)
            filtered_data[amount_sign_col] = pd.to_numeric(filtered_data[amount_sign_col], errors='coerce').fillna(0)
            filtered_data[card_deduction_col] = pd.to_numeric(filtered_data[card_deduction_col], errors='coerce').fillna(0)
        except Exception as e:
            print(f"データ変換エラー: {e}")
            return
        
        # 選択された伝票種別に応じてデータを絞り込み
        receipt_type = self.receipt_type_combo.currentText()
        
        if receipt_type == "現金売上":
            # 現金売上: 金額符号=0 且つ カード減算額=0
            target_data = filtered_data[
                (filtered_data[amount_sign_col] == 0) & 
                (filtered_data[card_deduction_col] == 0)
            ]
        elif receipt_type == "キャッシュレス決済":
            # キャッシュレス決済: 金額符号=0 且つ カード減算額≠0
            target_data = filtered_data[
                (filtered_data[amount_sign_col] == 0) & 
                (filtered_data[card_deduction_col] != 0)
            ]
        elif receipt_type == "赤伝処理":
            # 赤伝処理: 金額符号=1
            target_data = filtered_data[filtered_data[amount_sign_col] == 1]
        else:
            return
        
        if target_data.empty:
            return
        
        # 商品別に集計
        product_summary = target_data.groupby([product_code_col, product_name_col]).agg({
            count_col: 'sum',
            amount_col: 'sum'
        }).reset_index()
        
        # テーブルに表示
        for _, row in product_summary.iterrows():
            row_position = self.receipt_detail_table.rowCount()
            self.receipt_detail_table.insertRow(row_position)
            
            # 商品コード
            try:
                product_code = str(row.iloc[0])
                if product_code.isdigit():
                    self.receipt_detail_table.setItem(row_position, 0, NumericTableWidgetItem(int(product_code)))
                else:
                    self.receipt_detail_table.setItem(row_position, 0, QTableWidgetItem(product_code))
            except:
                self.receipt_detail_table.setItem(row_position, 0, QTableWidgetItem(str(row.iloc[0])))
                
            # 商品名称
            self.receipt_detail_table.setItem(row_position, 1, QTableWidgetItem(str(row.iloc[1])))
            
            # 枚数
            count_value = int(row.iloc[2])
            count_item = NumericTableWidgetItem(count_value, str(count_value))
            count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            
            # 赤伝処理の場合は赤文字で表示
            if receipt_type == "赤伝処理":
                count_item.setForeground(QColor(255, 0, 0))
                count_item.setText(f"({count_value})")
            
            self.receipt_detail_table.setItem(row_position, 2, count_item)
            
            # 金額
            amount_value = int(row.iloc[3])
            amount_item = NumericTableWidgetItem(amount_value, f"{amount_value:,}")
            amount_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            
            # 赤伝処理の場合は赤文字で表示
            if receipt_type == "赤伝処理":
                amount_item.setForeground(QColor(255, 0, 0))
                amount_item.setText(f"({amount_value:,})")
            
            self.receipt_detail_table.setItem(row_position, 3, amount_item)

        # 合計計算と表示更新（新規追加）
            total_count = 0
            total_amount = 0
            
            for row in range(self.receipt_detail_table.rowCount()):
                count_item = self.receipt_detail_table.item(row, 2)
                amount_item = self.receipt_detail_table.item(row, 3)
                
                if count_item and amount_item:
                    # NumericTableWidgetItemから数値を取得
                    if hasattr(count_item, 'numeric_value'):
                        count_value = count_item.numeric_value
                    else:
                        count_text = count_item.text().replace('(', '').replace(')', '').replace(',', '')
                        count_value = int(count_text) if count_text.isdigit() else 0
                    
                    if hasattr(amount_item, 'numeric_value'):
                        amount_value = amount_item.numeric_value
                    else:
                        amount_text = amount_item.text().replace('(', '').replace(')', '').replace(',', '')
                        amount_value = int(amount_text) if amount_text.replace('-', '').isdigit() else 0
                    
                    total_count += count_value
                    total_amount += amount_value
            
            # 合計表示を更新
            self.receipt_total_count_label.setText(f"合計枚数: {total_count}")
            self.receipt_total_amount_label.setText(f"合計金額: {total_amount:,}円")
            
            # ソートがある場合は適用（新規追加）
            if self.sort_column_r is not None:
                self.receipt_detail_table.sortItems(self.sort_column_r, self.sort_order_r)
                self._update_header_sort_indicators(self.receipt_detail_table, self.sort_column_r, is_receipt=True)              
    
    def sort_product_table(self, column_index):
        """商品別テーブルのソート処理"""
        if self.sort_column == column_index:
            # 同じカラムをクリックした場合、ソート順を逆転
            self.sort_order = Qt.DescendingOrder if self.sort_order == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            # 新しいカラムの場合、昇順でソート
            self.sort_column = column_index
            self.sort_order = Qt.AscendingOrder
        
        # ソート適用
        self.product_table.sortItems(column_index, self.sort_order)
        
        # ヘッダーにソート方向表示
        self._update_header_sort_indicators(self.product_table, column_index)
    
    def sort_group_table(self, column_index):
        """集計グループ別テーブルのソート処理"""
        if self.sort_column_g == column_index:
            # 同じカラムをクリックした場合、ソート順を逆転
            self.sort_order_g = Qt.DescendingOrder if self.sort_order_g == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            # 新しいカラムの場合、昇順でソート
            self.sort_column_g = column_index
            self.sort_order_g = Qt.AscendingOrder
        
        # ソート適用
        self.group_table.sortItems(column_index, self.sort_order_g)
        
        # ヘッダーにソート方向表示
        self._update_header_sort_indicators(self.group_table, column_index, is_group=True)
    
    def sort_receipt_table(self, column_index):
        """伝票別テーブルのソート処理"""
        if self.sort_column_r == column_index:
            # 同じカラムをクリックした場合、ソート順を逆転
            self.sort_order_r = Qt.DescendingOrder if self.sort_order_r == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            # 新しいカラムの場合、昇順でソート
            self.sort_column_r = column_index
            self.sort_order_r = Qt.AscendingOrder
        
        # ソート適用
        self.receipt_detail_table.sortItems(column_index, self.sort_order_r)
        
        # ヘッダーにソート方向表示
        self._update_header_sort_indicators(self.receipt_detail_table, column_index, is_receipt=True)

    def _update_header_sort_indicators(self, table, active_column, is_group=False, is_receipt=False):
        """テーブルヘッダーにソート方向を表示"""
        if is_receipt:
            sort_order = self.sort_order_r
        elif is_group:
            sort_order = self.sort_order_g
        else:
            sort_order = self.sort_order
        
        for i in range(table.columnCount()):
            header_text = table.horizontalHeaderItem(i).text().split(" ")[0]
            if i == active_column:
                direction = "▼" if sort_order == Qt.DescendingOrder else "▲"
                table.setHorizontalHeaderItem(i, QTableWidgetItem(f"{header_text} {direction}"))
            else:
                table.setHorizontalHeaderItem(i, QTableWidgetItem(header_text))