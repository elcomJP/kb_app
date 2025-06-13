import os
from datetime import datetime
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PyQt5.QtCore import QDate

from excel_exporter import ExcelExporter
from pdf_exporter import PDFExporter

class ExportHandler:
    def __init__(self, parent=None):
        self.parent = parent
        self.excel_exporter = ExcelExporter()
        self.pdf_exporter = PDFExporter()
        # 最後に使用したエクスポート形式を記憶
        self.last_export_filter = "Excel ファイル (*.xlsx)"
    
    def _format_date(self, date):
        """QDateをYYYY/MM/DD形式の文字列に変換"""
        if isinstance(date, QDate):
            return date.toString('yyyy/MM/dd')
        return str(date)
        
    
    def export_data(self, data, shop_name, start_date, end_date, export_type=None):
        """
        データをエクスポートするメインメソッド
        形式を選んでファイルに出力する
        """
        # 開始日と終了日をフォーマット
        start_date_str = self._format_date(start_date)
        end_date_str = self._format_date(end_date)
        date_range_str = f"{start_date_str}～{end_date_str}"
        
        # ファイルシステムで使用できない文字を置換
        safe_shop_name = shop_name.replace('/', '_').replace('\\', '_').replace(':', '_')
        safe_start_date = start_date_str.replace('/', '')
        safe_end_date = end_date_str.replace('/', '')
        
        # 単日か複数日かに基づいてタイトルを設定
        if start_date == end_date:
            report_title = "売上日計表"
            file_date_part = safe_start_date
        else:
            report_title = "売上月計表"
            file_date_part = f"{safe_start_date}-{safe_end_date}"
        
        # export_typeに基づいてデフォルトのフィルターを設定
        if export_type is not None:
            if export_type.lower() == "excel":
                self.last_export_filter = "Excel ファイル (*.xlsx)"
            elif export_type.lower() == "pdf":
                self.last_export_filter = "PDF ファイル (*.pdf)"
            else:
                print(f"不明なエクスポート形式: {export_type}")
                if self.parent:
                    QMessageBox.warning(self.parent, "警告", f"不明なエクスポート形式: {export_type}")
                return False
        
        # エクスポート形式を選択するダイアログを表示
        export_filter = "Excel ファイル (*.xlsx);;PDF ファイル (*.pdf)"
        file_path, selected_filter = QFileDialog.getSaveFileName(
            self.parent,
            "エクスポート",
            f"{safe_shop_name}_{report_title}_{file_date_part}",
            export_filter,
            self.last_export_filter
        )
        
        if not file_path:
            return False  # キャンセルされた場合
        
        # 選択された形式を記憶
        self.last_export_filter = selected_filter
        
        # 選択された形式に応じてエクスポート処理を実行
        if selected_filter == "Excel ファイル (*.xlsx)":
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            return self.excel_exporter.export_to_excel(data, file_path, safe_shop_name, report_title, date_range_str, self.parent)
        elif selected_filter == "PDF ファイル (*.pdf)":
            if not file_path.endswith('.pdf'):
                file_path += '.pdf'
            return self.pdf_exporter.export_to_pdf(data, file_path, safe_shop_name, report_title, date_range_str, self.parent)
        
        return False