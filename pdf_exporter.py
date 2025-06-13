import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import QDate

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

from pdf_footer import PDFFooterCanvas

class PDFExporter:
    def __init__(self):
        self.jp_font_registered = self._register_japanese_fonts()
        
    def _register_japanese_fonts(self):
        """日本語フォントを登録する"""
        try:
            fonts_registered = False

            try:
                pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
                fonts_registered = True
                print("HeiseiKakuGo-W5 フォントを登録しました")
                return fonts_registered
            except Exception as e:
                print(f"HeiseiKakuGo-W5 フォント登録エラー: {e}")
                
            if os.name == 'nt':
                windows_fonts = [
                    ('MS Gothic', 'C:/Windows/Fonts/msgothic.ttc'),
                    ('Yu Gothic', 'C:/Windows/Fonts/YuGothR.ttc'),
                    ('Meiryo', 'C:/Windows/Fonts/meiryo.ttc')
                ]
                for font_name, font_path in windows_fonts:
                    try:
                        if os.path.exists(font_path):
                            pdfmetrics.registerFont(TTFont(font_name, font_path))
                            fonts_registered = True
                            print(f"{font_name} フォントを登録しました")
                            break
                    except Exception as font_e:
                        print(f"{font_name} フォント登録エラー: {font_e}")
            elif os.name == 'posix':
                font_paths = [
                    ('/usr/share/fonts/truetype/ipafont-gothic/ipag.ttf', 'IPAGothic'),
                    ('/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf', 'IPAGothic'),
                    ('/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc', 'Hiragino'),
                    ('/Library/Fonts/ヒラギノ角ゴシック W3.ttc', 'Hiragino')
                ]
                for path, name in font_paths:
                    if os.path.exists(path):
                        try:
                            pdfmetrics.registerFont(TTFont(name, path))
                            fonts_registered = True
                            print(f"{name} フォントを登録しました")
                            break
                        except Exception as font_e:
                            print(f"{name} フォント登録エラー: {font_e}")
            
            return fonts_registered
        except Exception as e:
            print(f"フォント登録処理全体エラー: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _format_date(self, date):
        """QDateをYYYY/MM/DD形式の文字列に変換"""
        if isinstance(date, QDate):
            return date.toString('yyyy/MM/dd')
        return str(date)

    def _get_report_title(self, date_str):
        """日付文字列から適切なレポートタイトルを決定する"""
        if "～" in date_str:
            start_date, end_date = date_str.split("～")
            if start_date.strip() != end_date.strip():
                return "売上月計表"
        return "売上日計表"

    def _format_menu_number(self, menu_num):
        """メニュー番号から先頭の0を削除する"""
        try:
            menu_str = str(menu_num)
            return str(int(menu_str)) if menu_str.isdigit() else menu_str
        except (ValueError, TypeError):
            return str(menu_num)
    
    def export_to_pdf(self, data, file_path, shop_name, title, date_str, parent=None):
        """データをPDFファイルにエクスポート（日本語対応版）"""
        try:
            if data is None or len(data) == 0:
                print("エクスポートするデータがありません")
                if parent:
                    QMessageBox.warning(parent, "警告", "エクスポートするデータがありません")
                return False
            
            if isinstance(title, QDate):
                title = self._format_date(title)
            
            try:
                data_copy = data.copy()
            except:
                try:
                    data_copy = pd.DataFrame(data)
                except:
                    print("データをDataFrameに変換できません")
                    if parent:
                        QMessageBox.warning(parent, "エラー", "データ形式が不正です")
                    return False
            
            column_mapping = {
                'T': '集計Ｇ名称',
                'Q': '商品コード',
                'R': '論理口座名称',
                'J': '枚数',
                'L': '金額',
                'K': '金額符号',
                'M': 'カード減算額'
            }
            
            for col_letter, col_name in column_mapping.items():
                if col_name not in data_copy.columns:
                    print(f"警告: 列 '{col_name}' がデータに存在しません")
                    
            doc = SimpleDocTemplate(
                file_path,
                pagesize=landscape(A4),
                rightMargin=10*mm,
                leftMargin=10*mm,
                topMargin=10*mm,
                bottomMargin=15*mm
            )
            styles = getSampleStyleSheet()
            elements = []
            
            jp_font_name = 'HeiseiKakuGo-W5' if self.jp_font_registered else 'Helvetica'
            
            title_style = ParagraphStyle(
                'TitleJP',
                parent=styles['Title'],
                fontName=jp_font_name,
                fontSize=18,
                alignment=0,
                spaceAfter=5,
                encoding='utf-8'
            )
            
            date_style = ParagraphStyle(
                'DateJP',
                parent=styles['Normal'],
                fontName=jp_font_name,
                fontSize=14,
                alignment=0,
                spaceAfter=10,
                encoding='utf-8'
            )
            
            normal_style = ParagraphStyle(
                'NormalJP',
                parent=styles['Normal'],
                fontName=jp_font_name,
                fontSize=10,
                encoding='utf-8'
            )
            
            pdf_title = self._get_report_title(date_str)
            
            elements.append(Paragraph(f"{pdf_title} 店舗名: {shop_name}", title_style))
            elements.append(Paragraph(f"集計日: {date_str}", date_style))
            elements.append(Spacer(1, 10*mm))
            
            table_header = ['グループ名', 'メニュー番号', 'メニュー名', '数量', '金額']
            
            amount_sign_col = '金額符号' if '金額符号' in data_copy.columns else 'K'
            card_deduction_col = 'カード減算額' if 'カード減算額' in data_copy.columns else 'M'

            normal_data = data_copy[
                (data_copy[amount_sign_col] == '0') & 
                (data_copy[card_deduction_col] == '0')
            ].copy() if amount_sign_col in data_copy.columns and card_deduction_col in data_copy.columns else pd.DataFrame(columns=data_copy.columns)

            red_data = data_copy[
                (data_copy[amount_sign_col] == '1')
            ].copy() if amount_sign_col in data_copy.columns else pd.DataFrame(columns=data_copy.columns)

            cashless_data = data_copy[
                (data_copy[amount_sign_col] == '0') & 
                (data_copy[card_deduction_col] != '0')
            ].copy() if amount_sign_col in data_copy.columns and card_deduction_col in data_copy.columns else pd.DataFrame(columns=data_copy.columns)
            
            from reportlab.platypus import CondPageBreak
            
            elements.append(Paragraph("【現金売上】", normal_style))
            elements.append(Spacer(1, 5*mm))
            normal_table_data = [table_header]
            normal_total = {'数量': 0, '金額': 0}
            self._add_group_data_to_table(normal_data, normal_table_data, normal_total)
            
            if len(normal_table_data) == 1:
                normal_table_data.append(['データなし', '', '', '', ''])
            
            normal_table_data.append([
                '現金売上 計', '', '',
                f"{normal_total['数量']:,}", f"{normal_total['金額']:,}"
            ])
            
            normal_table = self._create_table(normal_table_data)
            elements.append(normal_table)
            elements.append(Spacer(1, 10*mm))
            
            elements.append(PageBreak()) 
            elements.append(Paragraph("【キャッシュレス決済】", normal_style))
            elements.append(Spacer(1, 5*mm))
            cashless_table_data = [table_header]
            cashless_total = {'数量': 0, '金額': 0}
            self._add_group_data_to_table(cashless_data, cashless_table_data, cashless_total)
            
            if len(cashless_table_data) == 1:
                cashless_table_data.append(['データなし', '', '', '', ''])
            
            cashless_table_data.append([
                'キャッシュレス決済 計', '', '',
                f"{cashless_total['数量']:,}", f"{cashless_total['金額']:,}"
            ])
            
            cashless_table = self._create_table(cashless_table_data)
            elements.append(cashless_table)
            elements.append(Spacer(1, 10*mm))
            
            elements.append(PageBreak()) 
            elements.append(Paragraph("【赤伝】", normal_style))
            elements.append(Spacer(1, 5*mm))
            red_table_data = [table_header]
            red_total = {'数量': 0, '金額': 0}
            self._add_group_data_to_table(red_data, red_table_data, red_total, is_red_slip=True)
            
            if len(red_table_data) == 1:
                red_table_data.append(['データなし', '', '', '', ''])
            
            red_table_data.append([
                '赤伝 計', '', '',
                f"{red_total['数量']:,}", f"{red_total['金額']:,}"
            ])
            
            red_table = self._create_table(red_table_data)
            elements.append(red_table)
            elements.append(Spacer(1, 10*mm))
            
            elements.append(Paragraph("【総計】", normal_style))
            elements.append(Spacer(1, 5*mm))
            total_table_data = [
                table_header,
                [
                    '総計(現金･キャッシュレス)', '', '',
                    f"{normal_total['数量'] + cashless_total['数量']:,}",
                    f"{normal_total['金額'] + cashless_total['金額']:,}"
                ]
            ]
            
            total_table = self._create_table(total_table_data)
            elements.append(total_table)
            
            footer = PDFFooterCanvas()
            doc.build(elements, onFirstPage=footer, onLaterPages=footer)
            
            return True
                
        except Exception as e:
            print(f"PDF出力エラー: {e}")
            import traceback
            traceback.print_exc()
            if parent:
                QMessageBox.critical(parent, "エラー", f"PDF出力に失敗しました:\n{str(e)}")
            return False

    def _add_group_data_to_table(self, data, table_data, total_accumulator, is_red_slip=False):
        """メニュー名ごとにデータを集計してテーブルに追加する"""
        if data.empty:
            return
        
        group_name_column = '集計Ｇ名称' if '集計Ｇ名称' in data.columns else 'T'
        menu_number_column = '商品コード' if '商品コード' in data.columns else 'Q'
        menu_name_column = '論理口座名称' if '論理口座名称' in data.columns else 'R'
        quantity_column = '枚数' if '枚数' in data.columns else 'J'
        amount_column = '金額' if '金額' in data.columns else 'L'
        
        menu_summary = {}
        
        for _, row in data.iterrows():
            group_name = str(row.get(group_name_column, "")).strip()
            if not group_name:
                group_name = "その他"
                
            menu_num = row.get(menu_number_column, "")
            menu_name = row.get(menu_name_column, "")
            
            menu_key = (group_name, menu_num, menu_name)
            
            try:
                quantity = int(row.get(quantity_column, 0))
            except (ValueError, TypeError):
                quantity = 0
                
            try:
                amount = int(row.get(amount_column, 0))
            except (ValueError, TypeError):
                amount = 0
            
            if is_red_slip:
                quantity = -quantity
                amount = -amount
            
            if menu_key in menu_summary:
                menu_summary[menu_key]['quantity'] += quantity
                menu_summary[menu_key]['amount'] += amount
            else:
                menu_summary[menu_key] = {
                    'quantity': quantity,
                    'amount': amount
                }
        
        current_group = None
        group_subtotal = {'数量': 0, '金額': 0}
        
        sorted_keys = sorted(menu_summary.keys(), key=lambda x: (x[0], x[1]))
        
        for key in sorted_keys:
            group_name, menu_num, menu_name = key
            summary = menu_summary[key]
            
            if current_group != group_name:
                if current_group is not None:
                    table_data.append([
                        f"{current_group} 計", '', '',
                        f"{group_subtotal['数量']:,}", f"{group_subtotal['金額']:,}"
                    ])
                    group_subtotal = {'数量': 0, '金額': 0}
                
                current_group = group_name
                table_data.append([group_name, '', '', '', ''])
            
            quantity = summary['quantity']
            amount = summary['amount']
            
            table_data.append([
                '', self._format_menu_number(menu_num), menu_name,
                f"{quantity:,}", f"{amount:,}"
            ])
            
            group_subtotal['数量'] += quantity
            group_subtotal['金額'] += amount
            total_accumulator['数量'] += quantity
            total_accumulator['金額'] += amount
        
        if current_group:
            table_data.append([
                f"{current_group} 計", '', '',
                f"{group_subtotal['数量']:,}", f"{group_subtotal['金額']:,}"
            ])
    
    def _create_table(self, table_data):
        """スタイル付きのテーブルを作成"""
        available_width = 257*mm
        
        col_widths = [
            available_width * 0.20,
            available_width * 0.10,
            available_width * 0.40,
            available_width * 0.15,
            available_width * 0.15
        ]
        
        row_heights = [16]
        for i in range(1, len(table_data)):
            row_heights.append(12)
        
        table = Table(table_data, colWidths=col_widths, rowHeights=row_heights, repeatRows=1)
        
        jp_font_name = 'HeiseiKakuGo-W5' if self.jp_font_registered else 'Helvetica'
        
        table_style = TableStyle([           
            ('LINEABOVE', (1,2), (-1,-2), 0.3, colors.black),

            ('TOPPADDING', (0,0), (-1,0), -5),
            ('BOTTOMPADDING', (0,0), (-1,0), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 5),
            ('RIGHTPADDING', (0,0), (-1,-1), 5),
            
            ('BACKGROUND', (0,0), (-1,0), colors.white),
            ('FONT', (0,0), (-1,-1), jp_font_name, 9),
            ('FONT', (0,0), (-1,0), jp_font_name, 10, True),
            ('ALIGNMENT', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),

            ('LINEABOVE', (0,0), (-1,0), 0.8, colors.black),
            ('LINEBELOW', (0,0), (-1,0), 0.8, colors.black),

            ('ALIGNMENT', (3,1), (4,-1), 'RIGHT'),
        ])
        
        for i, row in enumerate(table_data):
            if i > 0:
                if len(str(row[0])) > 0 and all(str(cell).strip() == '' for cell in row[1:]):
                    table_style.add('FONT', (0,i), (0,i), jp_font_name, 10, True)
                    table_style.add('ALIGNMENT', (0,i), (0,i), 'LEFT')
                    table_style.add('VALIGN', (0,i), (0,i), 'MIDDLE')  
                    table_style.add('TOPPADDING', (0,i), (0,i), -8)
                    table_style.add('BOTTOMPADDING', (0,i), (0,i), 3)

                elif '計' in str(row[0]):
                    if '総計' in str(row[0]) or '現金売上 計' in str(row[0]) or '赤伝 計' in str(row[0]) or 'キャッシュレス決済 計' in str(row[0]):
                        table_style.add('BACKGROUND', (0,i), (-1,i), colors.lightgrey)
                        table_style.add('FONT', (0,i), (-1,i), jp_font_name, 9, True)
                        table_style.add('VALIGN', (0,i), (-1,i), 'MIDDLE')
                        table_style.add('TOPPADDING', (0,i), (-1,i), -8)
                        table_style.add('BOTTOMPADDING', (0,i), (-1,i), 3)
                    else:
                        table_style.add('FONT', (0,i), (-1,i), jp_font_name, 9, True)
                        table_style.add('VALIGN', (0,i), (-1,i), 'MIDDLE')
                        table_style.add('TOPPADDING', (0,i), (-1,i), -8)
                        table_style.add('BOTTOMPADDING', (0,i), (-1,i), 3)
                        table_style.add('LINEBELOW', (0,i), (-1,i), 0.8, colors.black)

        table.setStyle(table_style)
        return table