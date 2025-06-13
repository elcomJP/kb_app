# pdf_footer.py
from datetime import datetime
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

class PDFFooterCanvas:
    """
    PDFのフッターを描画するクラス
    """
    def __init__(self, total_pages=None):
        """
        初期化
        total_pages: PDFの総ページ数（指定した場合は「n/total」形式で表示）
        """
        self.total_pages = total_pages
    
    def __call__(self, canvas, doc):
        """
        キャンバスコールバック
        """
        canvas.saveState()
        
        # フォント設定
        try:
            font_name = 'HeiseiKakuGo-W5'  # 日本語フォント
            canvas.setFont(font_name, 8)
        except:
            # フォントが使用できない場合はデフォルトフォントを使用
            canvas.setFont('Helvetica', 8)
        
        # ページサイズを取得
        width = doc.pagesize[0]
        
        # 左下に出力日時を表示
        formatted_date = datetime.now().strftime('%Y年%m月%d日 %H:%M')
        canvas.drawString(15*mm, 10*mm, f"出力日時: {formatted_date}")
        
        # 右下にページ番号を表示
        page_num = canvas._pageNumber
        # シンプルなページ番号表示
        page_text = f"{page_num}ページ"
        
        canvas.drawRightString(width - 15*mm, 10*mm, page_text)
        
        canvas.restoreState()