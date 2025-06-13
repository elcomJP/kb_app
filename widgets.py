from PyQt5.QtWidgets import QTableWidgetItem
from PyQt5.QtCore import Qt

class NumericTableWidgetItem(QTableWidgetItem):
    def __init__(self, value, formatted_text=None):
        # 表示用のテキストが指定されていれば、それを使用。なければ値をそのまま使用
        display_text = formatted_text if formatted_text is not None else str(value)
        super().__init__(display_text)
        
        # 数値データを保持（ソート用）
        self.value = float(value)
    
    def __lt__(self, other):
        # 数値比較でソートするためのメソッド
        if isinstance(other, NumericTableWidgetItem):
            return self.value < other.value
        return super().__lt__(other)