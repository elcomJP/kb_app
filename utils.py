from PyQt5.QtCore import QDate

class DateUtils:
    @staticmethod
    def set_today(start_date_widget, end_date_widget):
        """当日の日付を設定"""
        today = QDate.currentDate()
        start_date_widget.setDate(today)
        end_date_widget.setDate(today)
    
    @staticmethod
    def set_this_month(start_date_widget, end_date_widget):
        """今月の日付範囲を設定"""
        today = QDate.currentDate()
        first_day = QDate(today.year(), today.month(), 1)
        start_date_widget.setDate(first_day)
        end_date_widget.setDate(today)