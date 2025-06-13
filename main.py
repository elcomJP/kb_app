import sys
import streamlit as st
from PyQt5.QtWidgets import QApplication
from app import SalesAnalysisApp

if __name__ == "__main__":
    app = QApplication(sys.argv)

    app.setStyleSheet("""
        QMainWindow {
            background-color: #f5f5f5;
        }
        QGroupBox {
            font-weight: bold;
            border: 2px solid #cccccc;
            border-radius: 5px;
            margin-top: 1ex;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        QPushButton {
            background-color: #e0e0e0;
            border: 1px solid #cccccc;
            border-radius: 4px;
            padding: 5px 10px;
            min-width: 80px;
        }
        QPushButton:hover {
            background-color: #d0d0d0;
        }
        QPushButton:pressed {
            background-color: #c0c0c0;
        }
                      
    """)

    window = SalesAnalysisApp()
    window.show()
    sys.exit(app.exec_())
