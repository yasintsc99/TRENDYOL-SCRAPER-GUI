import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QTextEdit, QProgressBar
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from trendyol_scraper import TrendyolScraper, style_excel
import pandas as pd

class ScraperThread(QThread):
    progress = pyqtSignal(str)
    percent = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path

    def run(self):
        data = pd.read_excel(self.excel_path)
        total = len(data)
        scraper = TrendyolScraper(self.excel_path, logger=self.progress.emit)
        scraper.driver.implicitly_wait(10)
        scraper.driver.get('https://www.trendyol.com/')
        try:
            scraper.safe_find_element("css selector","button#onetrust-reject-all-handler",timeout=2).click()
        except:
            pass
        for index, sellerText in enumerate(data["Mağaza Adı"], start=1):
            scraper.log(f"🔄 [{index}/{total}] {sellerText} mağazası işleniyor...")
            scraper.scrape_single(sellerText)
            percent_value = int((index / total) * 100)
            self.percent.emit(percent_value)
        scraper.driver.quit()
        scraper.sellerData.to_excel("Trendyol Satıcı Bilgileri (Detaylı).xlsx", index=False)
        style_excel("Trendyol Satıcı Bilgileri (Detaylı).xlsx")
        self.finished.emit()

class TrendyolGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🛍️ Trendyol Satıcı Bilgi Toplayıcı")
        self.setGeometry(300, 300, 550, 450)
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f4f8;
                font-family: 'Segoe UI';
            }
            QPushButton {
                background-color: #0078d7;
                color: white;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #005fa3;
            }
            QLabel {
                font-size: 16px;
                margin-bottom: 10px;
            }
            QTextEdit {
                background-color: #ffffff;
                border: 1px solid #ccc;
                padding: 6px;
                font-size: 13px;
            }
            QProgressBar {
                height: 20px;
                border: 1px solid #aaa;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #0078d7;
                width: 10px;
            }
        """)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.label = QLabel("Excel dosyasını seçin:")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

        self.select_button = QPushButton("📂 Dosya Seç")
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        self.start_button = QPushButton("🚀 Başlat")
        self.start_button.setEnabled(False)
        self.start_button.clicked.connect(self.start_scraping)
        layout.addWidget(self.start_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)

        self.setLayout(layout)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Seç", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excel_path = file_path
            self.label.setText(f"Seçilen dosya:\n{file_path}")
            self.start_button.setEnabled(True)

    def start_scraping(self):
        self.log_output.clear()
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.thread = ScraperThread(self.excel_path)
        self.thread.progress.connect(self.update_log)
        self.thread.percent.connect(self.update_progress)
        self.thread.finished.connect(self.scraping_finished)
        self.thread.start()

    def update_log(self, message):
        self.log_output.append(message)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def scraping_finished(self):
        self.progress_bar.setVisible(False)
        self.log_output.append("\n✅ İşlem tamamlandı. Dosya oluşturuldu: Trendyol Satıcı Bilgileri (Detaylı).xlsx")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TrendyolGUI()
    window.show()
    sys.exit(app.exec())