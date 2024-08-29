from PyQt5.QtCore import QThread, pyqtSignal
import os

class Worker(QThread):
    update_output = pyqtSignal(str)

    def __init__(self, folder_path, database_manager):
        super().__init__()
        self.folder_path = folder_path
        self.database_manager = database_manager

    def run(self):
        self.update_output.emit("\nStarting processing...")
        self.database_manager.create_table()
        for root, dirs, files in os.walk(self.folder_path):
            for file in files:
                if file.endswith(".xlsx") or file.endswith(".xls") or file.endswith(".csv"):
                    file_path = os.path.join(root, file)
                    self.update_output.emit(f"Processing file: {file_path}")
                    try:
                        self.database_manager.insert_data("excel_data", file_path)
                        self.update_output.emit(f"  - Inserted data successfully.")
                    except Exception as e:
                        self.update_output.emit(f"  - Error: {str(e)}")
        self.update_output.emit("\nProcess completed.")


import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QVBoxLayout, QTextEdit, QFileDialog, QLineEdit
from excel_database_manager import ExcelDatabaseManager

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Database Manager")
        self.folder_path = None
        self.database_manager = ExcelDatabaseManager()
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.create_layout()
        self.resize(800, 800)

    def create_layout(self):
        layout = QVBoxLayout()
        
        self.folder_button = QPushButton("Select Folder")
        self.folder_button.clicked.connect(self.select_folder)
        layout.addWidget(self.folder_button)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter keyword to search")
        layout.addWidget(self.search_input)
        
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_keyword)
        layout.addWidget(self.search_button)
        
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        layout.addWidget(self.output_text)
        
        self.central_widget.setLayout(layout)

    def select_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if self.folder_path:
            self.output_text.append("Selected folder: " + self.folder_path)
            self.process_files()

    def process_files(self):
        self.worker = Worker(self.folder_path, self.database_manager)
        self.worker.update_output.connect(self.output_text.append)
        self.worker.start()

    def search_keyword(self):
        keyword = self.search_input.text()
        if keyword:
            result = self.database_manager.search_keyword("excel_data", keyword)
            self.output_text.append(result)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
