import sys
import os
import sqlite3
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit, QVBoxLayout, QWidget, QLineEdit, QLabel, QHBoxLayout

class ExcelFileFinderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Data Manager V 1.0.0")
        self.resize(800, 600)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        self.text_edit = QTextEdit(self)
        layout.addWidget(self.text_edit)

        self.button = QPushButton("Select Root Folder", self)
        layout.addWidget(self.button)
        self.button.clicked.connect(self.select_root_folder)

        search_layout = QHBoxLayout()
        self.search_input = QLineEdit(self)
        search_button = QPushButton("Search Keyword", self)
        search_button.clicked.connect(self.search_keyword)

        search_layout.addWidget(QLabel("Search Keyword:", self))
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(search_button)
        layout.addLayout(search_layout)

        # Initialize the SQLite database and create the table
        self.database_name = "excel_database.db"
        self.table_name = "excel_data"
        self.create_database()
        self.create_table()

    def create_database(self):
        connection = sqlite3.connect(self.database_name)
        connection.close()

    def create_table(self):
        connection = sqlite3.connect(self.database_name)
        cursor = connection.cursor()
        cursor.execute(f'''CREATE TABLE IF NOT EXISTS {self.table_name} (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            path TEXT,
                            file_name TEXT,
                            data TEXT
                        )''')
        connection.commit()
        connection.close()

    def select_root_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Root Folder")
        if folder_path:
            excel_files = self.find_excel_files(folder_path)
            if excel_files:
                self.text_edit.setPlainText("\n".join(excel_files))
                self.read_excel_files_and_insert(folder_path)
            else:
                self.text_edit.setPlainText("No Excel files found in the specified folder and subfolders.")

    def find_excel_files(self, root_folder):
        excel_files = []
        for root, dirs, files in os.walk(root_folder):
            for file in files:
                if file.lower().endswith((".xlsx", ".xls", ".csv")):
                    file_path = os.path.join(root, file)
                    excel_files.append(file_path)
        return excel_files

    def read_excel_files_and_insert(self, folder_path):
        connection = sqlite3.connect(self.database_name)
        cursor = connection.cursor()
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith((".xlsx", ".xls", ".csv")):
                    file_path = os.path.join(root, file)
                    cursor.execute(f'SELECT COUNT(*) FROM {self.table_name} WHERE path = ?', (file_path,))
                    count = cursor.fetchone()[0]
                    if count == 0:
                        try:
                            if file.endswith(".csv"):
                                file_data = pd.read_csv(file_path)
                            elif file.endswith(".xlsx") or file.endswith(".xls"):
                                file_data = pd.read_excel(file_path, engine='openpyxl')
                            cursor.execute(f'''INSERT INTO {self.table_name} (path, file_name, data) VALUES (?, ?, ?)''', (file_path, file, file_data.to_csv(index=False) if file.endswith(".csv") else file_data.to_string(index=False)))
                        except:
                           pass
        connection.commit()
        connection.close()
        self.text_edit.append("Excel files processed and data inserted into the database.")

    def search_keyword(self):
        keyword = self.search_input.text().strip()
        if keyword:
            self.search_keyword_in_database(keyword)

    def search_keyword_in_database(self, keyword):
        connection = sqlite3.connect(self.database_name)
        cursor = connection.cursor()
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (self.table_name,))
        if not cursor.fetchone():
            self.text_edit.setPlainText(f"Table '{self.table_name}' not found in the database.")
            connection.close()
            return

        cursor.execute(f"SELECT path, file_name, data FROM {self.table_name} WHERE lower(data) LIKE ?;", ('%' + keyword.lower() + '%',))
        results = cursor.fetchall()
        if results:
            result_texts = []
            for result in results:
                file_path, file_name, data = result
                relevant_row = next(row for row in data.split('\n') if keyword.lower() in row.lower())
                result_texts.append(f"File Path: {file_path}\nFile Name: {file_name}\nRelevant Data: {relevant_row}\n")
            self.text_edit.setPlainText("\n".join(result_texts))
        else:
            self.text_edit.setPlainText("No results found.")
        connection.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelFileFinderApp()
    window.show()
    sys.exit(app.exec_())
