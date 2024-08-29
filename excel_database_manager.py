import os
import sqlite3
import pandas as pd

class ExcelDatabaseManager:
    """
    A class to manage an SQLite database for storing information from Excel files.
    """

    def __init__(self, database_name="excel_database.db"):
        self.database_name = database_name
        self.connection = None
        self.cursor = None

    def connect(self):
        """Connects to the SQLite database."""
        self.connection = sqlite3.connect(self.database_name)
        self.cursor = self.connection.cursor()

    def disconnect(self):
        """Disconnects from the SQLite database."""
        if self.connection:
            self.connection.commit()
            self.connection.close()
            self.connection = None
            self.cursor = None

    def create_table(self, table_name="excel_data"):
        """Creates a table named 'table_name' within the database."""
        self.connect()
        query = f'''CREATE TABLE IF NOT EXISTS {table_name} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    path TEXT,
                    file_name TEXT,
                    data TEXT
                )'''
        self.cursor.execute(query)
        self.disconnect()

    def read_excel_file(self, file_path):
        """Reads data from an Excel file using pandas."""
        try:
            if file_path.endswith(".csv"):
                return pd.read_csv(file_path)
            elif file_path.endswith(".xlsx") or file_path.endswith(".xls"):
                return pd.read_excel(file_path, engine="openpyxl")
        except:
            pass

    def insert_data(self, table_name, file_path):
        """Inserts data from an Excel file into the specified table."""
        self.connect()
        file_data = self.read_excel_file(file_path)
        if file_data is not None:
            file_name = os.path.basename(file_path)
            file_data_str = file_data.to_string(index=False)
            query = f'''INSERT INTO {table_name} (path, file_name, data)
                        VALUES (?, ?, ?)'''
            self.cursor.execute(query, (file_path, file_name, file_data_str))
        else:
            print(f"Failed to read data from {file_path}")
        self.disconnect()


    def read_table_data(self, table_name):
        """Retrieves and returns all data from the specified table."""
        self.connect()
        query = f'SELECT * FROM {table_name}'
        self.cursor.execute(query)
        data = self.cursor.fetchall()
        self.disconnect()
        return data

    def search_keyword(self, table_name, keyword):
        self.connect()
        self.cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (table_name,))
        if not self.cursor.fetchone():
            self.disconnect()
            return f"Table '{table_name}' not found in the database."

        self.cursor.execute(f"SELECT path, file_name, data FROM {table_name} WHERE lower(data) LIKE ?;", ('%' + keyword.lower() + '%',))
        results = self.cursor.fetchall()
        self.disconnect()

        if results:
            result_texts = []
            for result in results:
                file_path, file_name, data = result
                relevant_row = next(row for row in data.split('\n') if keyword.lower() in row.lower())
                result_texts.append(f"File Path: {file_path}\nFile Name: {file_name}\nRelevant Data: {relevant_row}\n")
            return "\n".join(result_texts)
        else:
            return "No results found."

if __name__ == "__main__":
    # Example usage
    database_manager = ExcelDatabaseManager()
    database_manager.create_table()

    folder_path = "./"
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx") or file.endswith(".xls") or file.endswith(".csv"):
                file_path = os.path.join(root, file)
                database_manager.insert_data("excel_data", file_path)

    print("Process completed successfully.")
