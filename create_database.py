import os
import sqlite3
import pandas as pd

def create_database(database_name):
    connection = sqlite3.connect(database_name)
    connection.close()

def create_table(database_name, table_name):
    connection = sqlite3.connect(database_name)
    cursor = connection.cursor()

    # Create a table with columns 'path', 'file_name', and 'data'
    cursor.execute(f'''CREATE TABLE IF NOT EXISTS {table_name} (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        path TEXT,
                        file_name TEXT,
                        data TEXT
                    )''')

    connection.commit()
    connection.close()

def read_excel_files_and_insert(database_name, table_name, folder_path):
    connection = sqlite3.connect(database_name)
    cursor = connection.cursor()

    # Iterate through all files in the specified folder and its subfolders
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx") or file.endswith(".xls") or file.endswith('.csv'):
                file_path = os.path.join(root, file)

                # Check if the record already exists in the table
                cursor.execute(f'SELECT COUNT(*) FROM {table_name} WHERE path = ?', (file_path,))
                count = cursor.fetchone()[0]

                if count == 0:
                    # Read file using pandas
                    if file.endswith(".csv"):
                        file_data = pd.read_csv(file_path)
                    elif file.endswith(".xlsx") or file.endswith(".xls"):
                        file_data = pd.read_excel(file_path, engine='openpyxl')
                    else:
                        continue  # Skip files with unsupported extensions

                    # Insert data into the SQLite table
                    cursor.execute(f'''INSERT INTO {table_name} (path, file_name, data)
                                       VALUES (?, ?, ?)''', (file_path, file, str(file_data)))


    connection.commit()
    connection.close()

# Database and table names
database_name = "excel_database.db"
table_name = "excel_data"

# Folder path containing Excel files
folder_path = "./"

# Create the SQLite database and table
create_database(database_name)
create_table(database_name, table_name)

# Read Excel files and insert data into the table
read_excel_files_and_insert(database_name, table_name, folder_path)

print("Process completed successfully.")


def print_table_data(database_name, table_name):
    connection = sqlite3.connect(database_name)
    cursor = connection.cursor()

    # Select all data from the table
    cursor.execute(f'SELECT * FROM {table_name}')
    rows = cursor.fetchall()

    # Print the table data
    print(f"Table: {table_name}")
    print("ID  | Path                         | File Name         | Data")
    print("-" * 70)

    for row in rows:
        print(f"{row[0]:<4} | {row[1]:<30} | {row[2]:<17} | {row[3]}")

    connection.close()

# Database and table names
database_name = "excel_database.db"
table_name = "excel_data"

# Print the contents of the table
print_table_data(database_name, table_name)