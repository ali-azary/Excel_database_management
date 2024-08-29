# Excel_database_management
A proof of concept or MVP of unifying and integrating all excel data files into a single SQL database, that makes it easy to look for a specific keyword and find the excel files that contains it. It can be expanded to add more functionalitie for data analysis and reporting.

Streamlining Data Management: Excel Files to SQLite Database
Oftentimes you have your data in several Excel files scattered across various folders. Extracting insights from these files can be time-consuming and prone to errors, especially when dealing with large number of files. Remembering what is stored where and also finding data from different files will become impossible when you have a large number of files and folders.As a simple solution, with a Python script, you can seamlessly transition your Excel files into a centralized SQL database, and also you can easily search for any data and find the values and the file and path where it is stored. 

Database Creation
The first step involves creating a SQLite database. This database serves as a unified repository for all your Excel data.
import sqlite3

def create_database(database_name):
    connection = sqlite3.connect(database_name)
    connection.close()

# Specify the SQLite database name
database_name = "excel_database.db"

# Create the SQLite database
create_database(database_name)

Table Structure Definition
Within the database, define a table structure that mirrors the columns in your Excel files. This ensures consistency and uniformity in data storage.
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

# Database and table names
database_name = "excel_database.db"
table_name = "excel_data"

# Create the SQLite table
create_table(database_name, table_name)


Data Import and Search 
Read Excel files from a specified folder, insert their data into the SQLite table, and enable users to search for specific keywords within the data.
import os
import sqlite3

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

def search_keyword_in_database(database_name, table_name, keyword):
    connection = sqlite3.connect(database_name)
    cursor = connection.cursor()

    # Check if the specified table exists
    cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (table_name,))
    if not cursor.fetchone():
        print(f"Table '{table_name}' not found in the database.")
        connection.close()
        return

    print(f"Searching for '{keyword}' in table: {table_name}")

    # Perform a SELECT query for the keyword (case-insensitive)
    cursor.execute(f"SELECT path, file_name, data FROM {table_name} WHERE lower(data) LIKE ?;", ('%' + keyword.lower() + '%',))
    results = cursor.fetchall()

    # Print the results
    if results:
        print("Results found:")
        for result in results:
            file_path, file_name, data = result
            # Find the specific row containing the keyword
            relevant_row = next(row for row in data.split('\n') if keyword.lower() in row.lower())
            print(f"File Path: {file_path}\nFile Name: {file_name}\nRelevant Data: {relevant_row}\n")
    else:
        print("No results found.")

    connection.close()

# Specify the SQLite database and the keyword to search for
database_name = "excel_database.db"
table_name = 'excel_data'
keyword_to_search = input('Enter keyword to search: ')

# Search for the keyword in the database
search_keyword_in_database(database_name, table_name, keyword_to_search)

![2024-08-29 15 04 59](https://github.com/user-attachments/assets/fbbcdc0e-6f97-4afe-b0a7-24d9b1f47cd5)

Conclusion
By transitioning from Excel files to a SQLite database, you streamline your data management process, enhance accessibility, and pave the way for more sophisticated data analysis. 
In conclusion, embracing the power of Python and SQL for managing Excel files not only simplifies data management but also unlocks the full potential of your data assets. With the right tools and methods in place, you can harness the power of centralized data storage, enabling informed decision-making and driving business growth.
