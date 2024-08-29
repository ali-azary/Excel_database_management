import sqlite3

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