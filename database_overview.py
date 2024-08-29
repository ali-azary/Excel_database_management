import sqlite3
import pandas as pd

def database_overview(database_name):
    # Connect to the SQLite database
    connection = sqlite3.connect(database_name)
    cursor = connection.cursor()
    
    # Retrieve the list of all tables in the database
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()

    # Print overview for each table
    for table_name in tables:
        table_name = table_name[0]
        print(f"Table: {table_name}")
        
        # Get the count of rows in each table
        cursor.execute(f"SELECT COUNT(*) FROM {table_name};")
        count = cursor.fetchone()[0]
        print(f"Number of Records: {count}")
        
        # Get a preview of the first few rows of the table
        if count > 0:
            data_preview = pd.read_sql_query(f"SELECT * FROM {table_name} LIMIT 5;", connection)
            print("Data Preview:")
            print(data_preview)
        else:
            print("No data available.")
        
        print("\n")  # Add a newline for better readability between tables
    
    # Close the database connection
    connection.close()

# Specify the SQLite database name
database_name = "excel_database.db"

# Display the database overview
database_overview(database_name)
