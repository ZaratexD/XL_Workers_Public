import sqlite3 as sq
import os
import pandas as pd
import openpyxl as xl
from openpyxl.styles import NamedStyle, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

"""
Converts Excel File into database for easy SQL queries.
Used for v1-v6 later Enterprise Data Warehouse queried instead of Excel File
"""
def generate_ADE_DB(path_xl, path_db):
    try:
        # Read the Excel file, skipping the first row and setting headers to the second row
        # NOTE: Important to skip first row as Workday exports it with title there.
        df = pd.read_excel(path_xl, header=1)


        # Create a connection to the SQLite database
        conn = sq.connect(path_db)
        cursor = conn.cursor()

        # Keep the table name as 'ADE WD'
        table_name = 'ADE WD'
        
        # Convert the DataFrame to SQL, using the fixed table name
        df.to_sql(table_name, conn, if_exists='replace', index=False)

        # Commit changes and close the connection
        conn.commit()
        print(f"Data from {path_xl} has been successfully imported into {path_db} as table '{table_name}'.")

    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        if conn:
            conn.close()
