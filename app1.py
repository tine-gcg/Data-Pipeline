import pandas as pd
import os
import sqlite3
from tkinter import Tk, filedialog

"""EXTRACT"""
# import excel file, allow user to select file
def extract_excel():
    """Allows user to select and load an Excel file."""
    Tk().withdraw()  
    excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if excel_file:
        print(f"File selected: {excel_file}")
    else:
        print("No file selected.")
    return excel_file

"""TRANSFORM"""
def clean_column_name(name):
    """Clean column names to be SQLite-safe."""
    return name.strip().replace(" ", "_").replace("-", "_").replace(".", "").lower()

def transform_and_load_to_sqlite(excel_file, sqlite_db_name="database.sqlite"):
    """Loads Excel sheets into a SQLite database."""
    dfs = pd.read_excel(excel_file, sheet_name=None)
    conn = sqlite3.connect(sqlite_db_name)

    for sheet_name, df in dfs.items():
        print(f"Processing sheet: {sheet_name}")

        # Clean column names
        df.columns = [clean_column_name(col) for col in df.columns]

        # Drop completely empty rows
        df.dropna(how='all', inplace=True)

        # Create a clean table name
        table_name = sheet_name.strip().replace(" ", "_").lower()

        # Load into SQLite
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        print(f"Loaded sheet '{sheet_name}' into table '{table_name}'")

    conn.close()
    print(f"All data loaded into '{sqlite_db_name}'")

"""MAIN"""
def main():
    excel_file = extract_excel()
    if not excel_file:
        return

    transform_and_load_to_sqlite(excel_file)

if __name__ == "__main__":
    main()
