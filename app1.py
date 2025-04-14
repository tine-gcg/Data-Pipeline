import pandas as pd
import hashlib
import os
import re
import sqlite3
from tkinter import Tk, filedialog
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

USERNAME = os.getenv("SHAREPOINT_USERNAME", "big@globalcomfortgroup.com")
PASSWORD = os.getenv("SHAREPOINT_PASSWORD", "Fxckbvrn0ut!!!")
SITE_URL = "https://globalcomfortgroup0.sharepoint.com/sites/Database"
FOLDER_PATH = "Shared Documents/Data"

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
    return re.sub(r'\W+0', '_', name).strip('_').lower()

def generate_row_hash(row):
    """Generates a hash from row contents."""
    row_str = "|".join(map(str, row.values))
    return hashlib.md5(row_str.encode()).hexdigest()

def transform_and_load_to_sqlite(excel_file, sqlite_db_name="database.sqlite"):
    """Loads Excel sheets into SQLite DB with row_hash as a unique identifier."""
    dfs = pd.read_excel(excel_file, sheet_name=None)
    conn = sqlite3.connect(sqlite_db_name)

    # Get base filename (used in table name to prevent conflicts)
    base_filename = os.path.splitext(os.path.basename(excel_file))[0]
    base_filename_clean = base_filename.strip().replace(" ", "_").lower()

    for sheet_name, df in dfs.items():
        print(f"Processing sheet: {sheet_name}")
        
        # Clean column names
        df.columns = [clean_column_name(col) for col in df.columns]

        # Drop completely empty rows
        df.dropna(how='all', inplace=True)

        # Generate row_hash for each row
        df["row_hash"] = df.apply(generate_row_hash, axis=1)

        # Generate unique table name: file_sheet
        sheet_name_clean = sheet_name.strip().replace(" ", "_").lower()
        table_name = f"{base_filename_clean}_{sheet_name_clean}"

        # Create table with row_hash and insert data (replace = full overwrite)
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        print(f"Loaded sheet '{sheet_name}' into table '{table_name}' with row_hash")

    conn.close()
    print(f"All data loaded into '{sqlite_db_name}'")

# def transform_and_load_to_sqlite(excel_file, sqlite_db_name="database.sqlite"):
#     """Loads Excel sheets into a SQLite database."""
#     dfs = pd.read_excel(excel_file, sheet_name=None)
#     conn = sqlite3.connect(sqlite_db_name)
    
#     base_filename = os.path.splitext(os.path.basename(excel_file))[0]
#     base_filename_clean = re.sub(r'\W+', '_', base_filename.strip()).lower()

#     for sheet_name, df in dfs.items():
#         print(f"Processing sheet: {sheet_name}")

#         # Clean column names
#         df.columns = [clean_column_name(col) for col in df.columns]

#         # Drop completely empty rows
#         df.dropna(how='all', inplace=True)

#         # Create a unique table name using file name + sheet name
#         sheet_name_clean = re.sub(r'\W+', '_', sheet_name.strip()).lower()
#         table_name = f"{base_filename_clean}_{sheet_name_clean}"

#         # Load into SQLite
#         df.to_sql(table_name, conn, if_exists='replace', index=False)
#         print(f"Loaded sheet '{sheet_name}' into table '{table_name}'")

#     conn.close()
#     print(f"All data loaded into '{sqlite_db_name}'")

def upload_sqlite_to_sharepoint(site_url, folder_url, db_path, username, password):
    ctx_auth = AuthenticationContext(site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, ctx_auth)
        with open(db_path, 'rb') as content_file:
            file_content = content_file.read()
            target_folder = ctx.web.get_folder_by_server_relative_url(folder_url)
            name = os.path.basename(db_path)
            target_file = target_folder.upload_file(name, file_content).execute_query()
            print(f"Uploaded {name} to SharePoint.")
    else:
        print("Authentication failed.")


def main():
    excel_file = extract_excel()
    if not excel_file:
        return

    sqlite_path = "database.sqlite"
    transform_and_load_to_sqlite(excel_file, sqlite_db_name=sqlite_path)
    upload_sqlite_to_sharepoint(SITE_URL, FOLDER_PATH, sqlite_path, USERNAME, PASSWORD)


if __name__ == "__main__":
    main()
