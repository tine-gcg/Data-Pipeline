import pandas as pd
import os
import mysql.connector
from tkinter import Tk, filedialog
import shutil
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

USERNAME = os.getenv("SHAREPOINT_USERNAME", "big@globalcomfortgroup.com")
PASSWORD = os.getenv("SHAREPOINT_PASSWORD", "Fxckbvrn0ut!!!")
SITE_URL = "https://globalcomfortgroup0.sharepoint.com/sites/Database"
FOLDER_PATH = "Shared Documents/Data"

"""EXTRACT"""
# import excel file, allow user to select file
# save as sql file

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
# how the eff

def transform_to_sql(excel_file, table_name):
    """Transform Excel file into SQL insert statements."""
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file)
    
    sql_statements = []

    for i, row in df.iterrows():
        values = ', '.join([f"'{str(val)}'" for val in row.values])
        sql_statement = f"INSERT INTO {table_name} ({', '.join(df.columns)}) VALUES ({values});"
        sql_statements.append(sql_statement)

    # Save the SQL statements to a file
    sql_file = excel_file.replace(".xlsx", ".sql")
    with open(sql_file, 'w', encoding="utf-8") as file:
        file.write("\n".join(sql_statements))
    
    print(f"SQL file created: {sql_file}")
    return sql_file

# def transform_to_sql(excel_file, table_name, sheet_name=0, skiprows=0):
#     """Transform Excel file into SQL insert statements with flexible options."""
#     # Read the sheet, skipping unnecessary headers
#     df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=skiprows)

#     # Remove empty or unnamed columns
#     df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

#     # Drop fully empty rows (optional)
#     df = df.dropna(how='all')

#     # Generate SQL insert statements
#     sql_statements = []
#     for _, row in df.iterrows():
#         values = ', '.join([f"'{str(val).replace('\'', '\'\'')}'" for val in row.values])
#         sql = f"INSERT INTO {table_name} ({', '.join(df.columns)}) VALUES ({values});"
#         sql_statements.append(sql)

#     # Save to .sql
#     sql_file = excel_file.replace(".xlsx", ".sql")
#     with open(sql_file, 'w', encoding='utf-8') as f:
#         f.write('\n'.join(sql_statements))

#     print(f"SQL file created: {sql_file}")
#     return sql_file

"""LOAD"""
def file_exists_on_sharepoint(file_name, folder_path, ctx):
    folder = ctx.web.get_folder_by_server_relative_url(folder_path)
    files = folder.files
    ctx.load(files)
    ctx.execute_query()
    
    for f in files:
        if f.properties["Name"] == file_name:
            return True
    return False

def extract_table_name_from_sql(sql_file):
    with open(sql_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.upper().startswith("INSERT INTO"):
                parts = line.split()
                if len(parts) >= 3:
                    return parts[2]  # Assumes: INSERT INTO tablename (...)
    return None

# upload file to sharepoint
def load_to_sharepoint(sql_file, ctx):
    file_name = os.path.basename(sql_file)

    with open(sql_file, "rb") as content_file:
        target_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_PATH)
        target_file = target_folder.upload_file(file_name, content_file.read())
        ctx.execute_query()
        print(f"File '{file_name}' uploaded successfully to {FOLDER_PATH}")


def main():
    excel_file = extract_excel()
    if not excel_file:
        return

    file_name = os.path.basename(excel_file).replace(".xlsx", ".sql")

    
    ctx_auth = AuthenticationContext(SITE_URL)
    if not ctx_auth.acquire_token_for_user(USERNAME, PASSWORD):
        print("Authentication failed.")
        return

    ctx = ClientContext(SITE_URL, ctx_auth)

    # Check if the file already exists
    exists = file_exists_on_sharepoint(file_name, FOLDER_PATH, ctx)

    if not exists:
        table_name = input("Enter table name for SQL generation: ")
    else:
        print("File already exists in SharePoint. Using default table name.")
        # instead of using a default name like this, read the sql file and extract the table name
        table_name = "table"
        
    sql_file = transform_to_sql(excel_file, table_name)
    load_to_sharepoint(sql_file, ctx)


if __name__ == "__main__":
    main()