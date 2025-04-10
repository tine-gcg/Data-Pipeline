import pandas as pd
import os
import mysql.connector
from tkinter import Tk, filedialog
import shutil
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File


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

# def transform_to_sql(excel_file, table_name):
#     """Transform Excel file into SQL insert statements."""
#     # Read the Excel file into a DataFrame
#     df = pd.read_excel(excel_file)
    
#     sql_statements = []

#     for i, row in df.iterrows():
#         values = ', '.join([f"'{str(val)}'" for val in row.values])
#         sql_statement = f"INSERT INTO {table_name} ({', '.join(df.columns)}) VALUES ({values});"
#         sql_statements.append(sql_statement)

#     # Save the SQL statements to a file
#     sql_file = excel_file.replace(".xlsx", ".sql")
#     with open(sql_file, 'w', encoding="utf-8") as file:
#         file.write("\n".join(sql_statements))
    
#     print(f"SQL file created: {sql_file}")
#     return sql_file

def transform_to_sql(excel_file, table_name, sheet_name=0, skiprows=0):
    """Transform Excel file into SQL insert statements with flexible options."""
    # Read the sheet, skipping unnecessary headers
    df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=skiprows)

    # Remove empty or unnamed columns
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # Drop fully empty rows (optional)
    df = df.dropna(how='all')

    # Generate SQL insert statements
    sql_statements = []
    for _, row in df.iterrows():
        values = ', '.join([f"'{str(val).replace('\'', '\'\'')}'" for val in row.values])
        sql = f"INSERT INTO {table_name} ({', '.join(df.columns)}) VALUES ({values});"
        sql_statements.append(sql)

    # Save to .sql
    sql_file = excel_file.replace(".xlsx", ".sql")
    with open(sql_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(sql_statements))

    print(f"SQL file created: {sql_file}")
    return sql_file

"""LOAD"""
# upload file to sharepoint
# check if file exists, promt user to overwrite or not

def load_to_sharepoint(sql_file):
    username = os.getenv("SHAREPOINT_USERNAME", "big@globalcomfortgroup.com")
    password = os.getenv("SHAREPOINT_PASSWORD", "Fxckbvrn0ut!!!")
    
    site_url = "https://globalcomfortgroup0.sharepoint.com/sites/Database"                                                                                                                                                                                                                                                       
    folder_path = "Shared Documents/Data"  
    
    ctx_auth = AuthenticationContext(site_url)
    
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, ctx_auth)

        file_name = os.path.basename(sql_file)
        full_server_path = f"/sites/Database/{folder_path}/{file_name}"

        with open(sql_file, "rb") as content_file:
            File.save_binary(ctx, full_server_path, content_file)
            print(f"File '{file_name}' uploaded to SharePoint")
    else:
        print("Authentication failed. Check username/password or app access.")

def main():
    excel_file = extract_excel()
    if excel_file:
        sql_file = transform_to_sql(excel_file, table_name="guests") 
        load_to_sharepoint(sql_file)

if __name__ == "__main__":
    main()