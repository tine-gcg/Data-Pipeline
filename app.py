import base64
import pandas as pd
import hashlib
import os
import re
import sqlite3
import streamlit as st
import tempfile
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from streamlit_option_menu import option_menu

USERNAME = os.getenv("SHAREPOINT_USERNAME", "big@globalcomfortgroup.com")
PASSWORD = os.getenv("SHAREPOINT_PASSWORD", "Fxckbvrn0ut!!!")
SITE_URL = "https://globalcomfortgroup0.sharepoint.com/sites/BIGDatabase"
FOLDER_PATH = "Shared Documents"

# """EXTRACT"""
def upload_excel():
    """Allows user to select and load an Excel file."""
    excel_file = st.file_uploader("", type=["xlsx", "xls"])
    return excel_file

# """TRANSFORM"""
# def clean_column_name(name):
#     return re.sub(r'\W+0', '_', name).strip('_').lower()

def generate_row_hash(row):
    """Generates a hash from row contents."""
    row_str = "|".join(map(str, row.values))
    return hashlib.md5(row_str.encode()).hexdigest()

# def transform_and_load_to_sqlite(excel_file, sqlite_db_name="database.sqlite"):
#     """Loads Excel sheets into SQLite DB with row_hash as a unique identifier."""
#     dfs = pd.read_excel(excel_file, sheet_name=None)
#     conn = sqlite3.connect(sqlite_db_name)

#     # Get base filename (used in table name to prevent conflicts)
#     base_filename = os.path.splitext(os.path.basename(excel_file.name))[0]
#     base_filename_clean = base_filename.strip().replace(" ", "_").lower()

#     # Prepend an underscore to base_filename if it starts with a digit
#     if base_filename_clean[0].isdigit():  # If base_filename starts with a digit
#         base_filename_clean = f"_{base_filename_clean}"

#     preview_data = {}

#     for sheet_name, df in dfs.items():
#         # Clean the sheet_name for use in the table name
#         sheet_name_clean = sheet_name.strip().replace(" ", "_").lower()
        
#         # Use the base_filename_clean (which has been modified if necessary) and append the cleaned sheet name
#         table_name = f"{base_filename_clean}_{sheet_name_clean}"

#         # Clean column names
#         df.columns = [clean_column_name(col) for col in df.columns]
#         df.dropna(how='all', inplace=True)
#         df["row_hash"] = df.apply(generate_row_hash, axis=1)

#         # Create table with row_hash and insert data (replace = full overwrite)
#         df.to_sql(table_name, conn, if_exists='replace', index=False)

#         preview_data[sheet_name] = df.head()

#     conn.commit()
#     conn.close()
#     return preview_data

def clean_name(name):
    # Replace any non-word character (anything not a letter, number, or underscore) with underscores
    name = re.sub(r'\W+', '_', name)
    # Remove leading underscores just to be clean
    name = name.lstrip('_')
    # If after cleaning it starts with a digit, add an underscore in front
    if name and name[0].isdigit():
        name = '_' + name
    return name.lower()

def transform_and_load_to_sqlite(excel_file, sqlite_db_name="database.sqlite"):
    dfs = pd.read_excel(excel_file, sheet_name=None)
    conn = sqlite3.connect(sqlite_db_name)

    base_filename = os.path.splitext(os.path.basename(excel_file.name))[0]
    base_filename_clean = clean_name(base_filename)

    preview_data = {}

    for sheet_name, df in dfs.items():
        sheet_name_clean = clean_name(sheet_name)
        table_name = f"{base_filename_clean}_{sheet_name_clean}"

        # Clean column names
        df.columns = [clean_name(col) for col in df.columns]
        df.dropna(how='all', inplace=True)
        df["row_hash"] = df.apply(generate_row_hash, axis=1)

        df.to_sql(table_name, conn, if_exists='replace', index=False)

        preview_data[sheet_name] = df.head()

    conn.commit()
    conn.close()
    return preview_data


def upload_sqlite_to_sharepoint(site_url, folder_url, db_path, username, password):
    ctx_auth = AuthenticationContext(site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, ctx_auth)
        with open(db_path, 'rb') as content_file:
            file_content = content_file.read()
            target_folder = ctx.web.get_folder_by_server_relative_url(folder_url)
            name = os.path.basename(db_path)
            target_folder.upload_file(name, file_content).execute_query()
            st.success(f"Uploaded {name} to SharePoint.")
    else:
        print("Authentication failed.")

def convert_tab():
    st.subheader("Convert Excel to SQLite")
    excel_file = upload_excel()

    if excel_file:
        sqlite_path = "database.sqlite"
        preview_data = transform_and_load_to_sqlite(excel_file, sqlite_db_name=sqlite_path)

        st.success("Excel file has been converted into SQLite database.")
        st.subheader("Preview of the converted sheets:")

        # Display previews
        for sheet_name, df_preview in preview_data.items():
            st.write(f"**Sheet: {sheet_name}**")
            st.dataframe(df_preview)

        if st.button("Upload SQLite file to SharePoint"):
            upload_sqlite_to_sharepoint(SITE_URL, FOLDER_PATH, sqlite_path, USERNAME, PASSWORD)
    
def combine_tab(): 
    st.subheader("Combine Excel Files")

    uploaded_files = st.file_uploader(
        "",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        combined_df = pd.DataFrame()

        for file in uploaded_files:
            # Read all sheets from each uploaded file
            dfs = pd.read_excel(file, sheet_name=None)

            for sheet_name, df in dfs.items():
                df['source_file'] = file.name 
                df['source_sheet'] = sheet_name
                combined_df = pd.concat([combined_df, df], ignore_index=True)

        st.success(f"Successfully combined {len(uploaded_files)} files!")

        # Show combined data
        st.subheader("Preview of Combined Data")
        st.dataframe(combined_df)

        if st.button("Create SQLite and Upload to SharePoint"):
            with tempfile.NamedTemporaryFile(suffix=".sqlite", delete=False) as tmpfile:
                sqlite_path = tmpfile.name

            # Save combined_df into SQLite
            conn = sqlite3.connect(sqlite_path)
            combined_df.to_sql('combined_data', conn, index=False, if_exists='replace')
            conn.close()

            st.success("SQLite database created successfully!")

            # Now upload the SQLite file to SharePoint
            upload_sqlite_to_sharepoint(SITE_URL, FOLDER_PATH, sqlite_path, USERNAME, PASSWORD)

            # Optionally delete the temp file after upload
            os.remove(sqlite_path)

            st.success("SQLite database uploaded to SharePoint successfully!")
        
    else:
        st.info("Please upload Excel files to start combining.")

def embed_tab():
    st.subheader("Embed Files into SQLite Database")

    uploaded_file = st.file_uploader("")

    if uploaded_file:
        file_name = uploaded_file.name
        file_data = uploaded_file.read()  # This will read the file as bytes (BLOB)

        # Connect to SQLite
        conn = sqlite3.connect("database.sqlite")
        cursor = conn.cursor()

        # Create a table if it doesn't exist
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS embedded_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT,
                file_blob BLOB
            )
        ''')

        # Insert file into the table
        cursor.execute('''
            INSERT INTO embedded_files (file_name, file_blob) 
            VALUES (?, ?)
        ''', (file_name, file_data))

        conn.commit()
        conn.close()

        st.success(f"Successfully embedded file: {file_name}")

def view_tab():
    st.subheader("View SQLite Database")

    conn = sqlite3.connect("database.sqlite")
    cursor = conn.cursor()

    # Fetch all table names
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [table[0] for table in cursor.fetchall()]
    conn.close()

    if not tables:
        st.info("No tables found in the database.")
        return

    # User selects a table to view
    selected_table = st.selectbox("Select a table to view:", tables)

    if selected_table:
        conn = sqlite3.connect("database.sqlite")
        cursor = conn.cursor()

        try:
            cursor.execute(f"SELECT * FROM '{selected_table}'")  # Always wrap table names in quotes
            rows = cursor.fetchall()
            columns = [desc[0] for desc in cursor.description]

            df = pd.DataFrame(rows, columns=columns)
            st.dataframe(df)
        except Exception as e:
            st.error(f"Error loading table {selected_table}: {str(e)}")
        finally:
            conn.close()
    
    st.divider()  # A visual divider line (optional)

    # --- Viewing Section ---
    st.subheader("List of Embedded Files")
    # st.write("Click on the file to download.")
    
    conn = sqlite3.connect("database.sqlite")
    cursor = conn.cursor()

    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='embedded_files';")
    if cursor.fetchone():
        cursor.execute("SELECT id, file_name, file_blob FROM embedded_files")
        rows = cursor.fetchall()

        if rows:
            for file_id, file_name, file_blob in rows:
                # Encode file_blob to base64
                b64 = base64.b64encode(file_blob).decode()

                # Create downloadable link
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">📄 {file_name}</a>'
                st.markdown(href, unsafe_allow_html=True)

        else:
            st.info("No files embedded yet.")
    else:
        st.info("No files embedded yet.")

    conn.close()

def main():
    selected = option_menu( 
        menu_title=None, 
        # options=["Convert", "Combine", "Embed", "View"],
        options=["Convert", "Embed", "View"],
        icons=["file-earmark-spreadsheet", "file-earmark-plus", "file-earmark-lock", "eye"],
        menu_icon="cast",
        default_index=0,  
        orientation="horizontal"
    )    

    if selected == "Convert":
        convert_tab()
    # elif selected == "Combine":
    #     combine_tab()
    elif selected == "Embed":
        embed_tab()
    elif selected == "View":
        view_tab()
    
if __name__ == "__main__":
    st.set_page_config(page_title="SQLite Database Manager", layout="wide")
    main()
    
    