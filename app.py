import base64
import pandas as pd
import hashlib
import io
import os
import re
import sqlite3
import streamlit as st
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from streamlit_option_menu import option_menu

USERNAME = st.secrets["sharepoint"]["username"]
PASSWORD = st.secrets["sharepoint"]["password"]
SITE_URL = st.secrets["sharepoint"]["site_url"]
FOLDER_PATH = st.secrets["sharepoint"]["folder_path"]
MANAGE_DB = st.secrets["sharepoint"]["manage_db"]

SQLITE_FILENAME = "database.sqlite"

# """EXTRACT"""
def upload_excel():
    """Allows user to select and load an Excel file."""
    excel_file = st.file_uploader("", type=["xlsx", "xls"])
    return excel_file

# """TRANSFORM"""
def generate_row_hash(row):
    """Generates a hash from row contents."""
    row_str = "|".join(map(str, row.values))
    return hashlib.md5(row_str.encode()).hexdigest()

def clean_name(name):
    # Replace any non-word character (anything not a letter, number, or underscore) with underscores
    name = re.sub(r'\W+', '_', name)
    # Remove leading underscores just to be clean
    name = name.lstrip('_')
    # If after cleaning it starts with a digit, add an underscore in front
    if name and name[0].isdigit():
        name = '_' + name
    return name.lower()

def deduplicate_columns(columns):
    seen = {}
    new_columns = []
    for col in columns:
        col_clean = col.strip().lower()
        if col_clean in seen:
            seen[col_clean] += 1
            new_columns.append(f"{col_clean}_{seen[col_clean]}")
        else:
            seen[col_clean] = 0
            new_columns.append(col_clean)
    return new_columns

def transform_and_load_to_sqlite(excel_file, SQLITE_FILENAME):
    dfs = pd.read_excel(excel_file, sheet_name=None)
    conn = sqlite3.connect(SQLITE_FILENAME)

    base_filename = os.path.splitext(os.path.basename(excel_file.name))[0]
    base_filename_clean = clean_name(base_filename)

    preview_data = {}
    

    for sheet_name, df in dfs.items():
        sheet_name_clean = clean_name(sheet_name)
        table_name = f"{base_filename_clean}_{sheet_name_clean}"

        # Clean column names
        df.columns = [col.strip() for col in df.columns]    # Step 1: strip spaces
        df.columns = deduplicate_columns(df.columns)        # Step 2: deduplicate duplicates
        df.columns = [clean_name(col) for col in df.columns] # Step 3: clean special characters

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
        preview_data = transform_and_load_to_sqlite(excel_file, SQLITE_FILENAME)

        st.success("Excel file has been converted into SQLite database.")
        st.subheader("Preview of the converted sheets:")

        # Display previews
        for sheet_name, df_preview in preview_data.items():
            st.write(f"**Sheet: {sheet_name}**")
            st.dataframe(df_preview)

        if st.button("Upload SQLite file to SharePoint"):
            upload_sqlite_to_sharepoint(SITE_URL, FOLDER_PATH, SQLITE_FILENAME, USERNAME, PASSWORD)
            
def combine_excel_files(conn):            
    st.subheader("Combine Excel Files")

    # Upload two Excel files
    file1 = st.file_uploader("Upload First Excel File", type=["xlsx", "xls"], key="file1")
    file2 = st.file_uploader("Upload Second Excel File", type=["xlsx", "xls"], key="file2")

    if file1 and file2:
        # Read both Excel files
        dfs1 = pd.read_excel(file1, sheet_name=None)
        dfs2 = pd.read_excel(file2, sheet_name=None)

        # Normalize sheet names to lowercase
        dfs1 = {sheet.lower(): df for sheet, df in dfs1.items()}
        dfs2 = {sheet.lower(): df for sheet, df in dfs2.items()}

        # Check if both files have the same sheet names
        if set(dfs1.keys()) != set(dfs2.keys()):
            st.error("Files do not have the same sheet names (case-insensitive check). Cannot combine.")
            return

        combined_data = {}
        header_mismatch = False

        # For each normalized sheet
        for sheet_name in dfs1.keys():
            df1 = dfs1[sheet_name]
            df2 = dfs2[sheet_name]

            # Normalize column headers to lowercase
            df1.columns = [col.strip().lower() for col in df1.columns]
            df2.columns = [col.strip().lower() for col in df2.columns]

            # Compare headers
            if list(df1.columns) != list(df2.columns):
                st.error(f"Header mismatch in sheet: {sheet_name}")
                header_mismatch = True
                break
            else:
                combined_df = pd.concat([df1, df2], ignore_index=True)
                combined_df["row_hash"] = combined_df.apply(generate_row_hash, axis=1)
                combined_data[sheet_name] = combined_df

        if not header_mismatch:
            # Connect to SQLite and check existing tables
            conn = sqlite3.connect(SQLITE_FILENAME)
            existing_tables = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn)["name"].tolist()

            table_names = {}
            for sheet_name in combined_data.keys():
                table_name_input = st.text_input(f"Enter table name for sheet '{sheet_name}'", key=f"{sheet_name}_table_name")
                clean_table_name = clean_name(table_name_input) if table_name_input else None

                if not table_name_input:
                    st.warning(f"Please enter a table name for sheet '{sheet_name}'.")
                elif clean_table_name in existing_tables:
                    st.error(f"Table name '{clean_table_name}' already exists in the database.")
                else:
                    table_names[sheet_name] = clean_table_name

            # Proceed to save to SQLite if all names are valid and unique
            if len(table_names) == len(combined_data):
                for sheet_name, df in combined_data.items():
                    table_name = table_names.get(sheet_name)
                    if table_name:
                        df.columns = [clean_name(col) for col in df.columns]  # Clean column names
                        df.columns = deduplicate_columns(df.columns)          # Deduplicate after cleaning
                        df.to_sql(table_name, conn, if_exists="replace", index=False)

                conn.commit()
                conn.close()

                st.success(f"Combined data saved as `{SQLITE_FILENAME}`")

                with open(SQLITE_FILENAME, "rb") as f:
                    st.download_button("Download SQLite file", f, file_name=SQLITE_FILENAME)

                if st.button("Upload to SharePoint"):
                    upload_sqlite_to_sharepoint(SITE_URL, FOLDER_PATH, SQLITE_FILENAME, USERNAME, PASSWORD)
                    st.success(f"Uploaded `{SQLITE_FILENAME}` to SharePoint.")
            else:
                conn.close()
                
def append_excel_to_sqlite(conn):
    st.subheader("Append Excel Data to Existing SQLite Table")

    # Connect and list tables first
    conn = sqlite3.connect(SQLITE_FILENAME)
    existing_tables = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn)["name"].tolist()

    if not existing_tables:
        st.error("No existing tables found in the database.")
        conn.close()
        return

    selected_table = st.selectbox("Select a table to append data to", existing_tables)
    file = st.file_uploader("Upload Excel File to Append", type=["xlsx", "xls"], key="append_file")

    if file and selected_table:
        # Load destination table
        target_df = pd.read_sql(f"SELECT * FROM '{selected_table}'", conn)
        target_headers = [col.strip().lower() for col in target_df.columns]

        # Read uploaded Excel and lowercase sheet names
        dfs = pd.read_excel(file, sheet_name=None)
        dfs = {sheet.lower(): df for sheet, df in dfs.items()}

        sheet_to_append = st.selectbox("Select sheet to append", list(dfs.keys()))

        if sheet_to_append:
            source_df = dfs[sheet_to_append]
            source_df.columns = [col.strip().lower() for col in source_df.columns]
            
            if 'row_hash' not in source_df.columns:
                # If not, generate a row_hash for the new data
                source_df['row_hash'] = source_df.apply(generate_row_hash, axis=1)

            if list(source_df.columns) != target_headers:
                st.error("Header mismatch. Cannot append.")
                conn.close()
                return

            # Clean and deduplicate column names before appending
            source_df.columns = [clean_name(col) for col in source_df.columns]
            source_df.columns = deduplicate_columns(source_df.columns)

            # Append to existing table
            source_df.to_sql(selected_table, conn, if_exists="append", index=False)
            conn.commit()
            conn.close()

            st.success(f"Appended data from '{sheet_to_append}' to table '{selected_table}'")

            with open(SQLITE_FILENAME, "rb") as f:
                st.download_button("Download updated SQLite file", f, file_name=SQLITE_FILENAME)

            if st.button("Upload updated DB to SharePoint"):
                upload_sqlite_to_sharepoint(SITE_URL, FOLDER_PATH, SQLITE_FILENAME, USERNAME, PASSWORD)
                st.success(f"Uploaded `{SQLITE_FILENAME}` to SharePoint.")

def combine_tab():
    mode = st.radio("Select action", ["Combine Two Excel Files", "Combine Excel to Existing Table"])
    
    with st.container(border=True): 
        conn = sqlite3.connect(SQLITE_FILENAME)

        if mode == "Combine Two Excel Files":
            combine_excel_files(conn)
        elif mode == "Combine Excel to Existing Table":
            append_excel_to_sqlite(conn)

        conn.close()

def embed_tab():
    st.subheader("Embed Files into SQLite Database")

    uploaded_file = st.file_uploader("")

    if uploaded_file:
        file_name = uploaded_file.name
        file_data = uploaded_file.read()  # This will read the file as bytes (BLOB)

        # Connect to SQLite
        conn = sqlite3.connect(SQLITE_FILENAME)
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

    conn = sqlite3.connect(SQLITE_FILENAME)
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
        conn = sqlite3.connect(SQLITE_FILENAME)
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
    
    st.divider() 

    # --- Viewing Section ---
    st.subheader("List of Embedded Files")
    # st.write("Click on the file to download.")
    
    conn = sqlite3.connect(SQLITE_FILENAME)
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
    
def export_database_modal():
    password = st.text_input("Password", type="password", key="export_pw")
    if st.button("Submit", key="submit_export"):
        if password == MANAGE_DB:
            conn = sqlite3.connect(SQLITE_FILENAME)
            tables = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn)["name"].tolist()

            if not tables:
                st.warning("No tables found.")
                return

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                for table in tables:
                    df = pd.read_sql(f'SELECT * FROM "{table}"', conn)
                    df.to_excel(writer, sheet_name=table[:31], index=False)
            conn.close()

            st.download_button(
                "Download Database as Excel",
                data=excel_buffer.getvalue(),
                file_name="database_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Incorrect password.")

def delete_database_modal():
    password = st.text_input("Password", type="password", key="delete_pw")
    if st.button("Delete", key="submit_delete"):
        if password == MANAGE_DB:
            try:
                os.remove(SQLITE_FILENAME)
                st.success("Database deleted successfully.")
            except Exception as e:
                st.error(f"Error deleting database: {e}")
        else:
            st.error("Incorrect password.")

def manage_db_tab():
    st.subheader("Manage SQLite Database")

    if not os.path.exists(SQLITE_FILENAME):
        st.warning("Database does not exist.")
        return

    with st.container(border=True):
        st.markdown("### Export Database")
        if st.button("Export Database as Excel"):
            export_database_modal()

    with st.container(border=True):
        st.markdown("### Delete Entire Database")
        if st.button("Delete Database"):
            delete_database_modal()
    
# def manage_db_tab():
#     st.subheader("Manage SQLite Database")
#     if not os.path.exists(SQLITE_FILENAME):
#         st.warning("Database does not exist.")
#         return

#     # Export database to Excel
#     with st.container(border=True):

#         conn = sqlite3.connect(SQLITE_FILENAME)
#         query = "SELECT name FROM sqlite_master WHERE type='table';"
#         tables = pd.read_sql(query, conn)['name'].tolist()

#         if not tables:
#             st.info("No tables found in the database.")
#         else:
#             excel_buffer = io.BytesIO()
#             with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
#                 for table in tables:
#                     df = pd.read_sql(f'SELECT * FROM "{table}"', conn)
#                     df.to_excel(writer, sheet_name=table[:31], index=False)  # Excel limit: 31 char sheet name
#                 # writer.save()
#             conn.close()

#             st.download_button(
#                 "Download Entire Database as Excel",
#                 data=excel_buffer.getvalue(),
#                 file_name="database_export.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#     # Delete database with password
#     with st.container(border=True):
#         st.markdown("### ❌ Delete Entire Database")

#         password = st.text_input("Enter password to delete the database", type="password")
#         if st.button("Delete Database"):
#             if password == DELETE_DB:
#                 try:
#                     os.remove(SQLITE_FILENAME)
#                     st.success("Database deleted successfully.")
#                 except Exception as e:
#                     st.error(f"Error deleting database: {e}")
#             else:
#                 st.error("Incorrect password. Database not deleted.")

def main():
    selected = option_menu( 
        menu_title=None, 
        options=["Convert", "Combine", "Embed", "View", "Manage"],
        icons=["file-earmark-spreadsheet", "file-earmark-plus", "file-earmark-lock", "eye", "database"],
        menu_icon="cast",
        default_index=0,  
        orientation="horizontal"
    )    

    if selected == "Convert":
        convert_tab()
    elif selected == "Combine":
        combine_tab()
    elif selected == "Embed":
        embed_tab()
    elif selected == "View":
        view_tab()
    elif selected == "Manage":
       manage_db_tab() 
       
       
if __name__ == "__main__":
    st.set_page_config(page_title="SQLite Database Manager", layout="wide")
    main()
    
    