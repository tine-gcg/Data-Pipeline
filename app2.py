import pandas as pd
import sqlite3
import hashlib
from tkinter import Tk, filedialog, Label, Button




def generate_row_hash(row):
    return hashlib.sha256(str(tuple(row)).encode("utf-8")).hexdigest()


def clean_column_name(col):
    return str(col).strip().replace(" ", "_").lower()


def load_excel_to_sqlite(excel_file, db_name="backup.sqlite"):
    dfs = pd.read_excel(excel_file, sheet_name=None)
    conn = sqlite3.connect(db_name)
    base_file = os.path.splitext(os.path.basename(excel_file))[0].lower()

    for sheet_name, df in dfs.items():
        if df.empty:
            continue

        df.columns = [clean_column_name(c) for c in df.columns]
        df.dropna(how='all', inplace=True)

        table_name = f"{base_file}_{sheet_name.strip().replace(' ', '_').lower()}"

        # Generate row_hash for each row
        df["row_hash"] = df.apply(generate_row_hash, axis=1)

        try:
            existing = pd.read_sql(f"SELECT row_hash FROM '{table_name}'", conn)
            existing_hashes = set(existing["row_hash"].tolist())
        except Exception:
            existing_hashes = set()

        new_rows = df[~df["row_hash"].isin(existing_hashes)]

        if not new_rows.empty:
            new_rows.to_sql(table_name, conn, if_exists='append', index=False)

    conn.close()
    return True


def select_file_and_load(label):
    Tk().withdraw()
    excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not excel_file:
        label.config(text="No file selected.")
        return

    label.config(text=f"Processing: {excel_file}")
    try:
        load_excel_to_sqlite(excel_file)
        label.config(text="✅ Successfully updated the database!")
    except Exception as e:
        label.config(text=f"❌ Failed: {e}")


def create_gui():
    root = Tk()
    root.title("Excel to SQLite Backup")
    root.geometry("400x200")

    label = Label(root, text="Click to upload Excel file", wraplength=300)
    label.pack(pady=20)

    btn = Button(root, text="Upload Excel File", command=lambda: select_file_and_load(label))
    btn.pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    import os
    create_gui()
