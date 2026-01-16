import pandas as pd
import os
from sqlalchemy import create_engine

# --- CONFIGURATION ---
FILE_PATH = r'C:\Users\HP\Documents\[01] Document Control'
MASTER_FILE = 'Document Control.xlsx'
SUB_FILES = [
    'G10.xlsx', 
    'H10 Berapit.xlsx', 
    'H10 TRC.xlsx', 
    'M10.xlsx', 
    'N10.xlsx', 
    'Wheel Press Machine.xlsx'
]

# Database Credentials
DB_CONFIG = {
    'user': 'postgres',
    'password': '1234',
    'host': 'localhost',
    'port': '5432',
    'database': 'dcde'
}
TABLE_NAME = 'project_data'

def sync_document_control():
    try:
        print("--- Starting Synchronization Process ---")

        # 1. Load Master File
        master_full_path = os.path.join(FILE_PATH, MASTER_FILE)
        print(f"Reading Master File: {MASTER_FILE}...")
        df_master = pd.read_excel(master_full_path)
        
        # Clean column names (strip spaces)
        df_master.columns = df_master.columns.str.strip()

        # 2. Extract unique drawing names from all Sub-Files
        sub_drawing_names = set()
        for file in SUB_FILES:
            sub_path = os.path.join(FILE_PATH, file)
            if os.path.exists(sub_path):
                print(f"Scanning Sub-File: {file}...")
                # We only need the 'drawing name' column for filtering
                df_sub = pd.read_excel(sub_path, usecols=['drawing name'])
                # Clean and collect names
                names = df_sub['drawing name'].dropna().str.strip().unique()
                sub_drawing_names.update(names)
            else:
                print(f"Warning: File {file} not found. Skipping.")

        print(f"Found {len(sub_drawing_names)} unique drawings in sub-files.")

        # 3. Filter Master File Data
        # We only keep rows where 'drawing name' exists in our sub-files set
        df_filtered = df_master[df_master['drawing name'].str.strip().isin(sub_drawing_names)].copy()

        # 4. Map Excel Columns to SQL Table Columns
        column_mapping = {
            'project name': 'project_name',
            'country': 'country',
            'batch': 'batch',
            'main assembly': 'main_assembly',
            'drawing name': 'drawing_name',
            'part number': 'part_number',
            'revision': 'revision',
            'total sheet': 'total_sheets',
            'engineer': 'engineer',
            'date approved': 'date_approved',
            'remarks': 'remarks'
        }
        
        df_final = df_filtered.rename(columns=column_mapping)
        
        # Keep only the columns that exist in the mapping
        df_final = df_final[list(column_mapping.values())]

        # 5. Upload to PostgreSQL
        connection_string = f"postgresql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
        engine = create_engine(connection_string)

        print(f"Uploading {len(df_final)} records to table '{TABLE_NAME}'...")
        
        # 'append' adds new data. Use 'replace' if you want to clear the table first.
        df_final.to_sql(TABLE_NAME, engine, if_exists='append', index=False)

        print("--- Success: Database Updated ---")

    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    sync_document_control()