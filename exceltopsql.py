import pandas as pd
from sqlalchemy import create_engine, Date
import os

# --- 1. CONFIGURATION ---
DB_URL = 'postgresql://postgres:1234@localhost:5432/dcde'
FOLDER_PATH = r'C:\Users\HP\Documents\[01] Document Control'
MASTER_FILE_NAME = 'Document Control.xlsx'

# Tambah semua file project kau kat sini
TARGET_FILES = [
    'H10 Berapit.xlsx', 
    'H10 TRC.xlsx', 
    'M10.xlsx', 
    'N10.xlsx', 
    'Wheel Press Machine.xlsx'
]

engine = create_engine(DB_URL)

def run_full_pipeline():
    print(f"--- [START] Full Data Pipeline ---")
    
    all_sub_data = []

    # --- LANGKAH 1: SEDUT DATA DARI SEMUA SUB-FILES ---
    for file_name in TARGET_FILES:
        full_path = os.path.join(FOLDER_PATH, file_name)
        
        if not os.path.exists(full_path):
            print(f"  [!] Skip: File {file_name} tidak dijumpai.")
            continue

        print(f"\nProcessing File: {file_name}")
        try:
            xls = pd.ExcelFile(full_path)
            for sheet_name in xls.sheet_names:
                # Baca: Skip Row 1, Header Row 2, Skip 10 Row bawah
                df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=1, skipfooter=10)
                
                # Standardkan nama kolum
                df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

                # Mapping kolum Sub-File ke SQL
                sub_mapping = {
                    'drawing_number': 'part_number',
                    'approved_date': 'date_approved',
                    'total_sheet': 'total_sheets'
                }
                df = df.rename(columns=sub_mapping)

                # Buang baris jika part_number kosong
                df = df.dropna(subset=['part_number'])
                
                if not df.empty:
                    # Metadata punca data
                    df['file_source'] = file_name
                    all_sub_data.append(df)
                    print(f"    -> Sheet '{sheet_name}': {len(df)} rows.")

        except Exception as e:
            print(f"  [ERROR] Gagal proses {file_name}: {e}")

    if not all_sub_data:
        print("Tiada data sub-file ditemui. Pipeline berhenti.")
        return

    # Gabungkan semua data mentah
    final_df = pd.concat(all_sub_data, ignore_index=True)

    # --- LANGKAH 2: BACA MASTER FILE UNTUK ENRICHMENT ---
    print(f"\nReading Master File for Enrichment: {MASTER_FILE_NAME}...")
    try:
        master_path = os.path.join(FOLDER_PATH, MASTER_FILE_NAME)
        df_master = pd.read_excel(master_path)
        df_master.columns = df_master.columns.str.strip().str.lower().str.replace(' ', '_')
        
        # Ambil kolum rujukan
        cols_ref = ['part_number', 'project_name', 'country', 'batch', 'main_assembly', 'engineer']
        df_master = df_master[cols_ref].drop_duplicates(subset=['part_number'])
        
        # --- LANGKAH 3: MAPPING/VLOOKUP (Enrichment) ---
        # Kita join data sub-file dengan master info
        final_df = pd.merge(final_df, df_master, on='part_number', how='left', suffixes=('', '_master'))
        
        # Jika drawing_name dalam sub-file kosong, boleh guna dari master (opsional)
        print("Enrichment selesai.")

    except Exception as e:
        print(f"  [AMARAN] Gagal enrichment dari master: {e}. Data akan masuk tanpa info lengkap.")

    # --- LANGKAH 4: FINAL CLEANING & UPLOAD ---
    # Pastikan kolum ikut struktur SQL kau
    sql_columns = [
        'project_name', 'country', 'batch', 'main_assembly', 'drawing_name', 
        'part_number', 'revision', 'total_sheets', 'engineer', 'date_approved', 'remarks'
    ]

    # Tambah kolum jika tiada
    for col in sql_columns:
        if col not in final_df.columns:
            final_df[col] = None

    # Final touch: Format date & revision
    final_df['date_approved'] = pd.to_datetime(final_df['date_approved'], errors='coerce').dt.date
    final_df['revision'] = pd.to_numeric(final_df['revision'], errors='coerce').fillna(0).astype(int)

    # Pilih kolum yang betul sahaja
    final_df = final_df[sql_columns]
    
    # Buang duplicate final
    final_df = final_df.drop_duplicates(subset=['part_number', 'revision'], keep='last')

    print(f"\nTotal Records sedia untuk SQL: {len(final_df)}")
    
    try:
        # Guna 'replace' supaya table sentiasa fresh dengan data terbaru dari semua file
        final_df.to_sql('project_data', engine, if_exists='replace', index=False, dtype={'date_approved': Date()})
        print("--- [SUCCESS] Full Pipeline Completed! Database 'dcde' is Updated. ---")
    except Exception as e:
        print(f"Error Upload SQL: {e}")

if __name__ == "__main__":
    run_full_pipeline()