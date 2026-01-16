import pandas as pd
from sqlalchemy import create_engine
import os

# --- 1. CONFIGURATION ---
DB_URL = 'postgresql://postgres:1234@localhost:5432/dcde'
FOLDER_PATH = r'C:\Users\HP\Documents\[01] Document Control'
MASTER_FILE_NAME = 'Document Control.xlsx'

engine = create_engine(DB_URL)

def enrich_all_null_data():
    print(f"--- Fasa Enrichment: Melengkapkan SEMUA Data NULL dalam SQL ---")

    # 2. BACA SEMUA DATA DARI SQL
    try:
        # Kita ambil semua data dari table project_data
        df_sql = pd.read_sql("SELECT * FROM project_data", engine)
        if df_sql.empty:
            print("Database kosong. Tiada data untuk dikemaskini.")
            return
        print(f"Membaca {len(df_sql)} baris dari database...")
    except Exception as e:
        print(f"Error baca SQL: {e}")
        return

    # 3. BACA MASTER FILE (Library Rujukan)
    try:
        master_path = os.path.join(FOLDER_PATH, MASTER_FILE_NAME)
        df_master = pd.read_excel(master_path)
        
        # Standardkan nama kolum Master (Lower & Underscore)
        df_master.columns = df_master.columns.str.strip().str.lower().str.replace(' ', '_')
        
        # Kolum yang kita nak sedut dari Master
        cols_ref = ['part_number', 'project_name', 'country', 'batch', 'main_assembly', 'engineer']
        
        # Pastikan kita hanya ambil rujukan unik untuk setiap part_number
        df_master = df_master[cols_ref].drop_duplicates(subset=['part_number'])
        
        # Set index pada part_number untuk pencarian pantas (VLOOKUP style)
        df_master.set_index('part_number', inplace=True)
        print("Master File sedia sebagai rujukan.")
    except Exception as e:
        print(f"Error baca Master File: {e}")
        return

    # 4. PROSES PADANAN & UPDATE
    print("Memulakan proses mengemaskini maklumat NULL...")
    updated_count = 0
    
    for index, row in df_sql.iterrows():
        pn = row['part_number']
        
        # Jika part_number ini wujud dalam Master File
        if pn in df_master.index:
            m_info = df_master.loc[pn]
            
            # Kita hanya update jika data sedia ada adalah NULL atau kosong
            # Tapi untuk selamat, kita overwrite semua info master supaya sentiasa 'Up-to-Date'
            df_sql.at[index, 'project_name'] = m_info['project_name']
            df_sql.at[index, 'country'] = m_info['country']
            df_sql.at[index, 'batch'] = m_info['batch']
            df_sql.at[index, 'main_assembly'] = m_info['main_assembly']
            df_sql.at[index, 'engineer'] = m_info['engineer']
            updated_count += 1

    # 5. UPLOAD SEMULA DATA YANG DAH LENGKAP KE SQL
    if updated_count > 0:
        try:
            # Guna 'replace' untuk overwrite table dengan data yang telah diperkayakan (enriched)
            df_sql.to_sql('project_data', engine, if_exists='replace', index=False)
            print(f"\n--- SUCCESS! ---")
            print(f"Sebanyak {updated_count} baris data telah dikemaskini dengan info Master.")
        except Exception as e:
            print(f"Error Upload ke SQL: {e}")
    else:
        print("Tiada padanan Part Number ditemui. Tiada data dikemaskini.")

if __name__ == "__main__":
    enrich_all_null_data()