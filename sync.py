import os
import win32com.client as win32
import time

# --- CONFIGURATION ---
SERVER_PATH = r"Y:\[04] ENGINEERING TEAM\[98] DOCUMENT CONTROL"
LOCAL_PATH = r"C:\Users\HP\Documents\[01] Document Control"
PASSWORD = "1234"

# Mapping: Server Filename -> Local Filename
file_mapping = {
    "H10 BeraPit Drawing List.xlsx": "H10 Berapit.xlsx",
    "H10 TRC Drawing List.xlsx": "H10 TRC.xlsx",
    "M10(N) Drawing List.xlsx": "M10(N).xlsx"
}

def sync_and_unlock_files():
    print("--- MULA SYNC (GUNA Y: DRIVE) ---")
    
    if not os.path.exists(LOCAL_PATH):
        print(f"[ERROR] Folder local tidak jumpa di {LOCAL_PATH}")
        return
    
    if not os.path.exists(SERVER_PATH):
        print(f"[ERROR] Folder server (Y:) tidak dapat dibaca. Pastikan drive Y: connected.")
        return

    try:
        # Buka Excel (Visible = True supaya nampak progress)
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True  
        excel.DisplayAlerts = False 
    except Exception as e:
        print(f"[ERROR] Gagal buka Excel app. Error: {e}")
        return

    success_count = 0

    for server_file, local_file in file_mapping.items():
        source_full_path = os.path.join(SERVER_PATH, server_file)
        dest_full_path = os.path.join(LOCAL_PATH, local_file)

        print(f"\n[PROCESS] Sedang memproses: {server_file}...")

        if os.path.exists(source_full_path):
            try:
                # 1. Buka File Server
                wb = excel.Workbooks.Open(source_full_path, Password=PASSWORD)
                
                time.sleep(1) # Tunggu sekejap

                # 2. Save As ke Local TANPA Password
                wb.SaveAs(dest_full_path, FileFormat=51, Password='', WriteResPassword='')
                
                wb.Close()
                print(f"   [OK] BERJAYA: {local_file} siap.")
                success_count += 1
                
            except Exception as e:
                print(f"   [FAIL] GAGAL process {local_file}. Error: {e}")
        else:
            print(f"   [MISSING] File server '{server_file}' tak jumpa di Y: drive.")

    print(f"\n--- TAMAT. Berjaya: {success_count}/{len(file_mapping)} ---")
    
    # excel.Quit() # Uncomment kalau nak auto-close Excel

if __name__ == "__main__":
    sync_and_unlock_files()