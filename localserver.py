import os
import win32com.client as win32
import time

# --- CONFIGURATION ---
SERVER_PATH = r"Y:\[04] ENGINEERING TEAM\[98] DOCUMENT CONTROL"
LOCAL_PATH = r"C:\Users\HP\Documents\[01] Document Control"
SERVER_PASSWORD = "1234" # Password yang akan diletakkan pada fail Server

# Mapping: "Nama File Local" : "Nama File Server"
# Pastikan ejaan sebiji (case sensitive kadang-kadang penting)
file_mapping = {
    "Document Control.xlsx": "All Projects Drawings Data.xlsx",
    "H10 Berapit.xlsx": "H10 BeraPit Drawing List.xlsx",
    "H10 TRC.xlsx": "H10 TRC Drawing List.xlsx",
    "M10(N).xlsx": "M10(N) Drawing List.xlsx",
    "N10(N).xlsx": "N10(N) Drawing List.xlsx",
    "Wheel Press Machine.xlsx": "Wheel Press Machine Drawing List.xlsx"
}

def publish_all_files():
    print("--- MULA PUBLISH SEMUA: LOCAL -> SERVER ---")
    
    # Check folder server (Y:)
    if not os.path.exists(SERVER_PATH):
        print("[ERROR] Drive Y: tidak dapat dibaca. Sila check network.")
        return

    try:
        # Buka Excel
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True  # Kita nampak progress
        excel.DisplayAlerts = False # Auto-overwrite file lama tanpa tanya
    except Exception as e:
        print(f"[ERROR] Gagal buka Excel. {e}")
        return

    success_count = 0
    fail_count = 0

    for local_file, server_file in file_mapping.items():
        local_full_path = os.path.join(LOCAL_PATH, local_file)
        server_full_path = os.path.join(SERVER_PATH, server_file)

        print(f"\n[PROCESS] {local_file} --> Server...")

        # 1. Pastikan file local wujud
        if os.path.exists(local_full_path):
            try:
                # 2. Check "Lock" (Adakah orang lain sedang buka file server?)
                file_is_locked = False
                if os.path.exists(server_full_path):
                    try:
                        os.rename(server_full_path, server_full_path)
                    except OSError:
                        file_is_locked = True
                        print(f"   [SKIP] File Server sedang DIBUKA oleh user lain. Tak boleh overwrite.")
                        fail_count += 1
                        continue 

                # 3. Proses Save As
                if not file_is_locked:
                    wb = excel.Workbooks.Open(local_full_path)
                    
                    # Save ke server dan letak password
                    wb.SaveAs(server_full_path, FileFormat=51, Password=SERVER_PASSWORD, WriteResPassword=SERVER_PASSWORD)
                    
                    wb.Close()
                    print(f"   [OK] BERJAYA update server.")
                    success_count += 1
                
            except Exception as e:
                print(f"   [FAIL] Error: {e}")
                fail_count += 1
        else:
            print(f"   [MISSING] File local '{local_file}' tiada. Check ejaan.")
            fail_count += 1

    print(f"\n--- TAMAT. Berjaya: {success_count} | Gagal/Skip: {fail_count} ---")
    # excel.Quit() # Uncomment kalau nak excel tutup sendiri

if __name__ == "__main__":
    publish_all_files()