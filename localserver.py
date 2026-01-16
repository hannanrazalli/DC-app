import os
import shutil  # Library for file operations (Copy/Paste)
import time

# ==========================================
# CONFIGURATION
# ==========================================
SERVER_PATH = r"Y:\[04] ENGINEERING TEAM\[98] DOCUMENT CONTROL"
LOCAL_PATH = r"C:\Users\HP\Documents\[01] Document Control"

# MAPPING: "Local File Name" : "Server File Name"
# Updated based on your latest request
file_mapping = {
    "Document Control.xlsx": "All Projects Drawings Data.xlsx",
    "H10 Berapit.xlsx": "H10 BeraPit Drawing List.xlsx",
    "H10 TRC.xlsx": "H10 TRC Drawing List.xlsx",
    "M10.xlsx": "M10(N) Drawing List.xlsx",
    "N10.xlsx": "N10(N) Drawing List.xlsx",
    "Wheel Press Machine.xlsx": "Wheel Press Machine Drawing List.xlsx",
    "G10.xlsx": "G10 Drawing List.xlsx"
}

# ==========================================
# MAIN FUNCTION
# ==========================================
def publish_all_files():
    print("--- START PUBLISH (DIRECT COPY MODE) ---")
    
    # 1. Check if server path exists
    if not os.path.exists(SERVER_PATH):
        print(f"[ERROR] Server path not found: {SERVER_PATH}")
        print("Please check your network connection.")
        return

    success_count = 0
    fail_count = 0

    for local_file, server_file in file_mapping.items():
        local_full_path = os.path.join(LOCAL_PATH, local_file)
        server_full_path = os.path.join(SERVER_PATH, server_file)

        print(f"\n[PROCESS] {local_file} --> Server...")

        # 2. Check if local file exists
        if not os.path.exists(local_full_path):
            print(f"   [MISSING] Local file not found. Skipping.")
            fail_count += 1
            continue

        # 3. Check for "File Lock" (Is someone else opening the server file?)
        file_is_locked = False
        if os.path.exists(server_full_path):
            try:
                # Try renaming the file to itself to test if it's locked
                os.rename(server_full_path, server_full_path)
            except OSError:
                file_is_locked = True
                print(f"   [LOCKED] File on Server is OPEN by another user. Cannot Overwrite.")
                fail_count += 1
                continue

        # 4. Process Copy (Direct Copy)
        if not file_is_locked:
            try:
                # shutil.copy2 copies the file data and metadata (timestamps)
                shutil.copy2(local_full_path, server_full_path)
                print(f"   [OK] Successfully copied to server.")
                success_count += 1
            except Exception as e:
                print(f"   [FAIL] Copy failed: {e}")
                fail_count += 1

    print(f"\n--- FINISHED. Success: {success_count} | Failed/Skipped: {fail_count} ---")
    
    # Optional: Pause specifically so you can read the output window before it closes
    # input("Press Enter to exit...") 

if __name__ == "__main__":
    publish_all_files()