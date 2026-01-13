import os
import shutil
from datetime import datetime

# ==========================================
# KONFIGURASI LOKASI (FASA 1 & 2)
# ==========================================

# Lokasi folder asal (seperti yang anda berikan)
source_folder = r"C:\Users\HP\Documents\[01] Document Control\Master Template"

# Nama fail Master Template (Sila ubah sambungan .xlsx jika ia adalah Word .docx)
master_filename = "Master Template.xlsx"

# Lokasi di mana fail baharu akan disimpan (Output Folder)
# Kod ini akan mencipta folder 'Generated Files' jika belum wujud
output_folder = os.path.join(source_folder, "Generated Files")

# ==========================================
# FUNGSI UTAMA (FASA 2)
# ==========================================

def generate_new_file(project_name="Projek_Baru"):
    """
    Fungsi ini menyalin Master Template dan menamakannya semula
    berdasarkan nama projek dan tarikh hari ini.
    """
    
    # 1. Pastikan folder output wujud
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"[INFO] Created output folder: {output_folder}")

    # 2. Tetapkan path penuh untuk sumber (Master)
    source_path = os.path.join(source_folder, master_filename)

    # 3. Semak jika Master Template wujud
    if not os.path.exists(source_path):
        print(f"[ERROR] Master Template is not found in: {source_path}")
        return

    # 4. Cipta nama fail baharu (Format: YYYY-MM-DD_NamaProjek_Master.xlsx)
    today_date = datetime.now().strftime("%Y-%m-%d")
    new_filename = f"{today_date}_{project_name}_{master_filename}"
    destination_path = os.path.join(output_folder, new_filename)

    # 5. Salin fail (Copy & Rename)
    try:
        shutil.copy2(source_path, destination_path)
        print("-" * 40)
        print(f"[SUCCESS] New File Created!")
        print(f"Location: {destination_path}")
        print("-" * 40)
    except Exception as e:
        print(f"[ERROR] Failed to copy file: {e}")

# ==========================================
# PELAKSANAAN (EXECUTION)
# ==========================================

if __name__ == "__main__":
    print("--- STARTING PHASE 2: CREATING DOCUMENT ---")
    
    # Contoh: Anda boleh letak senarai nama projek di sini untuk generate banyak fail sekali gus
    senarai_projek = ["H10_TRC", "H10_BERAPIT", "M10", "N10", "WHEEL_PRESS_MACHINE"]

    for projek in senarai_projek:
        generate_new_file(projek)
    
    print("--- FINISHED ---")