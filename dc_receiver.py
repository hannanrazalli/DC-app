import customtkinter as ctk
from tkinter import messagebox
import pandas as pd
import os
import threading
from datetime import datetime

# ==========================================
# ⚙️ CONFIGURATION
# ==========================================

# SERVER PATH
SERVER_ROOT = r"\\192.168.2.3\Drawings"

# 1. FAIL MASTER (SOURCE OF TRUTH - EXCEL)
PATH_MASTER_EXCEL = os.path.join(SERVER_ROOT, r"[04] ENGINEERING TEAM\[98] DOCUMENT CONTROL\All Projects Drawings Data.xlsx")

# 2. FAIL RECEIVE STATUS (DESTINATION - CSV)
PATH_STATUS_CSV = os.path.join(SERVER_ROOT, r"[04] ENGINEERING TEAM\[98] DOCUMENT CONTROL\Document Control Receive Status.csv")

# 3. PATH DRAWING (PDF)
PATH_PDF_FOLDER = os.path.join(SERVER_ROOT, r"[04] ENGINEERING TEAM\[98] DOCUMENT CONTROL\[00] 2D Drawings - Signed Copy")

# 4. MAPPING HEADER
# Updated based on user request
COL_MAPPING = {
    "PART_NO": "Part Number",
    "REVISION": "Revision",
    "PROJECT": "Project Name",       # Updated: "Project" -> "Project Name"
    "ASSEMBLY": "Main Assembly",
    "DRAWING_NAME": "Drawing Name",
    "TOTAL_SHEET": "Total Sheets"    # Updated: "Total Sheet" -> "Total Sheets"
}

# ==========================================

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class DrawingCard(ctk.CTkFrame):
    """Kad untuk paparkan drawing yang BELUM di-sync"""
    def __init__(self, master, data, on_receive, on_view):
        super().__init__(master, fg_color="#2b2b2b", corner_radius=6, border_width=1, border_color="#505050")
        self.pack(fill="x", padx=5, pady=4)

        self.data = data
        self.on_receive_callback = on_receive
        
        # Safe extraction
        part = str(data.get(COL_MAPPING['PART_NO'], '-'))
        rev = str(data.get(COL_MAPPING['REVISION'], '-'))
        name = str(data.get(COL_MAPPING['DRAWING_NAME'], ''))
        
        if len(name) > 50: name = name[:47] + "..."

        # --- UI LAYOUT ---
        # Left: Info
        info_frame = ctk.CTkFrame(self, fg_color="transparent")
        info_frame.pack(side="left", padx=10, pady=5)
        
        row1 = ctk.CTkFrame(info_frame, fg_color="transparent")
        row1.pack(anchor="w")
        ctk.CTkLabel(row1, text=part, font=("Arial", 14, "bold"), text_color="white").pack(side="left")
        ctk.CTkLabel(row1, text=f"  (Rev: {rev})", font=("Arial", 12, "bold"), text_color="#aaaaaa").pack(side="left")
        
        ctk.CTkLabel(info_frame, text=name, font=("Arial", 12), text_color="gray").pack(anchor="w")

        # Right: Buttons
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(side="right", padx=10)

        self.btn_view = ctk.CTkButton(self.btn_frame, text="View PDF", width=80, height=30, fg_color="#3B8ED0", 
                      command=lambda: on_view(data))
        self.btn_view.pack(side="left", padx=5)
        
        self.btn_receive = ctk.CTkButton(self.btn_frame, text="RECEIVE", width=100, height=30, 
                                       fg_color="#E0A710", hover_color="#B0850C",
                                       command=self.perform_receive)
        self.btn_receive.pack(side="left", padx=5)

    def perform_receive(self):
        # Panggil fungsi backend
        success = self.on_receive_callback(self.data)
        
        if success:
            # Tukar status butang (Visual Feedback)
            self.btn_receive.configure(text="RECEIVED ✓", fg_color="#00C853", state="disabled")
            self.configure(border_color="#00C853", border_width=2)

class DCFastSyncApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("DC Fast Sync (Master vs CSV)")
        self.geometry("1200x800")

        self.df_master = pd.DataFrame()
        self.df_missing = pd.DataFrame() 
        
        self.setup_ui()

    def setup_ui(self):
        # --- LEFT SIDEBAR ---
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        
        ctk.CTkLabel(self.sidebar, text="SYNC CONTROL", font=("Arial", 20, "bold")).pack(pady=30)
        
        self.btn_scan = ctk.CTkButton(self.sidebar, text="🔍 SCAN UPDATES", height=50, 
                                      fg_color="#2CC985", hover_color="#20A065", font=("Arial", 14, "bold"),
                                      command=self.start_scan_thread)
        self.btn_scan.pack(padx=20, pady=20)
        
        self.lbl_status = ctk.CTkLabel(self.sidebar, text="Status: Idle", text_color="gray")
        self.lbl_status.pack(side="bottom", pady=20)
        
        self.lbl_count = ctk.CTkLabel(self.sidebar, text="-", font=("Arial", 18, "bold"), text_color="#FF4444")
        self.lbl_count.pack(pady=20)

        # --- RIGHT MAIN AREA ---
        self.main_area = ctk.CTkFrame(self, fg_color="transparent")
        self.main_area.pack(side="right", fill="both", expand=True)

        # Top Bar
        self.top_bar = ctk.CTkFrame(self.main_area, height=60, fg_color="transparent")
        self.top_bar.pack(fill="x", padx=10, pady=10)

        self.search_var = ctk.StringVar()
        self.search_var.trace("w", lambda *args: self.search_var.set(self.search_var.get().upper()))

        self.entry_search = ctk.CTkEntry(self.top_bar, textvariable=self.search_var, placeholder_text="Search missing items...", width=400)
        self.entry_search.pack(side="left", padx=5)
        self.entry_search.bind("<Return>", lambda e: self.perform_search())
        
        ctk.CTkButton(self.top_bar, text="SEARCH", width=80, command=self.perform_search).pack(side="left", padx=5)
        ctk.CTkButton(self.top_bar, text="CLEAR", width=60, fg_color="gray", command=self.clear_search).pack(side="left", padx=5)

        # Tab View
        self.tab_view = ctk.CTkTabview(self.main_area)
        self.tab_view.pack(fill="both", expand=True, padx=10, pady=5)
        self.tab_view.add("Home")
        ctk.CTkLabel(self.tab_view.tab("Home"), text="Press 'SCAN UPDATES' to check for missing drawings.").pack(pady=40)

    def start_scan_thread(self):
        self.btn_scan.configure(state="disabled", text="READING EXCEL...")
        self.lbl_status.configure(text="Reading Master Excel...")
        threading.Thread(target=self.process_comparison).start()

    def process_comparison(self):
        try:
            # 1. READ CSV (Destination)
            if os.path.exists(PATH_STATUS_CSV):
                df_csv = pd.read_csv(PATH_STATUS_CSV, dtype=str)
                df_csv.columns = df_csv.columns.str.strip()
            else:
                df_csv = pd.DataFrame(columns=list(COL_MAPPING.values()))
                # Create file immediately with headers if not exists
                df_csv.to_csv(PATH_STATUS_CSV, index=False)

            # 2. READ EXCEL (Source)
            if not os.path.exists(PATH_MASTER_EXCEL):
                self.after(0, lambda: messagebox.showerror("File Error", f"Master Excel not found:\n{PATH_MASTER_EXCEL}"))
                return
            
            # Read Excel - Force String
            df_master = pd.read_excel(PATH_MASTER_EXCEL, dtype=str)
            df_master.columns = df_master.columns.str.strip()
            df_master = df_master.fillna("")

            # --- DEBUG COLUMN MAPPING ---
            required_cols = [COL_MAPPING['PART_NO'], COL_MAPPING['REVISION'], COL_MAPPING['PROJECT']]
            missing_cols = [c for c in required_cols if c not in df_master.columns]
            
            if missing_cols:
                # Papar error popup jika column tak jumpa
                found_cols = list(df_master.columns)
                msg = f"ERROR: Column names do not match!\n\nMissing: {missing_cols}\n\nFound in Excel: {found_cols}\n\nPlease update 'COL_MAPPING' in the code."
                self.after(0, lambda: messagebox.showerror("Column Mismatch", msg))
                return

            # 3. COMPARE LOGIC (String Comparison)
            # Buang whitespace (cth: "Rev 1 " -> "Rev 1")
            df_master[COL_MAPPING['PART_NO']] = df_master[COL_MAPPING['PART_NO']].str.strip()
            df_master[COL_MAPPING['REVISION']] = df_master[COL_MAPPING['REVISION']].str.strip()
            
            if not df_csv.empty:
                df_csv[COL_MAPPING['PART_NO']] = df_csv[COL_MAPPING['PART_NO']].str.strip()
                df_csv[COL_MAPPING['REVISION']] = df_csv[COL_MAPPING['REVISION']].str.strip()
            
            # Create Unique Keys
            master_keys = df_master[COL_MAPPING['PART_NO']] + "_REV_" + df_master[COL_MAPPING['REVISION']]
            
            if not df_csv.empty:
                csv_keys = set(df_csv[COL_MAPPING['PART_NO']] + "_REV_" + df_csv[COL_MAPPING['REVISION']])
            else:
                csv_keys = set()

            # Find Missing
            is_missing = ~master_keys.isin(csv_keys)
            df_missing = df_master[is_missing].copy()

            self.df_missing = df_missing
            self.after(0, self.finish_scan)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("System Error", f"Detailed error: {str(e)}"))
        finally:
            self.after(0, lambda: self.btn_scan.configure(state="normal", text="🔍 SCAN UPDATES"))

    def finish_scan(self):
        count = len(self.df_missing)
        self.lbl_count.configure(text=f"{count} Missing")
        self.lbl_status.configure(text=f"Updated: {datetime.now().strftime('%H:%M:%S')}")

        if count == 0:
            messagebox.showinfo("Synced", "Great! All drawings in Excel are already in CSV.")
            # Clear tabs
            self.clear_tabs()
            self.tab_view.add("Synced")
            ctk.CTkLabel(self.tab_view.tab("Synced"), text="No pending items.").pack(pady=40)
            return

        self.construct_tabs()

    def clear_tabs(self):
        try:
            for t in list(self.tab_view._tab_dict.keys()):
                self.tab_view.delete(t)
        except: pass

    def construct_tabs(self):
        self.clear_tabs()

        # Group by Project
        projects = sorted(self.df_missing[COL_MAPPING['PROJECT']].astype(str).unique().tolist())
        projects = [p for p in projects if p and p.lower() != 'nan' and p.strip() != '']

        if not projects:
            # Fallback jika project column kosong tapi ada data
            self.tab_view.add("General")
            scroll = ctk.CTkScrollableFrame(self.tab_view.tab("General"))
            scroll.pack(fill="both", expand=True)
            for _, row in self.df_missing.head(100).iterrows():
                DrawingCard(scroll, row, self.process_receive, self.open_pdf)
            return

        for proj in projects:
            self.tab_view.add(proj)
            
            # Filter data for this project
            df_proj = self.df_missing[self.df_missing[COL_MAPPING['PROJECT']] == proj]
            
            # Group by Assembly
            assemblies = sorted(df_proj[COL_MAPPING['ASSEMBLY']].astype(str).unique().tolist())
            assemblies = [a for a in assemblies if a and a.lower() != 'nan' and a.strip() != '']
            
            if not assemblies:
                # No assembly structure
                scroll = ctk.CTkScrollableFrame(self.tab_view.tab(proj))
                scroll.pack(fill="both", expand=True)
                for _, row in df_proj.iterrows():
                    DrawingCard(scroll, row, self.process_receive, self.open_pdf)
            else:
                # Sub-tabs for Assembly
                assy_tabs = ctk.CTkTabview(self.tab_view.tab(proj), height=30)
                assy_tabs.pack(fill="both", expand=True)
                
                for assy in assemblies:
                    assy_tabs.add(assy)
                    scroll = ctk.CTkScrollableFrame(assy_tabs.tab(assy))
                    scroll.pack(fill="both", expand=True)
                    
                    df_assy = df_proj[df_proj[COL_MAPPING['ASSEMBLY']] == assy]
                    for _, row in df_assy.iterrows():
                        DrawingCard(scroll, row, self.process_receive, self.open_pdf)

    def process_receive(self, row_data):
        """Return True jika berjaya save"""
        try:
            df_to_save = pd.DataFrame([row_data])
            # Append to CSV
            df_to_save.to_csv(PATH_STATUS_CSV, mode='a', header=False, index=False)
            
            # Remove from memory (search updates)
            part = row_data[COL_MAPPING['PART_NO']]
            rev = row_data[COL_MAPPING['REVISION']]
            
            idx = self.df_missing[
                (self.df_missing[COL_MAPPING['PART_NO']] == part) & 
                (self.df_missing[COL_MAPPING['REVISION']] == rev)
            ].index
            
            self.df_missing = self.df_missing.drop(idx)
            
            # Update Count Label
            current_count = len(self.df_missing)
            self.lbl_count.configure(text=f"{current_count} Missing")
            
            return True
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
            return False

    def open_pdf(self, row_data):
        part_no = str(row_data.get(COL_MAPPING['PART_NO'], ""))
        project = str(row_data.get(COL_MAPPING['PROJECT'], ""))
        assembly = str(row_data.get(COL_MAPPING['ASSEMBLY'], ""))
        
        target_path = os.path.join(PATH_PDF_FOLDER, project, assembly)
        found_path = None
        
        # Cari file
        if os.path.exists(target_path):
            for f in os.listdir(target_path):
                if f.lower().endswith(".pdf") and part_no in f:
                    found_path = os.path.join(target_path, f)
                    break
        
        # Fallback cari dalam project folder
        if not found_path:
            t_proj = os.path.join(PATH_PDF_FOLDER, project)
            if os.path.exists(t_proj):
                for f in os.listdir(t_proj):
                    if f.lower().endswith(".pdf") and part_no in f:
                        found_path = os.path.join(t_proj, f)
                        break

        if found_path:
            try: os.startfile(found_path)
            except Exception as e: messagebox.showerror("Error", f"Cannot open PDF: {e}")
        else:
            messagebox.showwarning("Not Found", f"PDF not found for {part_no}\nin {project}/{assembly}")

    def perform_search(self):
        query = self.search_var.get().strip()
        if not query or self.df_missing.empty: return

        if "SEARCH RESULTS" in self.tab_view._tab_dict:
            self.tab_view.delete("SEARCH RESULTS")
        
        self.tab_view.add("SEARCH RESULTS")
        self.tab_view.set("SEARCH RESULTS")
        
        scroll = ctk.CTkScrollableFrame(self.tab_view.tab("SEARCH RESULTS"))
        scroll.pack(fill="both", expand=True)

        mask = (
            self.df_missing[COL_MAPPING['PART_NO']].astype(str).str.upper().str.contains(query, na=False) |
            self.df_missing[COL_MAPPING['DRAWING_NAME']].astype(str).str.upper().str.contains(query, na=False)
        )
        df_result = self.df_missing[mask]
        
        ctk.CTkLabel(scroll, text=f"Found {len(df_result)} items.").pack(pady=10)

        for _, row in df_result.head(100).iterrows():
             DrawingCard(scroll, row, self.process_receive, self.open_pdf)

    def clear_search(self):
        self.search_var.set("")
        if "SEARCH RESULTS" in self.tab_view._tab_dict:
            self.tab_view.delete("SEARCH RESULTS")
        try:
            self.tab_view.set(list(self.tab_view._tab_dict.keys())[0])
        except: pass

if __name__ == "__main__":
    app = DCFastSyncApp()
    app.mainloop()