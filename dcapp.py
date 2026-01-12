import customtkinter as ctk
import psycopg2
import os
from tkinter import messagebox
from datetime import date, datetime  # Tambah datetime untuk parsing
from tkcalendar import DateEntry
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
# Import MergedCell untuk elak error "read-only"
from openpyxl.cell.cell import MergedCell 

# --- APPEARANCE ---
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# --- CONFIGURATION ---
BASE_PATH = r"C:\Users\HP\Documents\[01] Document Control"
DB_CONFIG = {"dbname": "dcde", "user": "postgres", "password": "1234", "host": "localhost"}
MASTER_FILE = "Document Control.xlsx"

PROJECTS = ["H10 TRC", "H10 BeraPit", "M10(N)", "N10(N)", "Wheel Press Machine"]
PROJ_MAP = {"H10 TRC": "H10", "H10 BeraPit": "H10", "M10(N)": "M10", "N10(N)": "N10", "Wheel Press Machine": "WPM"}
ASSEMBLIES = ["Bogie", "Underframe", "Cabin", "Engine Hood", "Radiator Hood", "Muffler", "Gear Case", "Water Expansion Tank", "Battery Box", "Fuel Tank", "Sand Box"]
ENGINEER_LIST = ["Baskaran", "Sathish", "Harrison", "Hannan", "Gokul", "Vimal", "Ram", "Vishwa", "Bruno"]
REMARKS_LIST = ["New", "Revised", "-"]

MASTER_HEADERS = ["Project", "Country", "Batch", "Main Assembly", "Drawing Name", "Part Number", "Revision", "Total Sheets", "Engineer", "Date Approved", "Remarks"]

# Styles
CREAM_YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

class DCDEApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DCDE Engineering Data Entry System")
        self.geometry("1100x820")
        self.configure(fg_color="#f5f6f7")
        self.grid_columnconfigure((1, 3), weight=1)
        self.setup_ui()

    def setup_ui(self):
        # Row 0
        self.add_label("Project:", 0, 0)
        self.proj_v = ctk.StringVar(value=PROJECTS[0])
        self.proj_drop = ctk.CTkOptionMenu(self, values=PROJECTS, variable=self.proj_v, command=self.update_logic, fg_color="#1f538d")
        self.proj_drop.grid(row=0, column=1, padx=(10, 30), pady=15, sticky="ew")

        self.add_label("Batch (-/N/R):", 0, 2)
        self.batch_ent = ctk.CTkEntry(self)
        self.batch_ent.grid(row=0, column=3, padx=(10, 30), pady=15, sticky="ew")

        # Row 1
        self.add_label("Main Assembly:", 1, 0)
        self.assembly_v = ctk.StringVar(value=ASSEMBLIES[0])
        self.assembly_drop = ctk.CTkOptionMenu(self, values=ASSEMBLIES, variable=self.assembly_v, fg_color="#1f538d")
        self.assembly_drop.grid(row=1, column=1, padx=(10, 30), pady=15, sticky="ew")

        self.add_label("Drawing Name:", 1, 2)
        self.draw_ent = ctk.CTkEntry(self)
        self.draw_ent.grid(row=1, column=3, padx=(10, 30), pady=15, sticky="ew")
        self.draw_ent.bind("<KeyRelease>", lambda e: self.to_uppercase(e, self.draw_ent))

        # Row 2
        self.add_label("Part Number:", 2, 0)
        self.part_ent = ctk.CTkEntry(self)
        self.part_ent.grid(row=2, column=1, padx=(10, 30), pady=15, sticky="ew")
        self.part_ent.bind("<KeyRelease>", lambda e: self.to_uppercase(e, self.part_ent))

        self.add_label("Revision:", 2, 2)
        self.rev_ent = ctk.CTkEntry(self)
        self.rev_ent.grid(row=2, column=3, padx=(10, 30), pady=15, sticky="ew")
        self.rev_ent.bind("<KeyRelease>", self.auto_remark_logic)

        # Row 3
        self.add_label("Total Sheets:", 3, 0)
        self.total_ent = ctk.CTkEntry(self)
        self.total_ent.grid(row=3, column=1, padx=(10, 30), pady=15, sticky="ew")

        self.add_label("Engineer:", 3, 2)
        self.eng_v = ctk.StringVar(value=ENGINEER_LIST[0])
        self.eng_drop = ctk.CTkOptionMenu(self, values=ENGINEER_LIST, variable=self.eng_v, fg_color="#1f538d")
        self.eng_drop.grid(row=3, column=3, padx=(10, 30), pady=15, sticky="ew")

        # Row 4
        self.add_label("Remarks:", 4, 0)
        self.remark_v = ctk.StringVar(value="New")
        self.remark_drop = ctk.CTkOptionMenu(self, values=REMARKS_LIST, variable=self.remark_v, fg_color="#1f538d")
        self.remark_drop.grid(row=4, column=1, padx=(10, 30), pady=15, sticky="ew")

        self.add_label("Date Approved:", 4, 2)
        self.date_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.date_frame.grid(row=4, column=3, padx=(10, 30), pady=15, sticky="w")
        self.date_picker = DateEntry(self.date_frame, width=15, background='#1f538d', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.date_picker.pack(side="left", padx=(0, 10))
        ctk.CTkButton(self.date_frame, text="Today", width=70, command=self.set_today, fg_color="#5d6d7e").pack(side="left")

        # Buttons
        self.btn_submit = ctk.CTkButton(self, text="SUBMIT DATA", fg_color="#28a745", hover_color="#218838", height=65, font=("Arial", 16, "bold"), command=self.submit)
        self.btn_submit.grid(row=6, column=0, columnspan=2, padx=30, pady=40, sticky="ew")

        self.btn_clear = ctk.CTkButton(self, text="CLEAR ALL", fg_color="#dc3545", hover_color="#c82333", height=65, font=("Arial", 16, "bold"), command=self.clear_all)
        self.btn_clear.grid(row=6, column=2, columnspan=2, padx=30, pady=40, sticky="ew")

        # Status Bar
        self.status_bar = ctk.CTkFrame(self, height=40, fg_color="#ebedef", border_width=1, border_color="#abb2b9")
        self.status_bar.grid(row=7, column=0, columnspan=4, sticky="ew", padx=20, pady=(0, 20))
        self.status_label = ctk.CTkLabel(self.status_bar, text="Ready", font=("Arial", 13, "bold"))
        self.status_label.pack(expand=True)

    def add_label(self, text, r, c):
        ctk.CTkLabel(self, text=text, font=("Arial", 12, "bold"), text_color="#2e4053").grid(row=r, column=c, padx=(30, 0), pady=15, sticky="w")

    def update_status(self, message, color="#2e4053"):
        self.status_label.configure(text=message, text_color=color)

    def to_uppercase(self, event, widget):
        ignored = ["Control_L", "Control_R", "Shift_L", "Shift_R", "Caps_Lock", "Left", "Right", "Up", "Down", "Home", "End"]
        if event.keysym in ignored or (event.state & 0x0004): return
        pos = widget.index(ctk.INSERT)
        val = widget.get().upper()
        if widget.get() != val:
            widget.delete(0, 'end'); widget.insert(0, val); widget.icursor(pos)

    def auto_remark_logic(self, event):
        val = self.rev_ent.get()
        if val.isdigit(): self.remark_v.set("Revised" if int(val) >= 1 else "New")

    def update_logic(self, choice):
        if choice == "Wheel Press Machine":
            self.assembly_drop.configure(values=["Wheel Press"]); self.assembly_v.set("Wheel Press")
            self.eng_drop.configure(values=["Baskaran"]); self.eng_v.set("Baskaran")
        else:
            self.assembly_drop.configure(values=ASSEMBLIES); self.assembly_v.set(ASSEMBLIES[0])
            self.eng_drop.configure(values=ENGINEER_LIST); self.eng_v.set(ENGINEER_LIST[0])

    def set_today(self): self.date_picker.set_date(date.today())

    def clear_all(self):
        for w in [self.batch_ent, self.draw_ent, self.part_ent, self.rev_ent, self.total_ent]: w.delete(0, 'end')

    def submit(self):
        # 1. AMBIL INPUT UI
        full_name = self.proj_v.get()
        short_proj = PROJ_MAP.get(full_name)
        proj_file = os.path.join(BASE_PATH, f"{full_name}.xlsx")
        master_file_path = os.path.join(BASE_PATH, MASTER_FILE)
        
        draw = self.draw_ent.get().upper()
        part = self.part_ent.get().upper()
        rev_str = self.rev_ent.get()
        total_str = self.total_ent.get()
        
        # --- DATE FORMAT LOGIC ---
        dt_obj = self.date_picker.get_date() # Objek Date Asli
        dt_sql = dt_obj.strftime('%Y-%m-%d') # String untuk SQL (Standard)
        
        # Validation Mudah
        if not all([draw, part, rev_str, total_str]):
            self.update_status("Error: Please fill in mandatory fields (Drawing, Part, Rev, Total)!", "#c0392b")
            return
        
        try:
            rev = int(rev_str)
            total = int(total_str)
        except ValueError:
             self.update_status("Error: Revision and Total Sheets must be numbers!", "#c0392b")
             return

        # Data for SQL (String)
        sql_data = [short_proj, "Tanzania" if full_name == "H10 TRC" else "Malaysia", self.batch_ent.get(), self.assembly_v.get(), draw, part, rev, total, self.eng_v.get(), dt_sql, self.remark_v.get()]
        
        # Data for Excel (Object Date) - Kita hantar object date, bukan string
        excel_master_data = [short_proj, "Tanzania" if full_name == "H10 TRC" else "Malaysia", self.batch_ent.get(), self.assembly_v.get(), draw, part, rev, total, self.eng_v.get(), dt_obj, self.remark_v.get()]

        try:
            # ---------------------------------------------------------
            # 1. SQL INSERT
            # ---------------------------------------------------------
            try:
                conn = psycopg2.connect(**DB_CONFIG)
                cur = conn.cursor()
                cur.execute("INSERT INTO project_data (project_name, country, batch, main_assembly, drawing_name, part_number, revision, total_sheets, engineer, date_approved, remarks) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", sql_data)
                conn.commit(); cur.close(); conn.close()
            except Exception as db_err:
                print(f"Database Error (Skipped): {db_err}")

            # ---------------------------------------------------------
            # 2. DOCUMENT CONTROL (MASTER)
            # ---------------------------------------------------------
            if not os.path.exists(master_file_path):
                wb_m = Workbook(); ws_m = wb_m.active; ws_m.title = "MasterList"; ws_m.append(MASTER_HEADERS); wb_m.save(master_file_path)
            
            wb_m = load_workbook(master_file_path); ws_m = wb_m.active
            ws_m.append(excel_master_data)
            
            # --- APPLY STYLE FOR MASTER LIST ---
            last_row_m = ws_m.max_row
            for col_idx in range(1, len(excel_master_data) + 1):
                cell_m = ws_m.cell(row=last_row_m, column=col_idx)
                
                # Format Date Column (Col 10)
                if col_idx == 10:
                    cell_m.number_format = 'DD/MM/YYYY'

                if col_idx in [4, 5, 6]:
                    cell_m.alignment = LEFT_ALIGN
                else:
                    cell_m.alignment = CENTER_ALIGN
                cell_m.border = THIN_BORDER
            
            wb_m.save(master_file_path)

            # ---------------------------------------------------------
            # 3. PROJECT SPECIFIC FILE (LOGIK UTAMA)
            # ---------------------------------------------------------
            sheet_name = self.assembly_v.get()[:31]
            
            if not os.path.exists(proj_file):
                messagebox.showerror("Error", f"File {full_name}.xlsx not found!\nPlease ensure the file exists and headers (Row 1 & 2) are set up.")
                return

            wb_p = load_workbook(proj_file)
            
            # Semak Sheet Wujud
            if sheet_name not in wb_p.sheetnames:
                 ws_p = wb_p.create_sheet(sheet_name)
                 ws_p.append(["Sl. No", "Drawing Name", "Part Number", "Revision", "Total Drawings", "Date Approved", "Remarks"])
                 ws_p.cell(row=15, column=1, value="Drawing issued by:-")
            else:
                ws_p = wb_p[sheet_name]

            # --- A. CARI LOKASI SIGNATURE BLOCK ---
            sig_row = None
            for r in range(1, ws_p.max_row + 50):
                found_sig = False
                for c in range(1, 6): 
                    val = ws_p.cell(row=r, column=c).value
                    if val and ("issued by" in str(val).lower() or "drawing issued" in str(val).lower()):
                        sig_row = r
                        found_sig = True
                        break
                if found_sig:
                    break
            
            if not sig_row:
                sig_row = max(ws_p.max_row + 2, 4)
                ws_p.cell(row=sig_row, column=1, value="Drawing issued by:-")

            # --- B. SCAN DATA UNTUK OVERRIDE ---
            target_row = None
            is_override = False
            remarks_to_save = self.remark_v.get()

            for r in range(3, sig_row):
                existing_part = ws_p.cell(row=r, column=3).value
                if existing_part and str(existing_part).strip().upper() == part:
                    existing_rev = ws_p.cell(row=r, column=4).value
                    if str(existing_rev) == str(rev):
                        messagebox.showerror("Data Duplication", f"Part '{part}' with Revision '{rev}' already exists at row {r}!\nPlease check again.")
                        return 
                    else:
                        target_row = r
                        is_override = True
                        remarks_to_save = "Revised"
                        break
            
            # --- C. JIKA TIADA MATCH (DATA BARU) ---
            if not target_row:
                last_data_row = 2
                start_scan = sig_row - 1
                if start_scan < 2: start_scan = 2

                for r in range(start_scan, 2, -1):
                    val_draw = ws_p.cell(row=r, column=2).value
                    val_part = ws_p.cell(row=r, column=3).value
                    if (val_draw and str(val_draw).strip()) or (val_part and str(val_part).strip()):
                        last_data_row = r
                        break
                
                target_row = last_data_row + 1
                ws_p.insert_rows(target_row)
                
                next_row = target_row + 1
                found_sig_below = False
                for c in range(1, 6):
                    val_below = ws_p.cell(row=next_row, column=c).value
                    if val_below and "issued by" in str(val_below).lower():
                        found_sig_below = True
                        break
                
                if found_sig_below:
                    ws_p.insert_rows(next_row) 

            # --- D. TULIS DATA KE EXCEL ---
            
            if not is_override:
                if target_row == 3:
                    sl_no = 1
                else:
                    try:
                        prev_sl_cell = ws_p.cell(row=target_row-1, column=1)
                        if isinstance(prev_sl_cell, MergedCell):
                             prev_sl = 0 
                        else:
                             prev_sl = prev_sl_cell.value

                        if prev_sl and str(prev_sl).isdigit():
                            sl_no = int(prev_sl) + 1
                        else:
                            sl_no = 1 
                    except:
                        sl_no = 1
            else:
                sl_no = ws_p.cell(row=target_row, column=1).value

            # Data for Final Project File (Uses dt_obj - Date Object)
            final_data = [sl_no, draw, part, rev, total, dt_obj, remarks_to_save]

            for col, val in enumerate(final_data, 1):
                cell = ws_p.cell(row=target_row, column=col)
                if isinstance(cell, MergedCell): continue 
                cell.value = val
                
                # Format Date Column (Col 6)
                if col == 6:
                    cell.number_format = 'DD/MM/YYYY'

                if col in [2, 3]:
                    cell.alignment = LEFT_ALIGN
                else:
                    cell.alignment = CENTER_ALIGN
                cell.border = THIN_BORDER
            
            # --- E. HIGHLIGHT LATEST DATE ---
            # Re-scan signature location
            new_sig_row = None
            for r in range(1, ws_p.max_row + 10):
                for c in range(1, 6):
                    val = ws_p.cell(row=r, column=c).value
                    if val and "issued by" in str(val).lower():
                        new_sig_row = r
                        break
                if new_sig_row: break
            
            if not new_sig_row: new_sig_row = ws_p.max_row 

            # Kumpul Date Objects untuk Comparison (Robust checking)
            dates_objects = []
            for r in range(3, new_sig_row):
                d_val = ws_p.cell(row=r, column=6).value
                parsed_date = None
                
                # FIX ERROR ">" not supported between datetime and date
                # Kita akan tukar semua jadi datetime.date
                if isinstance(d_val, datetime):
                    parsed_date = d_val.date()
                elif isinstance(d_val, date):
                    parsed_date = d_val
                elif d_val:
                    # Cuba parse string lama (backward compatibility)
                    d_str = str(d_val).strip()
                    try: parsed_date = datetime.strptime(d_str, '%d/%m/%Y').date()
                    except ValueError:
                        try: parsed_date = datetime.strptime(d_str, '%Y-%m-%d').date()
                        except ValueError: pass
                
                if parsed_date:
                    dates_objects.append(parsed_date)
            
            latest_date_obj = max(dates_objects) if dates_objects else None

            # Apply Styles & Cleanup Remarks
            for r in range(3, new_sig_row):
                d_val = ws_p.cell(row=r, column=6).value
                current_date_obj = None
                
                # Parse semula date baris ini untuk compare (Guna Logik Sama)
                if isinstance(d_val, datetime):
                    current_date_obj = d_val.date()
                elif isinstance(d_val, date):
                    current_date_obj = d_val
                elif d_val:
                     d_str = str(d_val).strip()
                     try: current_date_obj = datetime.strptime(d_str, '%d/%m/%Y').date()
                     except ValueError:
                        try: current_date_obj = datetime.strptime(d_str, '%Y-%m-%d').date()
                        except ValueError: pass

                is_latest = (current_date_obj == latest_date_obj) and (latest_date_obj is not None)
                
                for c in range(1, 8):
                    cell = ws_p.cell(row=r, column=c)
                    if isinstance(cell, MergedCell): continue
                    
                    # LOGIC: Padam remarks jika bukan latest date
                    if c == 7: 
                        if not is_latest:
                             cell.value = None

                    # Alignment
                    if c in [2, 3]: 
                        cell.alignment = LEFT_ALIGN
                    else:
                        cell.alignment = CENTER_ALIGN

                    if is_latest:
                        cell.fill = CREAM_YELLOW
                        cell.font = Font(bold=True)
                    else:
                        cell.fill = NO_FILL
                        cell.font = Font(bold=False)

            wb_p.save(proj_file)
            
            action_msg = "UPDATED (Override)" if is_override else "SAVED (New)"
            self.update_status(f"Success! Data '{part}' has been {action_msg} at Row {target_row}.", "#27ae60")
            # self.clear_all()  <-- Field kekal

        except PermissionError:
             messagebox.showerror("File Error", f"Please close file {full_name}.xlsx before submitting!")
        except Exception as e:
            self.update_status(f"System Error: {str(e)}", "#c0392b")
            print(e)

if __name__ == "__main__":
    app = DCDEApp()
    app.mainloop()