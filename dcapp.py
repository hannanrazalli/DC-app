import customtkinter as ctk
import psycopg2
import os
from tkinter import messagebox
from datetime import date, datetime
from tkcalendar import DateEntry
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.cell.cell import MergedCell 
from PIL import Image # Diperlukan untuk memaparkan logo

# --- APPEARANCE SETUP ---
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("dark-blue")

# --- CONFIGURATION ---
BASE_PATH = r"C:\Users\HP\Documents\[01] Document Control"
DB_CONFIG = {"dbname": "dcde", "user": "postgres", "password": "1234", "host": "localhost"}
MASTER_FILE = "Document Control.xlsx"
LOGO_FILENAME = "LMG Locomotive Logo.jpeg" # Nama fail logo

PROJECTS = ["H10 TRC", "H10 BeraPit", "M10(N)", "N10(N)", "Wheel Press Machine"]
PROJ_MAP = {"H10 TRC": "H10", "H10 BeraPit": "H10", "M10(N)": "M10", "N10(N)": "N10", "Wheel Press Machine": "WPM"}
ASSEMBLIES = ["Bogie", "Underframe", "Cabin", "Engine Hood", "Radiator Hood", "Muffler", "Gear Case", "Water Expansion Tank", "Battery Box", "Fuel Tank", "Sand Box"]
ENGINEER_LIST = ["Baskaran", "Sathish", "Harrison", "Hannan", "Gokul", "Vimal", "Ram", "Vishwa", "Bruno"]
REMARKS_LIST = ["New", "Revised", "-"]
MASTER_HEADERS = ["Project", "Country", "Batch", "Main Assembly", "Drawing Name", "Part Number", "Revision", "Total Sheets", "Engineer", "Date Approved", "Remarks"]

# --- COLORS & FONTS ---
COLOR_PRIMARY = "#2C3E50"    # Dark Slate Blue (Header)
COLOR_ACCENT = "#3498DB"     # Bright Blue (Highlights)
COLOR_SUCCESS = "#27AE60"    # Emerald Green (Submit)
COLOR_DANGER = "#C0392B"     # Red (Clear)
COLOR_BG = "#ECF0F1"         # Light Grey (Background)
COLOR_CARD = "#FFFFFF"       # White (Card Background)
COLOR_TEXT = "#34495E"       # Dark Grey (Text)

FONT_HEADER = ("Segoe UI", 24, "bold")
FONT_SECTION = ("Segoe UI", 16, "bold")
FONT_LABEL = ("Segoe UI", 12, "bold")
FONT_INPUT = ("Segoe UI", 13)

# Excel Styles
CREAM_YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

class DCDEApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DCDE Engineering Data Entry System")
        
        # --- ADJUSTED DIMENSIONS FOR COMPACT & FULL LOOK ---
        self.geometry("980x760")
        self.minsize(900, 720) 
        
        self.configure(fg_color=COLOR_BG)
        
        # Grid layout for the main window
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Content area expands
        
        self.setup_ui()

    def setup_ui(self):
        # --- 1. HEADER SECTION ---
        self.header_frame = ctk.CTkFrame(self, fg_color=COLOR_PRIMARY, corner_radius=0, height=70) # Slightly shorter header
        self.header_frame.grid(row=0, column=0, sticky="ew")
        self.header_frame.grid_propagate(False)

        self.title_label = ctk.CTkLabel(
            self.header_frame, 
            text="ENGINEERING DOCUMENT CONTROL", 
            font=FONT_HEADER, 
            text_color="white"
        )
        self.title_label.pack(side="left", padx=25, pady=15)
        
        self.subtitle_label = ctk.CTkLabel(
            self.header_frame, 
            text="|  Data Entry System", 
            font=("Segoe UI", 14), 
            text_color="#BDC3C7"
        )
        self.subtitle_label.pack(side="left", pady=20)

        # --- LOGO SETUP ---
        try:
            if os.path.exists(LOGO_FILENAME):
                # Load imej menggunakan PIL
                pil_img = Image.open(LOGO_FILENAME)
                
                # Kira saiz berdasarkan ketinggian teks header
                # Font header size 24 lebih kurang 32-35 pixel visual height
                # Kita set height=40 supaya jelas dan seimbang dalam header 70px
                target_height = 40
                aspect_ratio = pil_img.width / pil_img.height
                target_width = int(target_height * aspect_ratio)
                
                # Convert ke CTkImage untuk paparan tajam (HighDPI support)
                logo_ctk = ctk.CTkImage(light_image=pil_img, 
                                        dark_image=pil_img, 
                                        size=(target_width, target_height))
                
                self.logo_label = ctk.CTkLabel(self.header_frame, text="", image=logo_ctk)
                self.logo_label.pack(side="right", padx=30)
            else:
                print(f"Info: Logo file '{LOGO_FILENAME}' not found. Skipping logo.")
        except Exception as e:
            print(f"Error loading logo: {e}")

        # --- 2. MAIN CONTENT CARD ---
        self.content_frame = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=15, border_width=1, border_color="#BDC3C7")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_columnconfigure((0, 1), weight=1) # Two columns

        # --- SECTION: PROJECT DETAILS ---
        self.create_section_header("1. PROJECT DETAILS", row=0)

        # Project Dropdown
        self.add_input_field(label="Project Name", row=1, col=0)
        self.proj_v = ctk.StringVar(value=PROJECTS[0])
        self.proj_drop = ctk.CTkOptionMenu(self.content_frame, values=PROJECTS, variable=self.proj_v, command=self.update_logic, 
                                           fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.proj_drop.grid(row=2, column=0, padx=20, pady=(5, 15), sticky="ew")

        # Batch Input
        self.add_input_field(label="Batch Code (-/N/R)", row=1, col=1)
        self.batch_ent = ctk.CTkEntry(self.content_frame, font=FONT_INPUT, height=32, placeholder_text="e.g. N")
        self.batch_ent.grid(row=2, column=1, padx=20, pady=(5, 15), sticky="ew")

        # --- SECTION: TECHNICAL DATA ---
        self.create_section_header("2. DRAWING & TECHNICAL INFORMATION", row=3)

        # Main Assembly
        self.add_input_field(label="Main Assembly", row=4, col=0)
        self.assembly_v = ctk.StringVar(value=ASSEMBLIES[0])
        self.assembly_drop = ctk.CTkOptionMenu(self.content_frame, values=ASSEMBLIES, variable=self.assembly_v, 
                                               fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.assembly_drop.grid(row=5, column=0, padx=20, pady=(5, 10), sticky="ew")

        # Drawing Name
        self.add_input_field(label="Drawing Name", row=4, col=1)
        self.draw_ent = ctk.CTkEntry(self.content_frame, font=FONT_INPUT, height=32, placeholder_text="Full drawing title")
        self.draw_ent.grid(row=5, column=1, padx=20, pady=(5, 10), sticky="ew")
        self.draw_ent.bind("<KeyRelease>", lambda e: self.to_uppercase(e, self.draw_ent))

        # Part Number
        self.add_input_field(label="Part Number", row=6, col=0)
        self.part_ent = ctk.CTkEntry(self.content_frame, font=FONT_INPUT, height=32, placeholder_text="e.g. H10-100-001")
        self.part_ent.grid(row=7, column=0, padx=20, pady=(5, 10), sticky="ew")
        self.part_ent.bind("<KeyRelease>", lambda e: self.to_uppercase(e, self.part_ent))

        # Revision & Total Sheets (Side by Side in Column 1)
        self.sub_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.sub_frame.grid(row=6, column=1, rowspan=2, padx=20, pady=0, sticky="ew")
        self.sub_frame.grid_columnconfigure((0, 1), weight=1)

        # Rev
        ctk.CTkLabel(self.sub_frame, text="Revision", font=FONT_LABEL, text_color=COLOR_TEXT).grid(row=0, column=0, sticky="w")
        self.rev_ent = ctk.CTkEntry(self.sub_frame, font=FONT_INPUT, height=32, placeholder_text="0")
        self.rev_ent.grid(row=1, column=0, padx=(0, 10), pady=(5, 10), sticky="ew")
        self.rev_ent.bind("<KeyRelease>", self.auto_remark_logic)

        # Total Sheets
        ctk.CTkLabel(self.sub_frame, text="Total Sheets", font=FONT_LABEL, text_color=COLOR_TEXT).grid(row=0, column=1, sticky="w")
        self.total_ent = ctk.CTkEntry(self.sub_frame, font=FONT_INPUT, height=32, placeholder_text="1")
        self.total_ent.grid(row=1, column=1, padx=(10, 0), pady=(5, 10), sticky="ew")

        # --- SECTION: APPROVAL ---
        self.create_section_header("3. APPROVAL & STATUS", row=8)

        # Engineer
        self.add_input_field(label="Responsible Engineer", row=9, col=0)
        self.eng_v = ctk.StringVar(value=ENGINEER_LIST[0])
        self.eng_drop = ctk.CTkOptionMenu(self.content_frame, values=ENGINEER_LIST, variable=self.eng_v, 
                                          fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.eng_drop.grid(row=10, column=0, padx=20, pady=(5, 10), sticky="ew")

        # Date Approved
        self.add_input_field(label="Date Approved", row=9, col=1)
        self.date_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.date_frame.grid(row=10, column=1, padx=20, pady=(5, 10), sticky="w")
        
        self.date_picker = DateEntry(self.date_frame, width=20, background=COLOR_PRIMARY, foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', font=("Arial", 11))
        self.date_picker.pack(side="left", padx=(0, 15), ipady=3)
        
        self.btn_today = ctk.CTkButton(self.date_frame, text="Set Today", width=100, height=30, fg_color="#95a5a6", hover_color="#7f8c8d", command=self.set_today)
        self.btn_today.pack(side="left")

        # Remarks
        self.add_input_field(label="Remarks", row=11, col=0)
        self.remark_v = ctk.StringVar(value="New")
        self.remark_drop = ctk.CTkOptionMenu(self.content_frame, values=REMARKS_LIST, variable=self.remark_v, 
                                             fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.remark_drop.grid(row=12, column=0, padx=20, pady=(5, 15), sticky="ew")

        # --- 3. ACTION BUTTONS ---
        self.btn_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.btn_frame.grid(row=13, column=0, columnspan=2, pady=(20, 20), padx=20, sticky="ew")
        self.btn_frame.grid_columnconfigure((0, 1), weight=1)

        self.btn_submit = ctk.CTkButton(
            self.btn_frame, 
            text="SUBMIT DATA", 
            fg_color=COLOR_SUCCESS, 
            hover_color="#219150", 
            height=50, 
            corner_radius=8,
            font=("Segoe UI", 15, "bold"), 
            command=self.submit
        )
        self.btn_submit.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        self.btn_clear = ctk.CTkButton(
            self.btn_frame, 
            text="CLEAR FORM", 
            fg_color=COLOR_DANGER, 
            hover_color="#A93226", 
            height=50, 
            corner_radius=8,
            font=("Segoe UI", 15, "bold"), 
            command=self.clear_all
        )
        self.btn_clear.grid(row=0, column=1, padx=(10, 0), sticky="ew")

        # --- 4. STATUS BAR ---
        self.status_bar = ctk.CTkFrame(self, height=30, fg_color="#BDC3C7", corner_radius=0)
        self.status_bar.grid(row=2, column=0, sticky="ew")
        self.status_label = ctk.CTkLabel(self.status_bar, text="System Ready", font=("Consolas", 11), text_color="#2C3E50")
        self.status_label.pack(side="left", padx=20)

    # --- UI HELPER FUNCTIONS ---
    def create_section_header(self, text, row):
        label = ctk.CTkLabel(self.content_frame, text=text, font=FONT_SECTION, text_color=COLOR_PRIMARY, anchor="w")
        label.grid(row=row, column=0, columnspan=2, padx=20, pady=(15, 5), sticky="w")
        
        # Divider Line
        line = ctk.CTkFrame(self.content_frame, height=2, fg_color="#E0E0E0")
        line.grid(row=row+1, column=0, columnspan=2, padx=20, pady=(0, 10), sticky="ew")

    def add_input_field(self, label, row, col):
        ctk.CTkLabel(self.content_frame, text=label, font=FONT_LABEL, text_color=COLOR_TEXT).grid(row=row, column=col, padx=20, pady=(5, 0), sticky="w")

    def update_status(self, message, color=COLOR_TEXT):
        self.status_label.configure(text=f"STATUS: {message}", text_color=color)

    # --- LOGIC FUNCTIONS (UNCHANGED) ---
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
        # 1. GET UI INPUT
        full_name = self.proj_v.get()
        short_proj = PROJ_MAP.get(full_name)
        proj_file = os.path.join(BASE_PATH, f"{full_name}.xlsx")
        master_file_path = os.path.join(BASE_PATH, MASTER_FILE)
        
        draw = self.draw_ent.get().upper()
        part = self.part_ent.get().upper()
        rev_str = self.rev_ent.get()
        total_str = self.total_ent.get()
        
        dt_obj = self.date_picker.get_date()
        dt_sql = dt_obj.strftime('%Y-%m-%d')
        
        # Validation
        if not all([draw, part, rev_str, total_str]):
            self.update_status("Error: Please fill in mandatory fields (Drawing, Part, Rev, Total)!", COLOR_DANGER)
            messagebox.showwarning("Missing Data", "Please fill in all mandatory fields.")
            return
        
        try:
            rev = int(rev_str)
            total = int(total_str)
        except ValueError:
             self.update_status("Error: Revision and Total Sheets must be numbers!", COLOR_DANGER)
             return

        # Data Payloads
        sql_data = [short_proj, "Tanzania" if full_name == "H10 TRC" else "Malaysia", self.batch_ent.get(), self.assembly_v.get(), draw, part, rev, total, self.eng_v.get(), dt_sql, self.remark_v.get()]
        excel_master_data = [short_proj, "Tanzania" if full_name == "H10 TRC" else "Malaysia", self.batch_ent.get(), self.assembly_v.get(), draw, part, rev, total, self.eng_v.get(), dt_obj, self.remark_v.get()]

        try:
            # --- PRE-CHECK: DUPLICATE IN PROJECT FILE ---
            skip_project_update = False
            wb_p = None 
            
            if os.path.exists(proj_file):
                sheet_name = self.assembly_v.get()[:31]
                wb_p = load_workbook(proj_file)
                
                if sheet_name in wb_p.sheetnames:
                    ws_p = wb_p[sheet_name]
                    sig_row = None
                    for r in range(1, ws_p.max_row + 50):
                        found_sig = False
                        for c in range(1, 6): 
                            val = ws_p.cell(row=r, column=c).value
                            if val and ("issued by" in str(val).lower() or "drawing issued" in str(val).lower()):
                                sig_row = r; found_sig = True; break
                        if found_sig: break
                    
                    if not sig_row: sig_row = max(ws_p.max_row + 2, 4)

                    duplicate_found = False
                    for r in range(3, sig_row):
                        existing_part = ws_p.cell(row=r, column=3).value
                        if existing_part and str(existing_part).strip().upper() == part:
                            existing_rev = ws_p.cell(row=r, column=4).value
                            if str(existing_rev) == str(rev):
                                duplicate_found = True
                                break
                    
                    if duplicate_found:
                        response = messagebox.askyesno(
                            "Duplicate Found", 
                            f"Part '{part}' with Revision '{rev}' already exists in the Project File.\n\n"
                            "Do you want to add this entry to the MASTER LIST only?\n"
                            "(Select 'No' to cancel operation)"
                        )
                        if response: skip_project_update = True
                        else:
                            self.update_status("Operation Cancelled (Duplicate Found).", COLOR_DANGER)
                            return 

            # 1. SQL INSERT
            try:
                conn = psycopg2.connect(**DB_CONFIG)
                cur = conn.cursor()
                cur.execute("INSERT INTO project_data (project_name, country, batch, main_assembly, drawing_name, part_number, revision, total_sheets, engineer, date_approved, remarks) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", sql_data)
                conn.commit(); cur.close(); conn.close()
            except Exception as db_err:
                print(f"Database Error (Skipped): {db_err}")

            # 2. DOCUMENT CONTROL (MASTER)
            if not os.path.exists(master_file_path):
                wb_m = Workbook(); ws_m = wb_m.active; ws_m.title = "MasterList"; ws_m.append(MASTER_HEADERS); wb_m.save(master_file_path)
            
            wb_m = load_workbook(master_file_path); ws_m = wb_m.active
            ws_m.append(excel_master_data)
            
            last_row_m = ws_m.max_row
            for col_idx in range(1, len(excel_master_data) + 1):
                cell_m = ws_m.cell(row=last_row_m, column=col_idx)
                if col_idx == 10: cell_m.number_format = 'DD/MM/YYYY'
                cell_m.alignment = LEFT_ALIGN if col_idx in [4, 5, 6] else CENTER_ALIGN
                cell_m.border = THIN_BORDER
            wb_m.save(master_file_path)

            # 3. PROJECT SPECIFIC FILE
            if skip_project_update:
                self.update_status(f"Success! Added to Master List ONLY (Skipped Project File).", "#27ae60")
                return 

            sheet_name = self.assembly_v.get()[:31]
            if not os.path.exists(proj_file):
                messagebox.showerror("Error", f"File {full_name}.xlsx not found!\nPlease ensure the file exists.")
                return

            if wb_p is None: wb_p = load_workbook(proj_file)
            
            if sheet_name not in wb_p.sheetnames:
                 ws_p = wb_p.create_sheet(sheet_name)
                 ws_p.append(["Sl. No", "Drawing Name", "Part Number", "Revision", "Total Drawings", "Date Approved", "Remarks"])
                 ws_p.cell(row=15, column=1, value="Drawing issued by:-")
            else:
                ws_p = wb_p[sheet_name]

            sig_row = None
            for r in range(1, ws_p.max_row + 50):
                found_sig = False
                for c in range(1, 6): 
                    val = ws_p.cell(row=r, column=c).value
                    if val and ("issued by" in str(val).lower() or "drawing issued" in str(val).lower()):
                        sig_row = r; found_sig = True; break
                if found_sig: break
            
            if not sig_row:
                sig_row = max(ws_p.max_row + 2, 4)
                ws_p.cell(row=sig_row, column=1, value="Drawing issued by:-")

            target_row = None; is_override = False; remarks_to_save = self.remark_v.get()

            for r in range(3, sig_row):
                existing_part = ws_p.cell(row=r, column=3).value
                if existing_part and str(existing_part).strip().upper() == part:
                    target_row = r; is_override = True; remarks_to_save = "Revised"; break
            
            if not target_row:
                last_data_row = 2
                start_scan = max(2, sig_row - 1)
                for r in range(start_scan, 2, -1):
                    val_draw = ws_p.cell(row=r, column=2).value
                    val_part = ws_p.cell(row=r, column=3).value
                    if (val_draw and str(val_draw).strip()) or (val_part and str(val_part).strip()):
                        last_data_row = r; break
                
                target_row = last_data_row + 1
                ws_p.insert_rows(target_row)
                
                next_row = target_row + 1
                found_sig_below = False
                for c in range(1, 6):
                    val_below = ws_p.cell(row=next_row, column=c).value
                    if val_below and "issued by" in str(val_below).lower():
                        found_sig_below = True; break
                
                if found_sig_below: ws_p.insert_rows(next_row) 

            if not is_override:
                if target_row == 3: sl_no = 1
                else:
                    try:
                        prev_sl_cell = ws_p.cell(row=target_row-1, column=1)
                        prev_sl = 0 if isinstance(prev_sl_cell, MergedCell) else prev_sl_cell.value
                        sl_no = int(prev_sl) + 1 if prev_sl and str(prev_sl).isdigit() else 1
                    except: sl_no = 1
            else:
                sl_no = ws_p.cell(row=target_row, column=1).value

            final_data = [sl_no, draw, part, rev, total, dt_obj, remarks_to_save]

            for col, val in enumerate(final_data, 1):
                cell = ws_p.cell(row=target_row, column=col)
                if isinstance(cell, MergedCell): continue 
                cell.value = val
                if col == 6: cell.number_format = 'DD/MM/YYYY'
                cell.alignment = LEFT_ALIGN if col in [2, 3] else CENTER_ALIGN
                cell.border = THIN_BORDER
            
            new_sig_row = None
            for r in range(1, ws_p.max_row + 10):
                for c in range(1, 6):
                    val = ws_p.cell(row=r, column=c).value
                    if val and "issued by" in str(val).lower():
                        new_sig_row = r; break
                if new_sig_row: break
            if not new_sig_row: new_sig_row = ws_p.max_row 

            dates_objects = []
            for r in range(3, new_sig_row):
                d_val = ws_p.cell(row=r, column=6).value
                parsed_date = None
                if isinstance(d_val, datetime): parsed_date = d_val.date()
                elif isinstance(d_val, date): parsed_date = d_val
                elif d_val:
                    d_str = str(d_val).strip()
                    try: parsed_date = datetime.strptime(d_str, '%d/%m/%Y').date()
                    except ValueError:
                        try: parsed_date = datetime.strptime(d_str, '%Y-%m-%d').date()
                        except ValueError: pass
                if parsed_date: dates_objects.append(parsed_date)
            
            latest_date_obj = max(dates_objects) if dates_objects else None

            for r in range(3, new_sig_row):
                d_val = ws_p.cell(row=r, column=6).value
                current_date_obj = None
                if isinstance(d_val, datetime): current_date_obj = d_val.date()
                elif isinstance(d_val, date): current_date_obj = d_val
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
                    if c == 7 and not is_latest: cell.value = None
                    cell.alignment = LEFT_ALIGN if c in [2, 3] else CENTER_ALIGN
                    if is_latest:
                        cell.fill = CREAM_YELLOW
                        cell.font = Font(bold=True)
                    else:
                        cell.fill = NO_FILL
                        cell.font = Font(bold=False)

            wb_p.save(proj_file)
            action_msg = "UPDATED (Override)" if is_override else "SAVED (New)"
            self.update_status(f"Success! Data '{part}' has been {action_msg} at Row {target_row}.", "#27ae60")

        except PermissionError:
             messagebox.showerror("File Error", f"Please close file {full_name}.xlsx before submitting!")
        except Exception as e:
            self.update_status(f"System Error: {str(e)}", COLOR_DANGER)
            print(e)

if __name__ == "__main__":
    app = DCDEApp()
    app.mainloop()