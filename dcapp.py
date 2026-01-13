import customtkinter as ctk
import psycopg2
import os
from tkinter import messagebox
from datetime import date, datetime
from tkcalendar import DateEntry
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.cell.cell import MergedCell 
from PIL import Image, ImageTk 

# --- APPEARANCE SETUP ---
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("dark-blue")

# --- CONFIGURATION ---
BASE_PATH = r"C:\Users\HP\Documents\[01] Document Control"
DB_CONFIG = {"dbname": "dcde", "user": "postgres", "password": "1234", "host": "localhost"}
MASTER_FILE = "Document Control.xlsx"
LOGO_FILENAME = "LMG Locomotive Logo.jpeg" 

# UPDATED: Project Names
PROJECTS = ["H10 TRC", "H10 BeraPit", "M10", "N10", "Wheel Press Machine"]
PROJ_MAP = {"H10 TRC": "H10", "H10 BeraPit": "H10", "M10": "M10", "N10": "N10", "Wheel Press Machine": "WPM"}

# UPDATED: Assembly names
ASSEMBLIES = ["Bogie", "Underframe", "Cabin", "Engine Hood", "Radiator Hood", "Muffler", "Gear Case", "Water Tank", "Battery Box", "Fuel Tank", "Sandbox"]
ENGINEER_LIST = ["Baskaran", "Sathish", "Harrison", "Hannan", "Gokul", "Vimal", "Ram", "Vishwa", "Bruno"]
REMARKS_LIST = ["New", "Revised", "-"]
BATCH_LIST = ["-", "1", "2", "N", "R"]
MASTER_HEADERS = ["Project", "Country", "Batch", "Main Assembly", "Drawing Name", "Part Number", "Revision", "Total Sheets", "Engineer", "Date Approved", "Remarks"]

# --- COLORS & FONTS ---
COLOR_PRIMARY = "#2C3E50"    
COLOR_ACCENT = "#3498DB"     
COLOR_SUCCESS = "#27AE60"    
COLOR_DANGER = "#C0392B"     
COLOR_WARNING = "#F39C12"    
COLOR_INFO = "#1ABC9C"       
COLOR_BG = "#ECF0F1"         
COLOR_CARD = "#FFFFFF"       
COLOR_TEXT = "#34495E"       

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
        
        self.geometry("1024x760") 
        self.minsize(950, 720) 
        
        self.configure(fg_color=COLOR_BG)
        
        # --- SET WINDOW ICON ---
        try:
            if os.path.exists(LOGO_FILENAME):
                self.window_icon = ImageTk.PhotoImage(file=LOGO_FILENAME)
                self.wm_iconphoto(False, self.window_icon)
        except Exception as e:
            print(f"Warning: Could not set window icon. {e}")

        # Grid layout for the main window
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) 
        
        self.setup_ui()

    def setup_ui(self):
        # --- 1. HEADER SECTION ---
        self.header_frame = ctk.CTkFrame(self, fg_color=COLOR_PRIMARY, corner_radius=0, height=70)
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
                pil_img = Image.open(LOGO_FILENAME)
                
                target_height = 40
                aspect_ratio = pil_img.width / pil_img.height
                target_width = int(target_height * aspect_ratio)
                
                self.header_logo = ctk.CTkImage(light_image=pil_img, 
                                                dark_image=pil_img, 
                                                size=(target_width, target_height))
                
                self.logo_label = ctk.CTkLabel(self.header_frame, text="", image=self.header_logo)
                self.logo_label.pack(side="right", padx=30)
            else:
                print(f"Info: Logo file '{LOGO_FILENAME}' not found. Skipping logo.")
        except Exception as e:
            print(f"Error loading logo: {e}")

        # --- 2. MAIN CONTENT CARD ---
        self.content_frame = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=15, border_width=1, border_color="#BDC3C7")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_columnconfigure((0, 1), weight=1) 

        # --- SECTION: PROJECT DETAILS ---
        self.create_section_header("1. PROJECT DETAILS", row=0)

        # Project Dropdown
        self.add_input_field(label="Project Name", row=1, col=0)
        self.proj_v = ctk.StringVar(value=PROJECTS[0])
        self.proj_drop = ctk.CTkOptionMenu(self.content_frame, values=PROJECTS, variable=self.proj_v, command=self.update_logic, 
                                           fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.proj_drop.grid(row=2, column=0, padx=20, pady=(5, 15), sticky="ew")

        # Batch Dropdown
        self.add_input_field(label="Batch Code", row=1, col=1)
        self.batch_v = ctk.StringVar(value="-")
        self.batch_drop = ctk.CTkOptionMenu(self.content_frame, values=BATCH_LIST, variable=self.batch_v,
                                            fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.batch_drop.grid(row=2, column=1, padx=20, pady=(5, 15), sticky="ew")

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
        self.part_ent.bind("<KeyRelease>", self.on_part_input)

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
        self.btn_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

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
        self.btn_submit.grid(row=0, column=0, padx=(0, 5), sticky="ew")

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
        self.btn_clear.grid(row=0, column=1, padx=(5, 5), sticky="ew")

        self.btn_open_folder = ctk.CTkButton(
            self.btn_frame,
            text="OPEN FOLDER",
            fg_color=COLOR_WARNING,
            hover_color="#D35400",
            height=50,
            corner_radius=8,
            font=("Segoe UI", 15, "bold"),
            command=self.open_folder
        )
        self.btn_open_folder.grid(row=0, column=2, padx=(5, 5), sticky="ew")
        
        self.btn_open_excel = ctk.CTkButton(
            self.btn_frame,
            text="OPEN EXCEL",
            fg_color=COLOR_INFO,
            hover_color="#16A085",
            height=50,
            corner_radius=8,
            font=("Segoe UI", 15, "bold"),
            command=self.open_project_file
        )
        self.btn_open_excel.grid(row=0, column=3, padx=(5, 0), sticky="ew")

        # --- 4. STATUS BAR ---
        self.status_bar = ctk.CTkFrame(self, height=30, fg_color="#BDC3C7", corner_radius=0)
        self.status_bar.grid(row=2, column=0, sticky="ew")
        self.status_label = ctk.CTkLabel(self.status_bar, text="System Ready", font=("Consolas", 11), text_color="#2C3E50")
        self.status_label.pack(side="left", padx=20)

    # --- UI HELPER FUNCTIONS ---
    def create_section_header(self, text, row):
        label = ctk.CTkLabel(self.content_frame, text=text, font=FONT_SECTION, text_color=COLOR_PRIMARY, anchor="w")
        label.grid(row=row, column=0, columnspan=2, padx=20, pady=(15, 5), sticky="w")
        
        line = ctk.CTkFrame(self.content_frame, height=2, fg_color="#E0E0E0")
        line.grid(row=row+1, column=0, columnspan=2, padx=20, pady=(0, 10), sticky="ew")

    def add_input_field(self, label, row, col):
        ctk.CTkLabel(self.content_frame, text=label, font=FONT_LABEL, text_color=COLOR_TEXT).grid(row=row, column=col, padx=20, pady=(5, 0), sticky="w")

    def update_status(self, message, color=COLOR_TEXT):
        self.status_label.configure(text=f"STATUS: {message}", text_color=color)
        self.update_idletasks() # Critical for preventing "Not Responding"

    def open_folder(self):
        try:
            if os.path.exists(BASE_PATH):
                os.startfile(BASE_PATH)
                self.update_status(f"Opened folder: {BASE_PATH}", COLOR_TEXT)
            else:
                messagebox.showerror("Folder Error", f"Folder not found:\n{BASE_PATH}")
                self.update_status("Error: Folder not found", COLOR_DANGER)
        except Exception as e:
            self.update_status(f"Error opening folder: {str(e)}", COLOR_DANGER)
            print(f"Error opening folder: {e}")

    def open_project_file(self):
        full_name = self.proj_v.get()
        file_path = os.path.join(BASE_PATH, f"{full_name}.xlsx")
        
        try:
            if os.path.exists(file_path):
                os.startfile(file_path)
                self.update_status(f"Opened file: {full_name}.xlsx", COLOR_TEXT)
            else:
                messagebox.showerror("File Error", f"File not found:\n{file_path}\nPlease submit data first to create the file.")
                self.update_status("Error: Excel file not found", COLOR_DANGER)
        except Exception as e:
             self.update_status(f"Error opening file: {str(e)}", COLOR_DANGER)

    # --- LOGIC FUNCTIONS ---
    def to_uppercase(self, event, widget):
        ignored = ["Control_L", "Control_R", "Shift_L", "Shift_R", "Caps_Lock", "Left", "Right", "Up", "Down", "Home", "End"]
        if event.keysym in ignored or (event.state & 0x0004): return
        pos = widget.index(ctk.INSERT)
        val = widget.get().upper()
        if widget.get() != val:
            widget.delete(0, 'end'); widget.insert(0, val); widget.icursor(pos)

    def on_part_input(self, event):
        self.to_uppercase(event, self.part_ent)
        val = self.part_ent.get()
        current_proj = self.proj_v.get()
        if current_proj != "Wheel Press Machine" and "(N)" in val:
            self.batch_v.set("N")

    def auto_remark_logic(self, event):
        val = self.rev_ent.get()
        if val.isdigit(): self.remark_v.set("Revised" if int(val) >= 1 else "New")

    def update_logic(self, choice):
        self.batch_drop.configure(state="normal")
        
        if choice == "Wheel Press Machine":
            self.batch_v.set("-")
            self.batch_drop.configure(state="disabled") 
            self.assembly_drop.configure(values=["Wheel Press"])
            self.assembly_v.set("Wheel Press")
            self.eng_drop.configure(values=["Baskaran"])
            self.eng_v.set("Baskaran")
            
        elif choice == "H10 TRC":
            self.batch_v.set("2")
            self.assembly_drop.configure(values=ASSEMBLIES)
            if self.assembly_v.get() == "Wheel Press": self.assembly_v.set(ASSEMBLIES[0])
            self.eng_drop.configure(values=ENGINEER_LIST)
            
        else:
            self.assembly_drop.configure(values=ASSEMBLIES)
            if self.assembly_v.get() == "Wheel Press": self.assembly_v.set(ASSEMBLIES[0])
            self.eng_drop.configure(values=ENGINEER_LIST)

    def set_today(self): self.date_picker.set_date(date.today())

    def clear_all(self):
        if self.proj_v.get() == "Wheel Press Machine":
            self.batch_v.set("-")
        elif self.proj_v.get() == "H10 TRC":
             self.batch_v.set("2")
        else:
            self.batch_v.set("-")
            
        for w in [self.draw_ent, self.part_ent, self.rev_ent, self.total_ent]: w.delete(0, 'end')

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
        
        batch_val = self.batch_v.get() 
        batch_val_excel = int(batch_val) if batch_val.isdigit() else batch_val
        
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

        sql_data = [short_proj, "Tanzania" if full_name == "H10 TRC" else "Malaysia", batch_val, self.assembly_v.get(), draw, part, rev, total, self.eng_v.get(), dt_sql, self.remark_v.get()]
        excel_master_data = [short_proj, "Tanzania" if full_name == "H10 TRC" else "Malaysia", batch_val_excel, self.assembly_v.get(), draw, part, rev, total, self.eng_v.get(), dt_obj, self.remark_v.get()]

        try:
            self.update_status("Checking Duplicates...", COLOR_WARNING)
            
            # ----------------------------------------------------
            #  STEP 1: CHECK MASTER FILE DUPLICATE (Strict Lock)
            # ----------------------------------------------------
            if os.path.exists(master_file_path):
                try:
                    wb_m_check = load_workbook(master_file_path, read_only=True)
                    ws_m_check = wb_m_check.active
                    
                    # Iterate rows safely (Limit to 5000 rows for speed)
                    scan_limit_m = min(ws_m_check.max_row + 1, 5000)
                    for row in ws_m_check.iter_rows(min_row=2, max_row=scan_limit_m, values_only=True):
                        if not row or len(row) < 7: continue
                        
                        m_proj = row[0] # Project Name (Col 1)
                        m_part = row[5] # Part Number (Col 6)
                        m_rev = row[6]  # Revision (Col 7)
                        
                        if m_proj and m_part:
                            if str(m_proj).strip() == short_proj and \
                               str(m_part).strip().upper() == part and \
                               str(m_rev).strip() == str(rev):
                                wb_m_check.close()
                                # --- HARD STOP HERE ---
                                messagebox.showerror("Duplicate Error", f"Data already exists in MASTER FILE!\n\nProject: {short_proj}\nPart: {part}\nRev: {rev}\n\nCannot add duplicate data.")
                                self.update_status("Error: Duplicate found in Master File.", COLOR_DANGER)
                                return 
                    wb_m_check.close()
                except Exception as e:
                    print(f"Master check error: {e}")

            # ----------------------------------------------------
            #  STEP 2: CHECK SUB FILE DUPLICATE (Skip Logic)
            # ----------------------------------------------------
            duplicate_in_sub = False
            target_sheet_name = self.assembly_v.get()[:31]
            skip_sub_write = False

            if os.path.exists(proj_file):
                try:
                    wb_temp = load_workbook(proj_file, read_only=True)
                    actual_sheet = None
                    for s in wb_temp.sheetnames:
                        if s.strip().lower() == target_sheet_name.lower(): actual_sheet = s; break
                    
                    if actual_sheet:
                        ws_temp = wb_temp[actual_sheet]
                        scan_limit = min(ws_temp.max_row + 10, 5000) 
                        data_start = 3
                        # Quick header check
                        for r in range(1, 11):
                            val = ws_temp.cell(row=r, column=3).value
                            if val and "part number" in str(val).lower(): data_start = r + 1; break
                        
                        sig = scan_limit
                        for r in range(data_start, scan_limit):
                            val = ws_temp.cell(row=r, column=1).value
                            if val and "issued by" in str(val).lower(): sig = r; break
                        
                        for r in range(data_start, sig):
                            e_part = ws_temp.cell(row=r, column=3).value
                            e_rev = ws_temp.cell(row=r, column=4).value
                            if str(e_part).strip().upper() == part and str(e_rev).strip() == str(rev):
                                duplicate_in_sub = True; break
                    wb_temp.close()
                except: pass

            if duplicate_in_sub:
                ans = messagebox.askyesno("Duplicate in Sub File", f"Part {part} Rev {rev} exists in Sub File (but NOT in Master).\n\nDo you want to add to MASTER LIST & SQL only?")
                if not ans: 
                    self.update_status("Cancelled.", COLOR_TEXT)
                    return
                skip_sub_write = True

            # ----------------------------------------------------
            #  STEP 3: WRITE TO SUB FILE (PROJECT FILE) - Gatekeeper
            # ----------------------------------------------------
            if not skip_sub_write:
                self.update_status("Saving to Sub-File...", COLOR_WARNING)
                try:
                    wb_p = None
                    if not os.path.exists(proj_file):
                        wb_p = Workbook()
                        ws_p = wb_p.active; ws_p.title = target_sheet_name
                        ws_p.append(["Document Control"])
                        ws_p.append(["Sl. No", "Drawing Name", "Part Number", "Revision", "Total Drawings", "Date Approved", "Remarks"])
                        ws_p.cell(row=15, column=1, value="Drawing issued by:-")
                        data_start_row = 3
                    else:
                        wb_p = load_workbook(proj_file)
                        actual_sheet_name = None
                        for s in wb_p.sheetnames:
                            if s.strip().lower() == target_sheet_name.lower(): actual_sheet_name = s; break
                        
                        data_start_row = 3
                        if actual_sheet_name:
                            ws_p = wb_p[actual_sheet_name]
                            for r in range(1, 11):
                                val = ws_p.cell(row=r, column=3).value
                                if val and "part number" in str(val).lower(): data_start_row = r + 1; break
                        else:
                            ws_p = wb_p.create_sheet(target_sheet_name)
                            ws_p.append(["Document Control"])
                            ws_p.append(["Sl. No", "Drawing Name", "Part Number", "Revision", "Total Drawings", "Date Approved", "Remarks"])
                            ws_p.cell(row=15, column=1, value="Drawing issued by:-")
                            data_start_row = 3

                    # Optimized Sig Detection
                    sig_row = None
                    limit_rows = min(ws_p.max_row + 50, 5000)
                    for r in range(data_start_row, limit_rows):
                        val = ws_p.cell(row=r, column=1).value
                        if val and "issued by" in str(val).lower(): sig_row = r; break
                    if not sig_row: 
                        sig_row = max(ws_p.max_row + 2, data_start_row + 1)
                        ws_p.cell(row=sig_row, column=1, value="Drawing issued by:-")

                    target_row = None
                    is_override = False
                    remarks_to_save = self.remark_v.get()

                    # Find existing row (Override Check)
                    for r in range(data_start_row, sig_row):
                        existing_part = ws_p.cell(row=r, column=3).value
                        if existing_part and str(existing_part).strip().upper() == part:
                            target_row = r
                            is_override = True
                            remarks_to_save = "Revised"
                            break
                    
                    if not target_row:
                        last_data_row = data_start_row - 1
                        start_scan = max(data_start_row, sig_row - 1)
                        for r in range(start_scan, data_start_row - 1, -1):
                            val_draw = ws_p.cell(row=r, column=2).value
                            if val_draw: last_data_row = r; break
                        
                        target_row = last_data_row + 1
                        ws_p.insert_rows(target_row)
                        
                        # --- OPTIMIZED ROW HEIGHT FIX (Limit Scan) ---
                        height_scan_limit = min(ws_p.max_row, sig_row + 20)
                        for r in range(height_scan_limit, target_row, -1):
                            if (r-1) in ws_p.row_dimensions:
                                ws_p.row_dimensions[r].height = ws_p.row_dimensions[r-1].height
                        ws_p.row_dimensions[target_row].height = None

                    if not is_override:
                        sl_no = target_row - data_start_row + 1
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

                    # --- Sorting Logic (Optimized) ---
                    self.update_status("Sorting Sub-File...", COLOR_INFO)
                    sig_row_current = None
                    limit_sort = min(ws_p.max_row + 10, 5000)
                    for r in range(data_start_row, limit_sort):
                        val = ws_p.cell(row=r, column=1).value
                        if val and "issued by" in str(val).lower(): sig_row_current = r; break
                    if not sig_row_current: sig_row_current = ws_p.max_row + 1
                    sort_end_row = sig_row_current - 1

                    data_matrix = []
                    if sort_end_row >= data_start_row:
                        for r in range(data_start_row, sort_end_row + 1):
                            row_data = []
                            has_data = False
                            for c in range(2, 8):
                                val = ws_p.cell(row=r, column=c).value
                                row_data.append(val)
                                if val is not None: has_data = True
                            if has_data: data_matrix.append(row_data)
                        
                        data_matrix.sort(key=lambda x: str(x[1]).strip().upper() if x[1] else "")

                        for i, row_data in enumerate(data_matrix):
                            current_r = data_start_row + i
                            sl_cell = ws_p.cell(row=current_r, column=1)
                            if not isinstance(sl_cell, MergedCell):
                                sl_cell.value = i + 1
                                sl_cell.alignment = CENTER_ALIGN
                                sl_cell.border = THIN_BORDER
                            
                            for idx, val in enumerate(row_data):
                                c = idx + 2
                                cell = ws_p.cell(row=current_r, column=c)
                                if isinstance(cell, MergedCell): continue
                                cell.value = val
                                if c == 6: cell.number_format = 'DD/MM/YYYY'
                                cell.alignment = LEFT_ALIGN if c in [2, 3] else CENTER_ALIGN
                                cell.border = THIN_BORDER

                    # --- Date Highlight ---
                    dates_objects = []
                    for r in range(data_start_row, sig_row_current):
                        d_val = ws_p.cell(row=r, column=6).value
                        parsed_date = None
                        if isinstance(d_val, datetime): parsed_date = d_val.date()
                        elif isinstance(d_val, date): parsed_date = d_val
                        elif d_val:
                            try: parsed_date = datetime.strptime(str(d_val).strip(), '%d/%m/%Y').date()
                            except: pass
                        if parsed_date: dates_objects.append(parsed_date)
                    
                    latest_date_obj = max(dates_objects) if dates_objects else None
                    for r in range(data_start_row, sig_row_current):
                        d_val = ws_p.cell(row=r, column=6).value
                        c_date = None
                        if isinstance(d_val, datetime): c_date = d_val.date()
                        elif isinstance(d_val, date): c_date = d_val
                        elif d_val:
                            try: c_date = datetime.strptime(str(d_val).strip(), '%d/%m/%Y').date()
                            except: pass
                        is_latest = (c_date == latest_date_obj) and (latest_date_obj is not None)
                        for c in range(1, 8):
                            cell = ws_p.cell(row=r, column=c)
                            if isinstance(cell, MergedCell): continue
                            if c == 7 and not is_latest: cell.value = None
                            if is_latest: cell.fill = CREAM_YELLOW; cell.font = Font(bold=True)
                            else: cell.fill = NO_FILL; cell.font = Font(bold=False)

                    wb_p.save(proj_file)
                
                except PermissionError:
                    messagebox.showerror("File Error", f"Cannot write to {full_name}.xlsx!\nPlease close the file.")
                    self.update_status("Aborted: Sub File Open. Master/SQL not updated.", COLOR_DANGER)
                    return # HARD STOP
                except Exception as e:
                    self.update_status(f"Sub File Error: {str(e)}", COLOR_DANGER)
                    print(e)
                    return # HARD STOP

            # ----------------------------------------------------
            #  STEP 4: WRITE TO MASTER FILE
            # ----------------------------------------------------
            self.update_status("Updating Master File...", COLOR_INFO)
            try:
                if not os.path.exists(master_file_path):
                    wb_m = Workbook(); ws_m = wb_m.active; ws_m.title = "MasterList"
                    ws_m.append(MASTER_HEADERS); wb_m.save(master_file_path)
                
                wb_m = load_workbook(master_file_path)
                ws_m = wb_m.active
                ws_m.append(excel_master_data)
                
                last_row_m = ws_m.max_row
                for col_idx in range(1, len(excel_master_data) + 1):
                    cell_m = ws_m.cell(row=last_row_m, column=col_idx)
                    if col_idx == 10: cell_m.number_format = 'DD/MM/YYYY'
                    cell_m.alignment = LEFT_ALIGN if col_idx in [4, 5, 6] else CENTER_ALIGN
                    cell_m.border = THIN_BORDER
                wb_m.save(master_file_path)
            except PermissionError:
                messagebox.showwarning("Master File Locked", "Sub File Saved.\nMaster File LOCKED (Update Manually).")
            except Exception as e:
                print(f"Master Error: {e}")

            # ----------------------------------------------------
            #  STEP 5: WRITE TO SQL
            # ----------------------------------------------------
            try:
                self.update_status("Updating SQL...", COLOR_INFO)
                conn = psycopg2.connect(**DB_CONFIG)
                cur = conn.cursor()
                cur.execute("INSERT INTO project_data (project_name, country, batch, main_assembly, drawing_name, part_number, revision, total_sheets, engineer, date_approved, remarks) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", sql_data)
                conn.commit(); cur.close(); conn.close()
            except Exception as db_err:
                self.update_status(f"SQL Error: {str(db_err)}", COLOR_DANGER)
                print(f"Database Error: {db_err}")

            self.update_status(f"Success! '{part}' saved.", "#27ae60")

        except Exception as e:
            self.update_status(f"System Error: {str(e)}", COLOR_DANGER)
            print(e)

if __name__ == "__main__":
    app = DCDEApp()
    app.mainloop()