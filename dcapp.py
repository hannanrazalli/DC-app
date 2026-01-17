import customtkinter as ctk
import psycopg2
import os
import threading
import shutil  # Dari Kod 2
import time    # Dari Kod 2
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

# --- CONFIGURATION (Kombinasi Kod 1 & Kod 2) ---
BASE_PATH = r"C:\Users\HP\Documents\[01] Document Control"
DB_CONFIG = {"dbname": "dcde", "user": "postgres", "password": "1234", "host": "localhost"}
MASTER_FILE = "Document Control.xlsx"
LOGO_FILENAME = "LMG Locomotive Logo.jpeg" 

# CONFIGURATION DARI KOD 2
SERVER_PATH = r"Y:\[04] ENGINEERING TEAM\[98] DOCUMENT CONTROL"
FILE_MAPPING = {
    "Document Control.xlsx": "All Projects Drawings Data.xlsx",
    "H10 Berapit.xlsx": "H10 BeraPit Drawing List.xlsx",
    "H10 TRC.xlsx": "H10 TRC Drawing List.xlsx",
    "M10.xlsx": "M10(N) Drawing List.xlsx",
    "N10.xlsx": "N10(N) Drawing List.xlsx",
    "Wheel Press Machine.xlsx": "Wheel Press Machine Drawing List.xlsx",
    "G10.xlsx": "G10 Drawing List.xlsx"
}

# PROJECT CONFIGURATION
PROJECTS = ["H10 TRC", "H10 BeraPit", "M10", "N10", "G10", "Wheel Press Machine"]

PROJ_MAP = {
    "H10 TRC": "H10", 
    "H10 BeraPit": "H10", 
    "M10": "M10", 
    "N10": "N10", 
    "G10": "G10 Drawing List", 
    "Wheel Press Machine": "WPM"
}

# Reverse Map for Auto-fill (Short Code -> Full Name)
REVERSE_PROJ_MAP = {}
for k, v in PROJ_MAP.items():
    if v not in REVERSE_PROJ_MAP: 
        REVERSE_PROJ_MAP[v] = k

ASSEMBLIES = ["Bogie", "Underframe", "Cabin", "Engine Hood", "Radiator Hood", "Muffler", "Gear Case", "Water Tank", "Battery Box", "Fuel Tank", "Sandbox"]
ENGINEER_LIST = ["Baskaran", "Sathish", "Harrison", "Hannan", "Gokul", "Vimal", "Ram", "Vishwa", "Bruno"]
REMARKS_LIST = ["New", "Revised", "Re-Release", "-"]
# BATCH_LIST dari kod anda
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
FONT_FEEDBACK = ("Segoe UI", 16, "bold") 

# Excel Styles
CREAM_YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

class DCDEApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Design Engineering Data Entry")
        self.geometry("1024x820") 
        self.minsize(950, 750) 
        self.configure(fg_color=COLOR_BG)
        
        # Window Icon
        try:
            if os.path.exists(LOGO_FILENAME):
                self.window_icon = ImageTk.PhotoImage(file=LOGO_FILENAME)
                self.wm_iconphoto(False, self.window_icon)
        except Exception as e:
            print(f"Warning: Could not set window icon. {e}")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) 
        
        self.setup_ui()

    def setup_ui(self):
        # Header
        self.header_frame = ctk.CTkFrame(self, fg_color=COLOR_PRIMARY, corner_radius=0, height=85)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        self.header_frame.grid_propagate(False)

        self.title_label = ctk.CTkLabel(self.header_frame, text="ENGINEERING DOCUMENT CONTROL", font=FONT_HEADER, text_color="white")
        self.title_label.pack(side="left", padx=25, pady=15)
        
        self.subtitle_label = ctk.CTkLabel(self.header_frame, text="|  Data Entry System", font=("Segoe UI", 14), text_color="#BDC3C7")
        self.subtitle_label.pack(side="left", pady=20)

        # Right Side Container (Publish + Logo)
        self.right_side_container = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.right_side_container.pack(side="right", padx=30)

        # Publish UI Block
        self.publish_ui_frame = ctk.CTkFrame(self.right_side_container, fg_color="transparent")
        self.publish_ui_frame.pack(side="left", padx=(0, 20))

        # 1. Publish Button (Di sebelah kiri logo)
        self.btn_publish = ctk.CTkButton(self.publish_ui_frame, text="PUBLISH TO SERVER", 
                                       fg_color=COLOR_INFO, hover_color="#16A085", 
                                       font=("Segoe UI", 12, "bold"), height=32,
                                       command=self.publish_to_server_thread)
        self.btn_publish.pack(pady=(0, 2))

        # 2. Last Sync Label (Bawah button)
        self.lbl_last_sync = ctk.CTkLabel(self.publish_ui_frame, text="Last Sync: Never", 
                                        font=("Segoe UI", 10), text_color="#BDC3C7")
        self.lbl_last_sync.pack()

        # Logo
        try:
            if os.path.exists(LOGO_FILENAME):
                pil_img = Image.open(LOGO_FILENAME)
                target_height = 40
                aspect_ratio = pil_img.width / pil_img.height
                target_width = int(target_height * aspect_ratio)
                self.header_logo = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(target_width, target_height))
                self.logo_label = ctk.CTkLabel(self.right_side_container, text="", image=self.header_logo)
                self.logo_label.pack(side="right")
        except Exception as e:
            print(f"Error loading logo: {e}")

        # Content Card
        self.content_frame = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=15, border_width=1, border_color="#BDC3C7")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_columnconfigure((0, 1), weight=1) 

        # 1. Project Details
        self.create_section_header("1. PROJECT DETAILS", row=0)

        self.add_input_field(label="Project Name", row=1, col=0)
        self.proj_v = ctk.StringVar(value=PROJECTS[0])
        self.proj_drop = ctk.CTkOptionMenu(self.content_frame, values=PROJECTS, variable=self.proj_v, command=self.update_logic, 
                                           fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.proj_drop.grid(row=2, column=0, padx=20, pady=(5, 15), sticky="ew")

        self.add_input_field(label="Batch Code", row=1, col=1)
        self.batch_v = ctk.StringVar(value="-")
        self.batch_drop = ctk.CTkOptionMenu(self.content_frame, values=BATCH_LIST, variable=self.batch_v,
                                            fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.batch_drop.grid(row=2, column=1, padx=20, pady=(5, 15), sticky="ew")

        # 2. Technical Data
        self.create_section_header("2. DRAWING & TECHNICAL INFORMATION", row=3)

        self.add_input_field(label="Main Assembly", row=4, col=0)
        self.assembly_v = ctk.StringVar(value=ASSEMBLIES[0])
        self.assembly_drop = ctk.CTkOptionMenu(self.content_frame, values=ASSEMBLIES, variable=self.assembly_v, 
                                                fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.assembly_drop.grid(row=5, column=0, padx=20, pady=(5, 10), sticky="ew")

        self.add_input_field(label="Drawing Name", row=4, col=1)
        self.draw_ent = ctk.CTkEntry(self.content_frame, font=FONT_INPUT, height=32, placeholder_text="Full drawing title")
        self.draw_ent.grid(row=5, column=1, padx=20, pady=(5, 10), sticky="ew")
        self.draw_ent.bind("<KeyRelease>", lambda e: self.to_uppercase(e, self.draw_ent))

        self.add_input_field(label="Part Number", row=6, col=0)
        self.part_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.part_frame.grid(row=7, column=0, padx=20, pady=(5, 10), sticky="ew")
        
        self.part_ent = ctk.CTkEntry(self.part_frame, font=FONT_INPUT, height=32, placeholder_text="e.g. H10-100-001")
        self.part_ent.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.part_ent.bind("<KeyRelease>", self.on_part_input)
        
        self.btn_check_part = ctk.CTkButton(self.part_frame, text="Search", width=60, height=32, 
                                            fg_color="#7f8c8d", hover_color="#95a5a6", 
                                            command=self.check_part_existence_thread)
        self.btn_check_part.pack(side="right")

        self.sub_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.sub_frame.grid(row=6, column=1, rowspan=2, padx=20, pady=0, sticky="ew")
        self.sub_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(self.sub_frame, text="Revision", font=FONT_LABEL, text_color=COLOR_TEXT).grid(row=0, column=0, sticky="w")
        self.rev_ent = ctk.CTkEntry(self.sub_frame, font=FONT_INPUT, height=32, placeholder_text="0")
        self.rev_ent.grid(row=1, column=0, padx=(0, 10), pady=(5, 10), sticky="ew")
        self.rev_ent.bind("<KeyRelease>", self.auto_remark_logic)

        ctk.CTkLabel(self.sub_frame, text="Total Sheets", font=FONT_LABEL, text_color=COLOR_TEXT).grid(row=0, column=1, sticky="w")
        self.total_ent = ctk.CTkEntry(self.sub_frame, font=FONT_INPUT, height=32, placeholder_text="1")
        self.total_ent.grid(row=1, column=1, padx=(10, 0), pady=(5, 10), sticky="ew")

        # 3. Approval
        self.create_section_header("3. APPROVAL & STATUS", row=8)

        self.add_input_field(label="Responsible Engineer", row=9, col=0)
        self.eng_v = ctk.StringVar(value=ENGINEER_LIST[0])
        self.eng_drop = ctk.CTkOptionMenu(self.content_frame, values=ENGINEER_LIST, variable=self.eng_v, 
                                          fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.eng_drop.grid(row=10, column=0, padx=20, pady=(5, 10), sticky="ew")

        self.add_input_field(label="Date Approved", row=9, col=1)
        self.date_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.date_frame.grid(row=10, column=1, padx=20, pady=(5, 10), sticky="w")
        
        self.date_picker = DateEntry(self.date_frame, width=20, background=COLOR_PRIMARY, foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', font=("Arial", 11))
        self.date_picker.pack(side="left", padx=(0, 15), ipady=3)
        
        self.btn_today = ctk.CTkButton(self.date_frame, text="Set Today", width=100, height=30, fg_color="#95a5a6", hover_color="#7f8c8d", command=self.set_today)
        self.btn_today.pack(side="left")

        self.add_input_field(label="Remarks", row=11, col=0)
        self.remark_v = ctk.StringVar(value="New")
        self.remark_drop = ctk.CTkOptionMenu(self.content_frame, values=REMARKS_LIST, variable=self.remark_v, 
                                             fg_color=COLOR_PRIMARY, button_color=COLOR_ACCENT, font=FONT_INPUT, height=32)
        self.remark_drop.grid(row=12, column=0, padx=20, pady=(5, 15), sticky="ew")

        # Feedback Banner
        self.feedback_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent", height=50)
        self.feedback_frame.grid(row=13, column=0, columnspan=2, pady=(10, 5), padx=20, sticky="ew")
        
        self.lbl_feedback = ctk.CTkLabel(self.feedback_frame, text="", font=FONT_FEEDBACK, text_color="white", corner_radius=6, fg_color="transparent")
        self.lbl_feedback.pack(fill="both", ipady=10)

        # Action Buttons
        self.btn_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.btn_frame.grid(row=14, column=0, columnspan=2, pady=(10, 20), padx=20, sticky="ew")
        self.btn_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.btn_submit = ctk.CTkButton(self.btn_frame, text="SUBMIT DATA", fg_color=COLOR_SUCCESS, hover_color="#219150", height=50, corner_radius=8, font=("Segoe UI", 15, "bold"), command=self.submit)
        self.btn_submit.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.btn_clear = ctk.CTkButton(self.btn_frame, text="CLEAR FORM", fg_color=COLOR_DANGER, hover_color="#A93226", height=50, corner_radius=8, font=("Segoe UI", 15, "bold"), command=self.clear_all)
        self.btn_clear.grid(row=0, column=1, padx=(5, 5), sticky="ew")

        self.btn_open_folder = ctk.CTkButton(self.btn_frame, text="OPEN FOLDER", fg_color=COLOR_WARNING, hover_color="#D35400", height=50, corner_radius=8, font=("Segoe UI", 15, "bold"), command=self.open_folder)
        self.btn_open_folder.grid(row=0, column=2, padx=(5, 5), sticky="ew")
        
        self.btn_open_excel = ctk.CTkButton(self.btn_frame, text="OPEN EXCEL", fg_color=COLOR_INFO, hover_color="#16A085", height=50, corner_radius=8, font=("Segoe UI", 15, "bold"), command=self.open_project_file)
        self.btn_open_excel.grid(row=0, column=3, padx=(5, 0), sticky="ew")

        # Status Bar
        self.status_bar = ctk.CTkFrame(self, height=30, fg_color="#BDC3C7", corner_radius=0)
        self.status_bar.grid(row=2, column=0, sticky="ew")
        self.status_label = ctk.CTkLabel(self.status_bar, text="System Ready", font=("Consolas", 11), text_color="#2C3E50")
        self.status_label.pack(side="left", padx=20)

    # --- INTEGRASI MEKANISME KOD 2 ---
    def publish_to_server_thread(self):
        """Memulakan proses publish dalam thread berasingan supaya GUI tidak hang."""
        self.btn_publish.configure(state="disabled", text="Publishing...")
        self.update_status("Publishing files to server...", COLOR_WARNING)
        threading.Thread(target=self._run_publish_logic, daemon=True).start()

    def _run_publish_logic(self):
        """Mekanisme asal dari Kod 2."""
        print("--- START PUBLISH (DIRECT COPY MODE) ---")
        
        if not os.path.exists(SERVER_PATH):
            self.after(0, lambda: messagebox.showerror("Error", f"Server path not found: {SERVER_PATH}\nPlease check network."))
            self.after(0, lambda: self.btn_publish.configure(state="normal", text="PUBLISH TO SERVER"))
            return

        success_count = 0
        fail_count = 0

        for local_file, server_file in FILE_MAPPING.items():
            local_full_path = os.path.join(BASE_PATH, local_file)
            server_full_path = os.path.join(SERVER_PATH, server_file)

            # Check local file existence
            if not os.path.exists(local_full_path):
                fail_count += 1
                continue

            # Check for "File Lock"
            file_is_locked = False
            if os.path.exists(server_full_path):
                try:
                    os.rename(server_full_path, server_full_path)
                except OSError:
                    file_is_locked = True
                    fail_count += 1
                    continue

            # Process Copy
            if not file_is_locked:
                try:
                    shutil.copy2(local_full_path, server_full_path)
                    success_count += 1
                except Exception as e:
                    print(f"Copy failed: {e}")
                    fail_count += 1

        # Update UI after finish
        self.after(0, self._finalize_publish, success_count, fail_count)

    def _finalize_publish(self, success, fail):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.lbl_last_sync.configure(text=f"Last Sync: {now}")
        self.btn_publish.configure(state="normal", text="PUBLISH TO SERVER")
        self.update_status(f"Publish Finished. Success: {success} | Failed: {fail}", COLOR_SUCCESS)
        self.trigger_feedback("success", "Publish Completed!")
        
        if fail > 0:
            messagebox.showwarning("Sync Result", f"Finished with issues.\nSuccess: {success}\nFailed/Locked: {fail}")
        else:
            messagebox.showinfo("Sync Success", f"All files successfully published to server.")

    # --- HELPERS (Dikekalkan 100%) ---
    def create_section_header(self, text, row):
        label = ctk.CTkLabel(self.content_frame, text=text, font=FONT_SECTION, text_color=COLOR_PRIMARY, anchor="w")
        label.grid(row=row, column=0, columnspan=2, padx=20, pady=(15, 5), sticky="w")
        line = ctk.CTkFrame(self.content_frame, height=2, fg_color="#E0E0E0")
        line.grid(row=row+1, column=0, columnspan=2, padx=20, pady=(0, 10), sticky="ew")

    def add_input_field(self, label, row, col):
        ctk.CTkLabel(self.content_frame, text=label, font=FONT_LABEL, text_color=COLOR_TEXT).grid(row=row, column=col, padx=20, pady=(5, 0), sticky="w")

    def update_status(self, message, color=COLOR_TEXT):
        self.after(0, lambda: self.status_label.configure(text=f"STATUS: {message}", text_color=color))

    def trigger_feedback(self, type="success", message="Saved Successfully"):
        if type == "success":
            color = COLOR_SUCCESS
            icon = "✔"
        elif type == "error":
            color = COLOR_DANGER
            icon = "✘"
        else:
            color = COLOR_INFO
            icon = "ℹ"
        self.lbl_feedback.configure(text=f"{icon}  {message}", fg_color=color)
        self.after(5000, lambda: self.lbl_feedback.configure(text="", fg_color="transparent"))

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

    def to_uppercase(self, event, widget):
        ignored = ["Control_L", "Control_R", "Shift_L", "Shift_R", "Caps_Lock", "Left", "Right", "Up", "Down", "Home", "End"]
        if event.keysym in ignored or (event.state & 0x0004): return
        pos = widget.index(ctk.INSERT)
        val = widget.get().upper()
        if widget.get() != val:
            widget.delete(0, 'end')
            widget.insert(0, val)
            widget.icursor(pos)

    def on_part_input(self, event):
        self.to_uppercase(event, self.part_ent)
        val = self.part_ent.get()
        current_proj = self.proj_v.get()
        if current_proj != "Wheel Press Machine" and "(N)" in val:
            self.batch_v.set("N")

    def auto_remark_logic(self, event):
        val = self.rev_ent.get()
        if val.isdigit():
            self.remark_v.set("Revised" if int(val) >= 1 else "New")

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
            if self.assembly_v.get() == "Wheel Press":
                self.assembly_v.set(ASSEMBLIES[0])
            self.eng_drop.configure(values=ENGINEER_LIST)
        else:
            self.assembly_drop.configure(values=ASSEMBLIES)
            if self.assembly_v.get() == "Wheel Press":
                self.assembly_v.set(ASSEMBLIES[0])
            self.eng_drop.configure(values=ENGINEER_LIST)

    def set_today(self):
        self.date_picker.set_date(date.today())

    def clear_all(self):
        if self.proj_v.get() == "Wheel Press Machine":
            self.batch_v.set("-")
        elif self.proj_v.get() == "H10 TRC":
            self.batch_v.set("2")
        else:
            self.batch_v.set("-")
        
        for w in [self.draw_ent, self.part_ent, self.rev_ent, self.total_ent]:
            w.delete(0, 'end')
        self.lbl_feedback.configure(text="", fg_color="transparent") 

    # --- SEARCH LOGIC (Dikekalkan 100%) ---
    def check_part_existence_thread(self):
        self.btn_check_part.configure(state="disabled", text="...")
        self.trigger_feedback("info", "Searching All Projects...")
        self.update_status("Searching globally...", COLOR_WARNING)
        part = self.part_ent.get().strip().upper()
        threading.Thread(target=self._run_global_check, args=(part,), daemon=True).start()

    def _run_global_check(self, part):
        if not part:
            self.after(0, lambda: self._post_check([]))
            return
        matches = []
        master_file_path = os.path.join(BASE_PATH, MASTER_FILE)
        
        try:
            if os.path.exists(master_file_path):
                wb = load_workbook(master_file_path, read_only=True)
                ws = wb.active
                for row in ws.iter_rows(min_row=2, max_row=8000, values_only=True): 
                    if not row or len(row) < 8:
                        continue
                    m_proj, m_batch, m_assy, m_draw, m_part, m_rev, m_total = row[0], row[2], row[3], row[4], row[5], row[6], row[7]
                    if str(m_part).strip().upper() == part:
                        matches.append({
                            'source': 'MASTER LIST', 
                            'project': m_proj, 
                            'batch': m_batch, 
                            'assembly': m_assy,
                            'draw': m_draw, 
                            'rev': m_rev, 
                            'total': m_total, 
                            'part': m_part
                        })
                wb.close()

            for ui_proj_name in PROJECTS:
                full_proj_name = ui_proj_name 
                short_proj_code = PROJ_MAP.get(ui_proj_name)
                proj_file = os.path.join(BASE_PATH, f"{full_proj_name}.xlsx")
                if os.path.exists(proj_file):
                    wb = load_workbook(proj_file, read_only=True)
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        for row in ws.iter_rows(min_row=3, max_row=2000, values_only=True): 
                            if not row or len(row) < 5:
                                continue
                            s_draw, s_part, s_rev, s_total = row[1], row[2], row[3], row[4]
                            if str(s_part).strip().upper() == part:
                                is_duplicate_match = False
                                for m in matches:
                                    if m['part'] == part and str(m['rev']) == str(s_rev) and m['project'] == short_proj_code:
                                        is_duplicate_match = True
                                        break
                                if not is_duplicate_match:
                                    matches.append({
                                        'source': f'FILE: {full_proj_name}', 
                                        'project': short_proj_code, 
                                        'batch': "-", 
                                        'assembly': sheet_name, 
                                        'draw': s_draw, 
                                        'rev': s_rev, 
                                        'total': s_total, 
                                        'part': s_part
                                    })
                    wb.close()
        except Exception as e:
            print(f"Error checking: {str(e)}")
        self.after(0, lambda: self._post_check(matches, part))

    def _post_check(self, matches, part_searched=None):
        self.btn_check_part.configure(state="normal", text="Search")
        self.update_status("Search complete.", COLOR_TEXT)
        if not matches:
             if part_searched:
                 self.trigger_feedback("info", "Part Not Found (Safe to add)")
             return

        candidates = {}
        for m in matches:
            key = f"{m['project']} | {m['draw']}"
            if key not in candidates:
                candidates[key] = []
            candidates[key].append(m)
        
        unique_keys = list(candidates.keys())
        if len(unique_keys) == 1:
            all_recs = candidates[unique_keys[0]]
            all_recs.sort(key=lambda x: str(x['rev']), reverse=True)
            final_candidate = all_recs[0]
            self.show_match_found_dialog(final_candidate)
        else:
            options = []
            for k in unique_keys:
                recs = candidates[k]
                recs.sort(key=lambda x: str(x['rev']), reverse=True)
                options.append(recs[0])
            self.open_selection_window(options)

    def show_match_found_dialog(self, m):
        top = ctk.CTkToplevel(self)
        top.title("Record Found")
        top.geometry("450x380")
        top.transient(self)
        top.grab_set()
        
        ctk.CTkLabel(top, text="EXISTING RECORD FOUND", font=("Segoe UI", 16, "bold"), text_color=COLOR_SUCCESS).pack(pady=15)
        info_frame = ctk.CTkFrame(top, fg_color="white", border_width=2, border_color="#BDC3C7", corner_radius=10)
        info_frame.pack(fill="x", padx=20, pady=10)
        info_frame.grid_columnconfigure(1, weight=1)

        def add_row(r, label, value):
            ctk.CTkLabel(info_frame, text=label, font=("Segoe UI", 12), text_color="gray", anchor="w").grid(row=r, column=0, padx=(15, 5), pady=5, sticky="w")
            ctk.CTkLabel(info_frame, text=value, font=("Segoe UI", 14, "bold"), text_color=COLOR_ACCENT, anchor="w").grid(row=r, column=1, padx=5, pady=5, sticky="w")

        add_row(0, "Project:", m['project'])
        add_row(1, "Assembly:", m['assembly'])
        add_row(2, "Drawing:", m['draw'])
        add_row(3, "Latest Rev:", str(m['rev']))
        add_row(4, "Part No:", m['part'])

        ctk.CTkLabel(top, text="Do you want to Auto-Fill this data?", font=("Segoe UI", 12)).pack(pady=(15, 10))
        btn_frame = ctk.CTkFrame(top, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        ctk.CTkButton(btn_frame, text="YES (Auto-Fill)", fg_color=COLOR_SUCCESS, hover_color="#219150", width=120,
                      command=lambda: [self.autofill_form(m), self.trigger_feedback("info", "Form Auto-filled!"), top.destroy()]).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="CANCEL", fg_color=COLOR_DANGER, hover_color="#C0392B", width=80, command=top.destroy).pack(side="left", padx=10)

    def open_selection_window(self, matches):
        top = ctk.CTkToplevel(self)
        top.title("Search Results")
        top.geometry("750x550")
        top.transient(self)
        top.grab_set()
        
        header_frame = ctk.CTkFrame(top, fg_color=COLOR_PRIMARY, height=70, corner_radius=0)
        header_frame.pack(fill="x")
        ctk.CTkLabel(header_frame, text="SELECT DATA TO AUTO-FILL", font=("Segoe UI", 18, "bold"), text_color="white").pack(pady=(15, 5))
        ctk.CTkLabel(header_frame, text="Please select the correct record from the list below:", font=("Segoe UI", 11), text_color="#BDC3C7").pack(pady=(0, 10))
        
        scroll = ctk.CTkScrollableFrame(top, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=10)
        
        for m in matches:
            card = ctk.CTkFrame(scroll, fg_color="white", border_width=2, border_color="#BDC3C7", corner_radius=8)
            card.pack(fill="x", pady=8, padx=5, ipady=5)
            row1 = ctk.CTkFrame(card, fg_color="transparent")
            row1.pack(fill="x", padx=15, pady=(8, 2))
            ctk.CTkLabel(row1, text=f"PROJECT: {m['project']}", font=("Segoe UI", 13, "bold"), text_color=COLOR_ACCENT).pack(side="left")
            ctk.CTkLabel(row1, text=f"{m['part']}", font=("Segoe UI", 14, "bold"), text_color=COLOR_ACCENT).pack(side="right")
            row2 = ctk.CTkFrame(card, fg_color="transparent")
            row2.pack(fill="x", padx=15, pady=2)
            ctk.CTkLabel(row2, text=f"{m['draw']}", font=("Segoe UI", 13, "bold"), text_color=COLOR_ACCENT).pack(anchor="w")
            row3 = ctk.CTkFrame(card, fg_color="transparent")
            row3.pack(fill="x", padx=15, pady=(5, 10))
            info_text = f"Assy: {m['assembly']}   |   Total Sheets: {m['total']}"
            ctk.CTkLabel(row3, text=info_text, font=("Segoe UI", 11), text_color="gray").pack(side="left")
            right_box = ctk.CTkFrame(row3, fg_color="transparent")
            right_box.pack(side="right")
            rev_lbl = ctk.CTkLabel(right_box, text=f" REV {m['rev']} ", font=("Segoe UI", 12, "bold"), 
                                   fg_color=COLOR_ACCENT, text_color="white", corner_radius=5)
            rev_lbl.pack(side="left", padx=(0, 10))
            btn_sel = ctk.CTkButton(right_box, text="SELECT", width=90, height=30, 
                                    fg_color=COLOR_SUCCESS, hover_color="#219150",
                                    font=("Segoe UI", 12, "bold"),
                                    command=lambda data=m: [self.autofill_form(data), top.destroy(), self.trigger_feedback("info", "Data Selected")])
            btn_sel.pack(side="left")
            
        ctk.CTkButton(top, text="CANCEL", fg_color=COLOR_DANGER, hover_color="#C0392B", width=120, command=top.destroy).pack(pady=15)

    def autofill_form(self, data):
        short_code = str(data.get('project')).strip()
        ui_name = REVERSE_PROJ_MAP.get(short_code)
        if ui_name:
            self.proj_v.set(ui_name)
            self.update_logic(ui_name) 
        if data.get('batch'):
            b_val = str(data.get('batch'))
            if b_val == "None" or b_val == "": 
                b_val = "-"
            self.batch_v.set(b_val)
        assy_val = str(data.get('assembly')).strip()
        self.assembly_v.set(assy_val)
        if data.get('draw'):
            self.draw_ent.delete(0, 'end')
            self.draw_ent.insert(0, str(data.get('draw')))
        if data.get('rev') is not None:
            self.rev_ent.delete(0, 'end')
            self.rev_ent.insert(0, str(data.get('rev')))
            self.auto_remark_logic(None)
        if data.get('total') is not None:
            self.total_ent.delete(0, 'end')
            self.total_ent.insert(0, str(data.get('total')))

    # --- DUPLICATE CHECKS (Dikekalkan 100%) ---
    def check_duplicate_entry(self, file_path, target_sheet_name, part_input, rev_input):
        if not os.path.exists(file_path):
            return False 
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            actual_sheet = None
            for s in wb.sheetnames:
                if s.strip().lower() == target_sheet_name.lower():
                    actual_sheet = s
                    break
            if not actual_sheet:
                wb.close()
                return False 
            ws = wb[actual_sheet]
            duplicate_found = False
            for row in ws.iter_rows(min_row=3, values_only=True):
                if row and len(row) < 4:
                    continue 
                excel_part = str(row[2]).strip().upper() if row[2] else ""
                excel_rev = str(row[3]).strip() if row[3] is not None else ""
                if excel_part == part_input and excel_rev == str(rev_input):
                    duplicate_found = True
                    break
            wb.close()
            return duplicate_found
        except Exception as e:
            print(f"Error checking duplicate: {e}")
            return False 

    def check_master_duplicate(self, part_input, rev_input, project_code):
        master_path = os.path.join(BASE_PATH, MASTER_FILE)
        if not os.path.exists(master_path):
            return False
        try:
            wb = load_workbook(master_path, read_only=True, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, max_row=8000, values_only=True):
                if row and len(row) < 7:
                    continue
                m_proj = str(row[0]).strip()
                m_part = str(row[5]).strip().upper()
                m_rev = str(row[6]).strip()
                if m_proj == project_code and m_part == part_input and m_rev == str(rev_input):
                    wb.close()
                    return True 
            wb.close()
            return False
        except Exception as e:
            print(f"Master Check Error: {e}")
            return False

    # --- FUNGSI BARU UNTUK PILIHAN DUPLICATE (ADJUSTMENT) ---
    def show_duplicate_choice_dialog(self, part, rev, sql_data, excel_master_data, master_file_path):
        """Dialog khas untuk membenarkan kemasukan ke SQL sahaja jika Excel pendua."""
        top = ctk.CTkToplevel(self)
        top.title("Duplicate Found")
        top.geometry("480x320")
        top.transient(self)
        top.grab_set()

        ctk.CTkLabel(top, text="DUPLICATE IN PROJECT EXCEL", font=("Segoe UI", 16, "bold"), text_color=COLOR_DANGER).pack(pady=15)
        ctk.CTkLabel(top, text=f"Part No: {part} (Rev {rev})\nalready exists in the sub-project Excel file.", justify="center").pack(pady=5)
        ctk.CTkLabel(top, text="What would you like to do?", font=("Segoe UI", 12, "bold")).pack(pady=15)

        btn_frame = ctk.CTkFrame(top, fg_color="transparent")
        btn_frame.pack(pady=10)

        def on_re_release():
            top.destroy()
            try:
                if not os.path.exists(master_file_path):
                    wb_m = Workbook()
                    ws_m = wb_m.active
                    ws_m.title = "MasterList"
                    ws_m.append(MASTER_HEADERS)
                    wb_m.save(master_file_path)
                
                wb_m = load_workbook(master_file_path)
                ws_m = wb_m.active
                re_release_data = list(excel_master_data)
                re_release_data[10] = "Re-Release"
                
                ws_m.append(re_release_data)
                last_row_m = ws_m.max_row
                for col_idx in range(1, len(re_release_data) + 1):
                    cell_m = ws_m.cell(row=last_row_m, column=col_idx)
                    if col_idx == 10:
                        cell_m.number_format = 'DD/MM/YYYY'
                    if col_idx in [4, 5, 6]:
                        cell_m.alignment = LEFT_ALIGN
                    else:
                        cell_m.alignment = CENTER_ALIGN
                    cell_m.border = THIN_BORDER
                wb_m.save(master_file_path)
                self.trigger_feedback("success", "Re-Release Logged (Master Only)")
                self.update_status("Success! Re-Release logged in Master List only.", "#27ae60")
            except Exception as e:
                print(e)

        def on_sql_only():
            top.destroy()
            try:
                conn = psycopg2.connect(**DB_CONFIG)
                cur = conn.cursor()
                cur.execute("INSERT INTO project_data (project_name, country, batch, main_assembly, drawing_name, part_number, revision, total_sheets, engineer, date_approved, remarks) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", sql_data)
                conn.commit()
                cur.close()
                conn.close()
                self.trigger_feedback("success", "Data Inserted into SQL Only")
                self.update_status("Success! Data added to database only.", COLOR_SUCCESS)
            except Exception as db_err:
                messagebox.showerror("SQL Error", f"Failed to insert: {db_err}")

        ctk.CTkButton(btn_frame, text="RE-RELEASE\n(MASTER ONLY)", fg_color=COLOR_PRIMARY, width=150, height=50, command=on_re_release).grid(row=0, column=0, padx=10)
        ctk.CTkButton(btn_frame, text="SQL ONLY\n(DATABASE)", fg_color=COLOR_INFO, width=150, height=50, command=on_sql_only).grid(row=0, column=1, padx=10)
        
        ctk.CTkButton(top, text="CANCEL", fg_color=COLOR_DANGER, width=80, command=top.destroy).pack(pady=15)

    # --- SUBMIT LOGIC (MEKANISME ASAL 100% DENGAN ADJUSTMENT) ---
    def submit(self):
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
            self.update_status("Error: Please fill in mandatory fields!", COLOR_DANGER)
            self.trigger_feedback("error", "Missing Mandatory Fields!")
            return
        
        try:
            rev = int(rev_str)
            total = int(total_str)
        except ValueError:
            self.trigger_feedback("error", "Revision/Total must be numbers")
            return

        target_sheet_name = self.assembly_v.get()[:31] 
        
        # STEP 0: WRITE-LOCK CHECK
        if os.path.exists(proj_file):
            try:
                with open(proj_file, "a"):
                    pass
            except PermissionError:
                self.trigger_feedback("error", "FILE OPEN! Please Close Excel.")
                messagebox.showerror("File Error", f"File '{full_name}.xlsx' is open!\nPlease close it before submitting.")
                return 

        if os.path.exists(master_file_path):
            try:
                with open(master_file_path, "a"):
                    pass
            except PermissionError:
                self.trigger_feedback("error", "MASTER FILE OPEN! Close it.")
                messagebox.showerror("File Error", f"File '{MASTER_FILE}' is open!\nPlease close it.")
                return 

        # Define Payloads
        sql_data = [short_proj, "Tanzania" if full_name == "H10 TRC" else "Malaysia", batch_val, self.assembly_v.get(), draw, part, rev, total, self.eng_v.get(), dt_sql, self.remark_v.get()]
        excel_master_data = [short_proj, "Tanzania" if full_name == "H10 TRC" else "Malaysia", batch_val_excel, self.assembly_v.get(), draw, part, rev, total, self.eng_v.get(), dt_obj, self.remark_v.get()]

        # --- ADJUSTMENT: Duplicate Check dengan SQL Only Option ---
        if self.check_duplicate_entry(proj_file, target_sheet_name, part, rev):
            self.show_duplicate_choice_dialog(part, rev, sql_data, excel_master_data, master_file_path)
            return

        # PROCEED AS NORMAL
        try:
            # SQL Insert
            try:
                conn = psycopg2.connect(**DB_CONFIG)
                cur = conn.cursor()
                cur.execute("INSERT INTO project_data (project_name, country, batch, main_assembly, drawing_name, part_number, revision, total_sheets, engineer, date_approved, remarks) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", sql_data)
                conn.commit()
                cur.close()
                conn.close()
            except Exception as db_err:
                print(f"Database Error (Skipped): {db_err}")

            # Master List Control
            if not os.path.exists(master_file_path):
                wb_m = Workbook()
                ws_m = wb_m.active
                ws_m.title = "MasterList"
                ws_m.append(MASTER_HEADERS)
                wb_m.save(master_file_path)
            
            add_to_master = True
            final_master_data = list(excel_master_data) 

            if self.check_master_duplicate(part, rev, short_proj):
                msg = (f"Record already exists in MASTER LIST (History Log):\n"
                        f"Part: {part}\nRev: {rev}\n\n"
                        f"Is this a 'Re-Release'?\n"
                        f"Click 'Yes' to add as new log (Remark: Re-Release).\n"
                        f"Click 'No' to skip Master List.")
                if messagebox.askyesno("Master List Log", msg):
                    final_master_data[10] = "Re-Release"
                    self.update_status("Logging as Re-Release in Master List...", COLOR_INFO)
                else:
                    add_to_master = False
                    self.update_status("Skipped Master List update.", COLOR_INFO)
            
            if add_to_master:
                wb_m = load_workbook(master_file_path)
                ws_m = wb_m.active
                ws_m.append(final_master_data)
                last_row_m = ws_m.max_row
                for col_idx in range(1, len(final_master_data) + 1):
                    cell_m = ws_m.cell(row=last_row_m, column=col_idx)
                    if col_idx == 10:
                        cell_m.number_format = 'DD/MM/YYYY'
                    if col_idx in [4, 5, 6]:
                        cell_m.alignment = LEFT_ALIGN
                    else:
                        cell_m.alignment = CENTER_ALIGN
                    cell_m.border = THIN_BORDER
                wb_m.save(master_file_path)

            # Project File Update
            if not os.path.exists(proj_file):
                messagebox.showerror("Error", f"File {full_name}.xlsx not found!")
                return
            
            wb_p = load_workbook(proj_file)
            actual_sheet_name = None
            for s in wb_p.sheetnames:
                if s.strip().lower() == target_sheet_name.lower():
                    actual_sheet_name = s
                    break
            
            data_start_row = 3 
            if actual_sheet_name:
                ws_p = wb_p[actual_sheet_name]
                for r in range(1, 11):
                    val = ws_p.cell(row=r, column=3).value
                    if val and "part number" in str(val).lower():
                        data_start_row = r + 1
                        break
            else:
                 ws_p = wb_p.create_sheet(target_sheet_name)
                 ws_p.append(["Document Control"])
                 ws_p.append(["Sl. No", "Drawing Name", "Part Number", "Revision", "Total Drawings", "Date Approved", "Remarks"]) 
                 ws_p.cell(row=15, column=1, value="Drawing issued by:-")
                 data_start_row = 3 

            sig_row = None
            for r in range(data_start_row, ws_p.max_row + 50):
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

            target_row = None
            is_override = False
            remarks_to_save = self.remark_v.get()
            
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
                    val_part = ws_p.cell(row=r, column=3).value
                    if (val_draw and str(val_draw).strip()) or (val_part and str(val_part).strip()):
                        last_data_row = r
                        break
                target_row = last_data_row + 1
                ws_p.insert_rows(target_row)
                current_max = ws_p.max_row
                for r in range(current_max, target_row, -1):
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
                if isinstance(cell, MergedCell):
                    continue 
                cell.value = val
                if col == 6:
                    cell.number_format = 'DD/MM/YYYY'
                if col in [2, 3]:
                    cell.alignment = LEFT_ALIGN
                else:
                    cell.alignment = CENTER_ALIGN
                cell.border = THIN_BORDER
            
            # --- SORTING ---
            sig_row_current = None
            for r in range(data_start_row, ws_p.max_row + 10):
                found = False
                for c in range(1, 6):
                    val = ws_p.cell(row=r, column=c).value
                    if val and "issued by" in str(val).lower():
                        sig_row_current = r
                        found = True
                        break
                if found:
                    break
            
            if not sig_row_current:
                sig_row_current = ws_p.max_row + 1
            sort_end_row = sig_row_current - 1
            
            data_matrix = []
            if sort_end_row >= data_start_row:
                for r in range(data_start_row, sort_end_row + 1):
                    row_data = []
                    has_data = False
                    for c in range(2, 8):
                        val = ws_p.cell(row=r, column=c).value
                        row_data.append(val)
                        if val is not None:
                            has_data = True
                    if has_data:
                        data_matrix.append(row_data)
                
                data_matrix.sort(key=lambda x: str(x[1]).strip().upper() if x[1] else "")
                
                for i, row_data in enumerate(data_matrix):
                    current_r = data_start_row + i
                    sl_cell = ws_p.cell(row=current_r, column=1)
                    sl_cell.value = i + 1
                    sl_cell.alignment = CENTER_ALIGN
                    sl_cell.border = THIN_BORDER
                    
                    for idx, val in enumerate(row_data):
                        c = idx + 2
                        cell = ws_p.cell(row=current_r, column=c)
                        cell.value = val
                        if c == 6:
                            cell.number_format = 'DD/MM/YYYY'
                        if c in [2, 3]:
                            cell.alignment = LEFT_ALIGN
                        else:
                            cell.alignment = CENTER_ALIGN
                        cell.border = THIN_BORDER

            # --- HIGHLIGHTING ---
            dates_objects = []
            for r in range(data_start_row, sig_row_current):
                d_val = ws_p.cell(row=r, column=6).value
                parsed_date = None
                if isinstance(d_val, datetime):
                    parsed_date = d_val.date()
                elif isinstance(d_val, date):
                    parsed_date = d_val
                elif d_val:
                    d_str = str(d_val).strip()
                    try:
                        parsed_date = datetime.strptime(d_str, '%d/%m/%Y').date()
                    except ValueError:
                        try:
                            parsed_date = datetime.strptime(d_str, '%Y-%m-%d').date()
                        except ValueError:
                            pass
                if parsed_date:
                    dates_objects.append(parsed_date)
            
            latest_date_obj = max(dates_objects) if dates_objects else None

            for r in range(data_start_row, sig_row_current):
                d_val = ws_p.cell(row=r, column=6).value
                current_date_obj = None
                if isinstance(d_val, datetime):
                    current_date_obj = d_val.date()
                elif isinstance(d_val, date):
                    current_date_obj = d_val
                elif d_val:
                    d_str = str(d_val).strip()
                    try:
                        current_date_obj = datetime.strptime(d_str, '%d/%m/%Y').date()
                    except ValueError:
                        try:
                            current_date_obj = datetime.strptime(d_str, '%Y-%m-%d').date()
                        except ValueError:
                            pass
                
                is_latest = (current_date_obj == latest_date_obj) and (latest_date_obj is not None)
                
                for c in range(1, 8):
                    cell = ws_p.cell(row=r, column=c)
                    if isinstance(cell, MergedCell):
                        continue
                    if c == 7 and not is_latest:
                        cell.value = None
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
            
            # --- ROW HEIGHT FIX ---
            for r in range(data_start_row, sig_row_current):
                ws_p.row_dimensions[r].height = 18

            wb_p.save(proj_file)
            action_msg = "UPDATED (Override)" if is_override else "SAVED (New)"
            self.trigger_feedback("success", f"SUCCESS! {part} {action_msg}")
            self.update_status(f"Success! Data '{part}' has been {action_msg}, SORTED & Saved.", "#27ae60")
        
        except PermissionError:
             self.trigger_feedback("error", "FILE OPEN! Please Close Excel.")
             messagebox.showerror("File Error", f"Please close file {full_name}.xlsx before submitting!")
        except Exception as e:
            self.trigger_feedback("error", "SYSTEM ERROR")
            self.update_status(f"System Error: {str(e)}", COLOR_DANGER)
            print(e)

if __name__ == "__main__":
    app = DCDEApp()
    app.mainloop()