import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from PIL import Image, ImageTk, ImageDraw
import datetime
import threading
 
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
 
app = ctk.CTk()
app.title("CCRIS Credit Report")
app.geometry("1800x900")

 
# --- Sidebar Toggle Button (3 vertical dash) ---
def create_menu_icon(size=32, color="#bbb"):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    x = size // 2
    for y in [size // 4, size // 2, 3 * size // 4]:
        draw.line([(x - 6, y), (x + 6, y)], fill=color, width=3)
    return ctk.CTkImage(img, size=(size, size))
 
menu_icon = create_menu_icon(32, "#bbb")
 
# --- Sidebar Setup ---
SIDEBAR_EXPANDED_WIDTH = 175
SIDEBAR_SHRUNK_WIDTH = 48

sidebar_container = ctk.CTkFrame(app, fg_color="transparent")
sidebar_container.pack(side="left", fill="y")

sidebar = ctk.CTkFrame(sidebar_container, width=SIDEBAR_EXPANDED_WIDTH)
sidebar.pack(side="left", fill="y")
sidebar.pack_propagate(False)  # Prevent auto-resize

# Hamburger Button (always at top left of sidebar)
hamburger_img = ctk.CTkImage(Image.open("hamburger.png"), size=(24, 24))
hamburger_btn = ctk.CTkButton(
    sidebar,
    text="",
    image=hamburger_img,
    width=40,
    height=40,
    fg_color="transparent",
    hover_color="#333",
    command=lambda: toggle_sidebar()
)
hamburger_btn.pack(pady=(8, 0), padx=(4, 0), anchor="nw")

# --- Hamburger Button ---
hamburger_img = ctk.CTkImage(Image.open("hamburger.png"), size=(24, 24))
hamburger_btn = ctk.CTkButton(
    app,
    text="",
    image=hamburger_img,
    width=32,
    height=40,
    fg_color="transparent",
    hover_color="#333",
    command=lambda: toggle_sidebar()
)
hamburger_btn.place(x=8, y=8)

# # --- Sidebar Toggle Logic ---
sidebar_expanded = True


# Menu label and buttons (put in a frame for easy show/hide)
menu_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
menu_frame.pack(fill="both", expand=True)

ctk.CTkLabel(menu_frame, text="Menu", font=("Arial", 18, "bold")).pack(pady=(12, 10))
btn_report = ctk.CTkButton(
    menu_frame,
    text="CCRIS Report",
    width=150,
    height=40,
    font=("Arial", 15, "bold"),
    corner_radius=10,
)
btn_report.pack(pady=10, padx=(0, 0))

btn_another = ctk.CTkButton(
    menu_frame,
    text="Excel All Task",
    width=150,
    height=40,
    font=("Arial", 15, "bold"),
    corner_radius=10,
)
btn_another.pack(pady=10, padx=(0, 0))

def toggle_sidebar():
    global sidebar_expanded
    if sidebar_expanded:
        sidebar.configure(width=SIDEBAR_SHRUNK_WIDTH)
        menu_frame.pack_forget()
        sidebar_expanded = False
    else:
        sidebar.configure(width=SIDEBAR_EXPANDED_WIDTH)
        menu_frame.pack(fill="both", expand=True)
        sidebar_expanded = True

# --- Start with sidebar expanded ---
sidebar.configure(width=SIDEBAR_EXPANDED_WIDTH)
menu_frame.pack(fill="both", expand=True)

 
# Placeholder for button commands
def do_nothing():
    pass
 
# --- CCRIS Report Class ---
class CCRISReport:
    def __init__(self, parent):
        self.parent = parent
        
        
        # Set Treeview style to dark before creating any Treeview
        self.set_treeview_style("dark")
        
        # --- Scrollable Frame Setup ---
        self.outer_frame = ctk.CTkFrame(parent)
        self.outer_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.task_tab = TaskTabBar(self.outer_frame)
        
        # --- Loading Overlay ---
        self.loading_gif = Image.open("loading.gif")
        self.loading_frames = []
        size = (64, 64)
        for angle in range(0, 360, 30):
            frame = self.loading_gif.copy().resize(size, Image.LANCZOS).convert("RGBA")
            rotated = frame.rotate(angle)
            self.loading_frames.append(ctk.CTkImage(rotated, size=size))
        self.loading_label = ctk.CTkLabel(self.outer_frame, text="")
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label.lower()
        self.loading_gif_running = False

        
        # Use a regular Canvas for scrolling
        self.canvas = tk.Canvas(self.outer_frame, borderwidth=0, highlightthickness=0, bg="#222222")
        self.scrollbar = ctk.CTkScrollbar(self.outer_frame, orientation="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Create a frame inside the canvas
        self.frame = ctk.CTkFrame(self.canvas)
        self.frame_id = self.canvas.create_window((0, 0), window=self.frame, anchor="nw")

        # Update scrollregion when the frame changes size
        def on_frame_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            # Make the canvas's width match the outer_frame's width
            self.canvas.itemconfig(self.frame_id, width=self.canvas.winfo_width())
        self.frame.bind("<Configure>", on_frame_configure)

        # Also update canvas width when the window is resized
        def on_canvas_configure(event):
            self.canvas.itemconfig(self.frame_id, width=event.width)
        self.canvas.bind("<Configure>", on_canvas_configure)

        # # Make mousewheel scroll work
        # def _on_mousewheel(event):
        #     self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # self.outer_frame.pack_forget()
 
        # Header
        self.header = ctk.CTkFrame(self.frame)
        self.header.pack(fill="x", pady=10)
        logo_image = ctk.CTkImage(Image.open("bnm_logo.png"), size=(300, 56))
        alrajhi_logo_image = ctk.CTkImage(Image.open("alrajhi_logo.png"), size=(170, 60))
        self.header.columnconfigure((0, 1, 2), weight=1)
        ctk.CTkLabel(self.header, image=logo_image, text="").grid(row=0, column=1, padx=10, pady=5, sticky="nsew")
        ctk.CTkLabel(self.header, text="CREDIT REPORT", font=("Arial", 22, "bold")).grid(row=0, column=0, padx=(20, 10), sticky="w")
        ctk.CTkLabel(self.header, image=alrajhi_logo_image, text="").grid(row=0, column=2, padx=10, pady=5, sticky="nsew")
 
        # Controls
        self.control_frame = ctk.CTkFrame(self.frame)
        self.control_frame.pack(fill="x", pady=5)
 
        ctk.CTkButton(self.control_frame, text="Import CCRIS Excel", command=self.load_excel).pack(side="left", padx=10)
        self.selected_pg_rqs = ctk.StringVar()
        style = ttk.Style()
        style.configure("Custom.TCombobox", font=("Arial", 16))
        self.pg_dropdown = ttk.Combobox(
            self.control_frame,
            textvariable=self.selected_pg_rqs,
            width=25,
            style="Custom.TCombobox"
        )
        self.pg_dropdown.pack(side="left", padx=10)
        self.pg_dropdown.bind("<<ComboboxSelected>>", lambda event: self.load_pg_data())
 
        self.arrears_label = ctk.CTkLabel(self.control_frame, text="Arrears in 12 Months:")
        self.arrears_label.pack(side="left", padx=20)
 
        # Dark mode toggle
        light_icon = ctk.CTkImage(Image.open("light_mode_icon.png"), size=(24, 24))
        dark_icon = ctk.CTkImage(Image.open("dark_mode_icon.png"), size=(24, 24))
        self.current_icon = {"mode": "dark"}
        self.mode_icon_btn = ctk.CTkButton(self.control_frame, text="", image=light_icon, width=32, height=32, command=self.toggle_mode_icon)
        self.mode_icon_btn.pack(side="right", padx=10)
        self.light_icon = light_icon
        self.dark_icon = dark_icon
 
        # Table Section
        self.table_section = ctk.CTkFrame(self.frame)
        self.table_section.pack(fill="both", expand=True, padx=10, pady=10)
        self.table_section.update_idletasks()
        self.table_section.configure(width=1800)  # or a value wide enough for all columns
 
        # Table columns
        self.outstanding_cols = ["No", "Approval Date", "Status", "Capacity", "Lender", "Branch", "Facility",
                                 "Total Outstanding", "Balance Date", "Limit", "Collateral", "Repayment Term",
                                 "12-Month Arrears", "Legal Status", "Legal Date"]
 
        # Outstanding Credit
        ctk.CTkLabel(self.table_section, text="Outstanding Credit", font=("Arial", 14, "bold")).pack(anchor="w")
        self.outstanding_tree = self.create_table(self.table_section, self.outstanding_cols, height=6)
 
        # Special Attention
        ctk.CTkLabel(self.table_section, text="Special Attention Account", font=("Arial", 14, "bold")).pack(anchor="w")
        self.attention_tree = self.create_table(self.table_section, self.outstanding_cols, height=4)
 
        # Application for Credit
        ctk.CTkLabel(self.table_section, text="Application for Credit", font=("Arial", 14, "bold")).pack(anchor="w")
        self.application_tree = self.create_table(self.table_section, self.outstanding_cols, height=4)
 
        # Data
        self.excel_data = {}
        self.pg_list = []
 
        # Hide by default (will be shown by sidebar button)
        self.frame.pack_forget()
       
    def show_loading(self):
        self.loading_label.lift()
        self.loading_gif_running = True
        self.animate_loading_gif(0)

    def hide_loading(self):
        self.loading_label.lower()
        self.loading_gif_running = False

    def animate_loading_gif(self, idx):
        if not self.loading_gif_running:
            return
        frame = self.loading_frames[idx]
        self.loading_label.configure(image=frame, text="")
        next_idx = (idx + 1) % len(self.loading_frames)
        self.outer_frame.after(60, lambda: self.animate_loading_gif(next_idx))

 
    def show(self):
        self.outer_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def hide(self):
        self.outer_frame.pack_forget()
 
    def load_excel(self):
        self.show_loading()
        def do_import():
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
            if not file_path:
                self.hide_loading()
                return
            xls = pd.ExcelFile(file_path)
            for sheet in ["part_1", "part_2", "part_3", "part_4"]:
                if sheet in xls.sheet_names:
                    df = xls.parse(sheet, dtype=str).fillna("NaN")
                    self.excel_data[sheet] = df

            if "part_1" in self.excel_data:
                self.pg_list = self.excel_data["part_1"]["PG_RQS"].dropna().unique().tolist()
                self.pg_dropdown.configure(values=self.pg_list)
                if self.pg_list:
                    self.selected_pg_rqs.set(self.pg_list[0])
                    self.load_pg_data()
            self.hide_loading()
        threading.Thread(target=do_import, daemon=True).start()
 
    def load_pg_data(self):
        pg = self.selected_pg_rqs.get()
        if not pg:
            return
 
        self.clear_table(self.outstanding_tree)
        self.clear_table(self.attention_tree)
        self.clear_table(self.application_tree)
 
        def format_date(val):
            # Try to parse and format as dd-mm-yyyy, else return as is or '-'
            if val in ["NaN", "-", "", None] or pd.isna(val):
                return "-"
            try:
                dt = pd.to_datetime(val, errors='coerce')
                if pd.isna(dt):
                    return val
                return dt.strftime("%d-%m-%Y")
            except Exception:
                return val
 
        # --- Calculate Report Date for part_4 (Application for Credit) ---
        report_date_str = "-"
        pending_count_last_month = 0
        if "part_4" in self.excel_data:
            df_part4 = self.excel_data["part_4"]
            df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg]
            if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                # Get the latest report date
                latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                if pd.notna(latest_report_date):
                    report_date_str = latest_report_date.strftime("%d-%m-%Y")
                    one_month_ago = latest_report_date - pd.DateOffset(months=1)
                    # Count pending where approval date is in last 1 month of report date
                    mask = (
                        (df_pg_part4["APPL_STS"] == "P") &
                        (pd.to_datetime(df_pg_part4["DT_APPL"], errors="coerce") >= one_month_ago) &
                        (pd.to_datetime(df_pg_part4["DT_APPL"], errors="coerce") <= latest_report_date)
                    )
                    pending_count_last_month = mask.sum()
 
        # ...existing code for filling tables...
        for part, tree in zip(["part_2", "part_3", "part_4"], [self.outstanding_tree, self.attention_tree, self.application_tree]):
            df = self.excel_data.get(part, pd.DataFrame())
            df_pg = df[df["PG_RQS"] == pg]
            for _, row in df_pg.iterrows():
                values = [
                    row.get("REC_CTR", "NaN"),
                    format_date(row.get("DT_APPL", "NaN")),  # Approval Date formatted
                    row.get("APPL_STS", "NaN"),
                    row.get("CPY", "NaN"),
                    row.get("LEND_TYPE", "NaN"),
                    row.get("BRANCH", "NaN"),
                    row.get("FCY_TYPE", "NaN"),
                    row.get("IM_AM", "NaN"),
                    format_date(row.get("DT_BAL", "NaN")),   # Balance Date formatted
                    row.get("IM_LIM_AM", "NaN"),
                    row.get("COL_TYPE", "NaN"),
                    row.get("RPY_TERM", "NaN"),
                    row.get("MTH_N", "NaN"),
                    row.get("LEGAL_STS", "NaN"),
                    row.get("DT_LEGAL", "NaN")
                ]
                values = ['-' if (v == "NaN" or str(v).strip() == "" or pd.isna(v)) else v for v in values]
                tree.insert("", "end", values=values)
            self.update_table_height(tree, len(df_pg), min_height=4, max_height=20)
 
        # Optional: Load MTH_C from part_1
        mth_c_value = "12-Month Arrears"
        if "part_1" in self.excel_data:
            arrears_df = self.excel_data["part_1"]
            mth_c = arrears_df.loc[arrears_df["PG_RQS"] == pg, "MTH_C"]
            if not mth_c.empty:
                mth_c_value = mth_c.iloc[0]
 
        # --- Show report date beside arrears label ---
        self.arrears_label.configure(
            text=f"Arrears in 12 Months: {mth_c_value}    |    Report Date: {report_date_str}"
        )
        self.attention_tree.heading("12-Month Arrears", text=mth_c_value)
        self.application_tree.heading("12-Month Arrears", text=mth_c_value)
        self.outstanding_tree.heading("12-Month Arrears", text=mth_c_value)
 
        # self.task_tab.show_content(self.excel_data, pg)
        self.task_tab.set_data(self.excel_data, pg)
        
       
    def clear_table(self, tree):
        for row in tree.get_children():
            tree.delete(row)
 
    def create_table(self, parent, columns, height=5):
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="both", expand=True, pady=5)

        tree = ttk.Treeview(frame, columns=columns, show="headings", height=height)
        tree.pack(side="left", fill="both", expand=True)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center", width=120, stretch=True)
        return tree
 
    def update_table_height(self, tree, data_len, min_height=4, max_height=20):
        tree.config(height=max(min_height, min(data_len, max_height)))
 
    def set_treeview_style(self, mode):
        style = ttk.Style()
        if mode == "dark":
            style = ttk.Style()
            style.theme_use("default")
            style.configure("Treeview",
                            background="#222222",
                            foreground="#ffffff",
                            fieldbackground="#222222",
                            rowheight=25)
            style.map("Treeview", background=[("selected", "#444444")])
        else:
            style.theme_use("default")
            style.configure("Treeview",
                            background="#ffffff",
                            foreground="#000000",
                            fieldbackground="#ffffff",
                            rowheight=25)
            style.map("Treeview", background=[("selected", "#cce6ff")])
 
    def toggle_mode_icon(self):
        current = ctk.get_appearance_mode()
        if current == "Light":
            ctk.set_appearance_mode("dark")
            self.set_treeview_style("dark")
            self.mode_icon_btn.configure(image=self.light_icon)
            self.current_icon["mode"] = "dark"
        else:
            ctk.set_appearance_mode("light")
            self.set_treeview_style("light")
            self.mode_icon_btn.configure(image=self.dark_icon)
            self.current_icon["mode"] = "light"

    
    

class TaskTabBar:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ctk.CTkFrame(parent)
        self.frame.pack(side="bottom", fill="x", pady=(0, 0))
        self.visible = False

        # Tab bar (always visible, at the top of the panel)
        self.tab_bar = ctk.CTkFrame(self.frame)
        self.tab_bar.pack(fill="x", side="top")
        self.task_btn = ctk.CTkButton(
            self.tab_bar,
            text="Task",
            width=120,
            height=32,
            fg_color="#222",
            text_color="#fff",
            hover_color="#333",
            corner_radius=0,
            font=("Arial", 15, "bold"),
            command=self._on_show
        )
        self.task_btn.pack(side="left", padx=(0, 2), pady=0)

        close_icon = ctk.CTkImage(Image.open("close.png"), size=(24, 24))
        self.close_btn = ctk.CTkButton(
            self.tab_bar,
            text="",
            image=close_icon,
            width=32,
            height=32,
            fg_color="transparent",
            hover_color="#d32f2f",
            command=self.hide_content,
            text_color="white"
        )
        self.close_btn.pack(side="right", padx=(4, 8), pady=0)

        # Content area (hidden by default)
        self.content_label = ctk.CTkLabel(
            self.frame,
            text="",
            font=("Arial", 15),
            anchor="w",
            justify="left"
        )

        # For animation
        self.animating = False
        self.target_height = 320  # Adjust as needed

        # These will be set by the parent before showing
        self.excel_data = None
        self.pg = None

    def set_data(self, excel_data, pg):
        self.excel_data = excel_data
        self.pg = pg

    def _on_show(self):
        if self.excel_data is not None and self.pg is not None:
            self.show_content(self.excel_data, self.pg)

    def show_content(self, excel_data, pg):
        # --- Build calculation text ---
        task1_text = ""
        pending_numbers = []
        pending_count_last_month = 0
        report_date_str = "-"
        mth_c_value = "12-Month Arrears"
        if "part_4" in excel_data and pg:
            df_part4 = excel_data["part_4"]
            df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg]
            if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                if pd.notna(latest_report_date):
                    report_date_str = latest_report_date.strftime("%d-%m-%Y")
                    one_month_ago = latest_report_date - pd.DateOffset(months=1)
                    mask = (
                        (df_pg_part4["APPL_STS"] == "P") &
                        (pd.to_datetime(df_pg_part4["DT_APPL"], errors="coerce") >= one_month_ago) &
                        (pd.to_datetime(df_pg_part4["DT_APPL"], errors="coerce") <= latest_report_date)
                    )
                    pending_count_last_month = mask.sum()
                    pending_numbers = df_pg_part4.loc[mask, "REC_CTR"].tolist()
        if "part_1" in excel_data and pg:
            arrears_df = excel_data["part_1"]
            mth_c = arrears_df.loc[arrears_df["PG_RQS"] == pg, "MTH_C"]
            if not mth_c.empty:
                mth_c_value = mth_c.iloc[0]
        if pending_numbers:
            pending_numbers_str = ", ".join(str(n) for n in pending_numbers)
        else:
            pending_numbers_str = "-"
        task1_text = (
            f"Task 1:\n"
            f"Pending applications in last One month: {pending_count_last_month}\n"
            f"Row No : {pending_numbers_str}\n"
            f"Report Date: {report_date_str}\n"
        )

        # --- Task 2: CRDTCARD Outstanding Ratio ---
        task2_result = "-"
        task2_text = "Task 2:\nCRDTCARD Outstanding : -\n"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce")
            df_pg_part2["IM_LIM_AM"] = pd.to_numeric(df_pg_part2["IM_LIM_AM"], errors="coerce")
            total_outstanding = df_pg_part2["IM_LIM_AM"].sum()
            crdtcard_outstanding = df_pg_part2.loc[df_pg_part2["FCY_TYPE"] == "CRDTCARD", "IM_AM"].sum()
            if total_outstanding > 0:
                ratio = crdtcard_outstanding / total_outstanding
                task2_result = f"{ratio:.2%}"
            else:
                task2_result = "No outstanding found"
            task2_text = f"Task 2:\nCRDTCARD Outstanding : {task2_result}\n"

        # --- Task 4: Number of unsecured facilities in last 12 months ---
        task4_result = "-"
        task4_rows = "-"
        if "part_2" in excel_data and "part_4" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_part4 = excel_data["part_4"]
            df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg]
            if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                if pd.notna(report_date):
                    one_year_ago = report_date - pd.DateOffset(months=12)
                    df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
                    df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
                    mask = (
                        (df_pg_part2["COL_TYPE"] == "0") &
                        (df_pg_part2["DT_APPL"] >= one_year_ago) &
                        (df_pg_part2["DT_APPL"] <= report_date)
                    )
                    unsecured_rows = df_pg_part2[mask]
                    task4_result = len(unsecured_rows)
                    task4_rows = ", ".join(unsecured_rows["REC_CTR"].astype(str).tolist()) if not unsecured_rows.empty else "-"
        task4_text = (
            f"Task 4:\n"
            f"Number of unsecured facilities in last 12 months: {task4_result}\n"
            f"Row No: {task4_rows}\n"
        )

        # --- Task 5: Number of unsecured facilities in last 18 months ---
        task5_result = "-"
        task5_rows = "-"
        if "part_2" in excel_data and "part_4" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_part4 = excel_data["part_4"]
            df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg]
            if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                if pd.notna(report_date):
                    eighteen_months_ago = report_date - pd.DateOffset(months=18)
                    df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
                    df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
                    mask = (
                        (df_pg_part2["COL_TYPE"] == "0") &
                        (df_pg_part2["DT_APPL"] >= eighteen_months_ago) &
                        (df_pg_part2["DT_APPL"] <= report_date)
                    )
                    unsecured_rows_18 = df_pg_part2[mask]
                    task5_result = len(unsecured_rows_18)
                    task5_rows = ", ".join(unsecured_rows_18["REC_CTR"].astype(str).tolist()) if not unsecured_rows_18.empty else "-"
        task5_text = (
            f"Task 5:\n"
            f"Number of unsecured facilities in last 18 months: {task5_result}\n"
            f"Row No: {task5_rows}\n"
        )

        # --- Task 7: Secured financing (Collateral ≠ 0) ---
        task7_count = "-"
        task7_outstanding = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce")
            secured = df_pg_part2[df_pg_part2["COL_TYPE"] != "0"]
            task7_count = len(secured)
            task7_outstanding = f"{secured['IM_AM'].sum():,.2f}" if not secured.empty else "0.00"
        task7_text = (
            f"Task 7:\n"
            f"a. Total number of secured facilities: {task7_count}\n"
            f"b. Total outstanding: {task7_outstanding}\n"
        )

        # --- Task 8: Unsecured financing (Collateral = 0) ---
        task8_count = "-"
        task8_outstanding = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce")
            unsecured = df_pg_part2[df_pg_part2["COL_TYPE"] == "0"]
            task8_count = len(unsecured)
            task8_outstanding = f"{unsecured['IM_AM'].sum():,.2f}" if not unsecured.empty else "0.00"
        task8_text = (
            f"Task 8:\n"
            f"a. Total number of unsecured facilities: {task8_count}\n"
            f"b. Total outstanding: {task8_outstanding}\n"
        )

        all_tasks_text = f"{task1_text}\n{task2_text}\n{task4_text}\n{task5_text}\n{task7_text}\n{task8_text}"
        self.content_label.configure(text=all_tasks_text)
        if not self.visible:
            self.content_label.pack(fill="both", padx=10, pady=10, side="top")
            self.content_label.update_idletasks()
            self.content_label.configure(height=0)
            self.visible = True
            self.animate_show()

    def hide_content(self):
        self.content_label.pack_forget()
        self.visible = False
        self.animating = False

    def animate_show(self):
        self.animating = True
        step = 32
        max_height = self.target_height
        def grow():
            h = self.content_label.winfo_height()
            if h < max_height:
                h = min(h + step, max_height)
                self.content_label.configure(height=h)
                self.frame.after(10, grow)
            else:
                self.content_label.configure(height=max_height)
                self.animating = False
        grow()

    def animate_hide(self):
        self.animating = True
        step = 32
        def shrink():
            h = self.content_label.winfo_height()
            if h > 0:
                h = max(h - step, 0)
                self.content_label.configure(height=h)
                self.frame.after(10, shrink)
            else:
                self.content_label.pack_forget()
                self.visible = False
                self.animating = False
        shrink()

class ExcelAllTask:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ctk.CTkFrame(parent)
        ctk.CTkLabel(self.frame, text="Excel All Task - Calculation Display", font=("Arial", 18, "bold")).pack(pady=10)

        
        # --- Animated GIF Loading (rotating and small) ---
        self.loading_gif = Image.open("loading.gif")
        self.loading_frames = []
        size = (42, 42)  # Small icon
        for angle in range(0, 360, 30):  # 12 frames for smooth rotation
            frame = self.loading_gif.copy().resize(size, Image.LANCZOS).convert("RGBA")
            rotated = frame.rotate(angle)
            self.loading_frames.append(ctk.CTkImage(rotated, size=size))

        self.loading_label = ctk.CTkLabel(self.frame, text="")
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label.lower()  # Hide by default
        self.loading_gif_running = False

        # Define columns: PG_RQS + Task 1-8
        self.columns = ["PG_RQS"] + [f"Task {i}" for i in range(1, 9)]

        # Create a frame to hold the Treeview and scrollbars
        tree_frame = ctk.CTkFrame(self.frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create the Treeview
        self.tree = ttk.Treeview(tree_frame, columns=self.columns, show="headings", height=20)
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=180, stretch=True)
        self.tree.pack(side="left", fill="both", expand=True)

        # Add vertical scrollbar
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        yscroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=yscroll.set)

    
    def show_loading(self):
        self.loading_label.lift()
        self.loading_gif_running = True
        self.animate_loading_gif(0)

    def hide_loading(self):
        self.loading_label.lower()
        self.loading_gif_running = False

    def animate_loading_gif(self, idx):
        if not self.loading_gif_running:
            return
        frame = self.loading_frames[idx]
        self.loading_label.configure(image=frame, text="")
        next_idx = (idx + 1) % len(self.loading_frames)
        self.frame.after(60, lambda: self.animate_loading_gif(next_idx))

    def show(self):
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.show_loading()
        threading.Thread(target=self.populate_grid, daemon=True).start()  # Run populate_grid in a thread

    def hide(self):
        self.frame.pack_forget()

    def populate_grid(self):
        # Clear previous data
        for row in self.tree.get_children():
            self.tree.delete(row)

        global ccris_report
        excel_data = ccris_report.excel_data

        if "part_1" not in excel_data:
            self.hide_loading()
            return
        pg_list = pd.Series(excel_data["part_1"]["PG_RQS"].unique()).dropna().tolist()

        batch_size = 100
        max_rows = 500  # or 7000 for production
        for start in range(0, min(max_rows, len(pg_list)), batch_size):
            end = min(start + batch_size, len(pg_list))
            for i in range(start, end):
                pg = pg_list[i]
                task_summaries = self.get_task_summaries_for_pg(pg, excel_data)
                self.tree.insert("", "end", values=[pg] + task_summaries)
            self.frame.update_idletasks()  # Update the UI after each batch
        self.hide_loading()


    def get_latest_report_date(self, df_pg_part4, df_pg_part2):
        # Try to get latest date from part_4
        if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
            date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], format="%d/%m/%Y", errors="coerce").max()
            if pd.notna(date):
                return date
        # Fallback: try from part_2 DT_APPL
        if not df_pg_part2.empty and "DT_APPL" in df_pg_part2.columns:
            date = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce").max()
            if pd.notna(date):
                return date
        return None

    def get_task_summaries_for_pg(self, pg, excel_data):
        # Prepare dataframes
        df_part2 = excel_data.get("part_2", pd.DataFrame())
        df_part4 = excel_data.get("part_4", pd.DataFrame())
        df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy() if not df_part2.empty else pd.DataFrame()
        df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg].copy() if not df_part4.empty else pd.DataFrame()

        # Task 1: Pending applications in last One month
        task1 = "-"
        pending_count_last_month = 0
        pending_numbers_str = "-"
        report_date_str = "-"
        mth_c_value = "12-Month Arrears"
        if "part_4" in excel_data and pg:
            df_part4 = excel_data["part_4"]
            df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg]
            if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], format="%d/%m/%Y", errors="coerce").max()
                if pd.notna(latest_report_date):
                    report_date_str = latest_report_date.strftime("%d-%m-%Y")
                    one_month_ago = latest_report_date - pd.DateOffset(months=1)
                    mask = (
                        (df_pg_part4["APPL_STS"] == "P") &
                        (pd.to_datetime(df_pg_part4["DT_APPL"], errors="coerce") >= one_month_ago) &
                        (pd.to_datetime(df_pg_part4["DT_APPL"], errors="coerce") <= latest_report_date)
                    )
                    pending_count_last_month = mask.sum()
                    pending_numbers = df_pg_part4.loc[mask, "REC_CTR"].tolist()
                    if pending_numbers:
                        pending_numbers_str = ", ".join(str(n) for n in pending_numbers)
        if "part_1" in excel_data and pg:
            arrears_df = excel_data["part_1"]
            mth_c = arrears_df.loc[arrears_df["PG_RQS"] == pg, "MTH_C"]
            if not mth_c.empty:
                mth_c_value = mth_c.iloc[0]
        task1 = f"{pending_count_last_month} (Rows: {pending_numbers_str}, Report: {report_date_str})"

        # Task 2: CRDTCARD Outstanding
        task2 = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce")
            df_pg_part2["IM_LIM_AM"] = pd.to_numeric(df_pg_part2["IM_LIM_AM"], errors="coerce")
            total_outstanding = df_pg_part2["IM_LIM_AM"].sum()
            crdtcard_outstanding = df_pg_part2.loc[df_pg_part2["FCY_TYPE"] == "CRDTCARD", "IM_AM"].sum()
            if total_outstanding > 0:
                ratio = crdtcard_outstanding / total_outstanding
                task2 = f" {ratio:.2%}"
            else:
                task2 = "No outstanding found"

        # Task 3: Oldest approval date in part_2 (all users, not per PG_RQS)
        task3 = "-"
        if "part_2" in excel_data:
            df_part2 = excel_data["part_2"].copy()
            df_part2["DT_APPL"] = pd.to_datetime(df_part2["DT_APPL"], errors="coerce")
            if not df_part2["DT_APPL"].isna().all():
                oldest_date = df_part2["DT_APPL"].min()
                oldest_rows = df_part2[df_part2["DT_APPL"] == oldest_date]
                row_numbers = ", ".join(oldest_rows["REC_CTR"].astype(str).tolist())
                task3 = f"{oldest_date.strftime('%d-%m-%Y')} (Rows: {row_numbers})"

        # Task 4: Number of unsecured facilities in last 12 months
        task4 = "-"
        latest_report_date = self.get_latest_report_date(df_pg_part4, df_pg_part2)
        if latest_report_date is not None and not df_pg_part2.empty:
            one_year_ago = latest_report_date - pd.DateOffset(months=12)
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            mask = (
                (df_pg_part2["COL_TYPE"] == "0") &
                (df_pg_part2["DT_APPL"] >= one_year_ago) &
                (df_pg_part2["DT_APPL"] <= latest_report_date)
            )
            unsecured_rows = df_pg_part2[mask]
            task4 = f"{len(unsecured_rows)} (Rows: {', '.join(unsecured_rows['REC_CTR'].astype(str).tolist()) if not unsecured_rows.empty else '-'})"

        # Task 5: Number of unsecured facilities in last 18 months
        task5 = "-"
        if latest_report_date is not None and not df_pg_part2.empty:
            eighteen_months_ago = latest_report_date - pd.DateOffset(months=18)
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            mask = (
                (df_pg_part2["COL_TYPE"] == "0") &
                (df_pg_part2["DT_APPL"] >= eighteen_months_ago) &
                (df_pg_part2["DT_APPL"] <= latest_report_date)
            )
            unsecured_rows_18 = df_pg_part2[mask]
            task5 = f"{len(unsecured_rows_18)} (Rows: {', '.join(unsecured_rows_18['REC_CTR'].astype(str).tolist()) if not unsecured_rows_18.empty else '-'})"

        # Task 6: (Placeholder, fill in your logic)
        task6 = "-"

        # Task 7: Secured financing (Collateral ≠ 0)
        task7 = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce")
            secured = df_pg_part2[df_pg_part2["COL_TYPE"] != "0"]
            count = len(secured)
            outstanding = f"{secured['IM_AM'].sum():,.2f}" if not secured.empty else "0.00"
            task7 = f"{count} (Outstanding: {outstanding})"

        # Task 8: Unsecured financing (Collateral = 0)
        task8 = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce")
            unsecured = df_pg_part2[df_pg_part2["COL_TYPE"] == "0"]
            count = len(unsecured)
            outstanding = f"{unsecured['IM_AM'].sum():,.2f}" if not unsecured.empty else "0.00"
            task8 = f"{count} (Outstanding: {outstanding})"

        return [task1, task2, task3, task4, task5, task6, task7, task8]
 
# --- Main Content Area ---
main_content = ctk.CTkFrame(app)
main_content.pack(side="left", fill="both", expand=True, padx=(18, 0))  # Increased left padding for more space between sidebar and main content

# Instantiate classes
ccris_report = CCRISReport(main_content)
excel_all_task = ExcelAllTask(main_content)
 
# --- Sidebar Button Commands ---
def show_ccris_report():
    excel_all_task.hide()
    ccris_report.show()
 
def show_excel_all_task():
    ccris_report.hide()
    excel_all_task.show()
 
btn_report.configure(command=show_ccris_report)
btn_another.configure(command=show_excel_all_task)
 
# Show CCRIS report by default
show_ccris_report()
 
app.mainloop()
 
