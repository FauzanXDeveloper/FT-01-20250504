import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from PIL import Image, ImageTk, ImageDraw
import datetime

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("CCRIS Credit Report")
app.state("zoomed")

# --- Sidebar Toggle Button (3 vertical dash) ---
def create_menu_icon(size=32, color="#bbb"):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    x = size // 2
    for y in [size // 4, size // 2, 3 * size // 4]:
        draw.line([(x - 6, y), (x + 6, y)], fill=color, width=3)
    return ctk.CTkImage(img, size=(size, size))

menu_icon = create_menu_icon(32, "#bbb")

# Place the sidebar in a container
sidebar_container = ctk.CTkFrame(app, fg_color="transparent")
sidebar_container.pack(side="left", fill="y")

sidebar = ctk.CTkFrame(sidebar_container, width=175)
sidebar.pack(side="left", fill="y")
# Load the PNG icon
hamburger_img = ctk.CTkImage(Image.open("hamburger.png"), size=(24, 24))

# Hamburger button (place it directly on the app, not in sidebar_container)
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
hamburger_btn.lower()  # Hide below sidebar at start

def toggle_sidebar():
    if sidebar_container.winfo_ismapped():
        sidebar_container.pack_forget()
        hamburger_btn.lift()  # Show hamburger button
    else:
        sidebar_container.pack(side="left", fill="y")
        hamburger_btn.lower()  # Hide hamburger button under sidebar
# ...rest of your sidebar code...
ctk.CTkLabel(sidebar, text="Menu", font=("Arial", 18, "bold")).pack(pady=(20, 10))
btn_report = ctk.CTkButton(
    sidebar,
    text="CCRIS Report",
    width=170,
    height=40,
    font=("Arial", 15, "bold"),
    corner_radius=2,  # Rounded all corners
)
btn_report.pack(pady=10, padx=(0, 0))

btn_another = ctk.CTkButton(
    sidebar,
    text="Excel All Task",
    width=170,
    height=40,
    font=("Arial", 15, "bold"),
    corner_radius=2,
)
btn_another.pack(pady=10, padx=(0, 0))

# Placeholder for button commands
def do_nothing():
    pass

# --- CCRIS Report Class ---
class CCRISReport:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ctk.CTkFrame(parent)
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)

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
        self.pg_dropdown = ttk.Combobox(self.control_frame, textvariable=self.selected_pg_rqs)
        self.pg_dropdown.pack(side="left", padx=10)
        self.pg_dropdown.bind("<<ComboboxSelected>>", lambda event: self.load_pg_data())

        self.arrears_label = ctk.CTkLabel(self.control_frame, text="Arrears in 12 Months:")
        self.arrears_label.pack(side="left", padx=20)

        # Dark mode toggle
        light_icon = ctk.CTkImage(Image.open("light_mode_icon.png"), size=(24, 24))
        dark_icon = ctk.CTkImage(Image.open("dark_mode_icon.png"), size=(24, 24))
        self.current_icon = {"mode": "dark"}
        self.mode_icon_btn = ctk.CTkButton(self.control_frame, text="", image=dark_icon, width=32, height=32, command=self.toggle_mode_icon)
        self.mode_icon_btn.pack(side="right", padx=10)
        self.light_icon = light_icon
        self.dark_icon = dark_icon

        # Table Section
        self.table_section = ctk.CTkFrame(self.frame)
        self.table_section.pack(fill="both", expand=True, padx=10, pady=10)

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
        
            # --- TASK Tab Bar (like VS Code panel) ---
        self.task_tab_frame = ctk.CTkFrame(self.table_section)
        self.task_tab_frame.pack(fill="x", pady=(10, 0), side="top")

        # Only one Task button
        self.task_btn = ctk.CTkButton(
            self.task_tab_frame,
            text="Task",
            width=120,
            height=32,
            fg_color="#222",
            text_color="#fff",
            hover_color="#333",
            corner_radius=0,
            font=("Arial", 15, "bold"),
            command=self.show_task_content
        )
        self.task_btn.pack(side="left", padx=(0, 2), pady=0)

        # Close button at the same level
        close_icon = ctk.CTkImage(Image.open("close.png"), size=(24, 24))
        self.close_btn = ctk.CTkButton(
            self.task_tab_frame,
            text="",
            image=close_icon,
            width=32,
            height=32,
            fg_color="transparent",
            hover_color="#d32f2f",
            command=self.hide_task_content,
            text_color="white"
        )
        self.close_btn.pack(side="right", padx=(4, 8), pady=0)

        # Content area for all tasks
        self.task_content_label = ctk.CTkLabel(
            self.table_section,
            text="",
            font=("Arial", 18, "bold"),
            anchor="w",
            justify="left"
        )
        self.show_task_content()

    def show_task_content(self):
        # Compose all tasks' info here
        task1_text = ""
        pending_numbers = []
        pending_count_last_month = 0
        report_date_str = "-"
        mth_c_value = "12-Month Arrears"
        pg = self.selected_pg_rqs.get()
        if "part_4" in self.excel_data and pg:
            df_part4 = self.excel_data["part_4"]
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
        if "part_1" in self.excel_data and pg:
            arrears_df = self.excel_data["part_1"]
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
        if "part_2" in self.excel_data and pg:
            df_part2 = self.excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()  # <-- add .copy()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce")
            total_outstanding = df_pg_part2["IM_AM"].sum()
            crdtcard_outstanding = df_pg_part2.loc[df_pg_part2["FCY_TYPE"] == "CRDTCARD", "IM_AM"].sum()
            if total_outstanding > 0:
                ratio = crdtcard_outstanding / total_outstanding
                task2_result = f"{crdtcard_outstanding:,.2f} / {total_outstanding:,.2f} = {ratio:.2%}"
            else:
                task2_result = "No outstanding found"
            
        task2_text = f"Task 2:\nCRDTCARD Outstanding Ratio: {task2_result}\n"
        task3_text = "Task 3:\n(Your logic here)\n"
        task4_text = "Task 4:\n(Your logic here)\n"
        all_tasks_text = f"{task1_text}\n{task2_text}\n{task3_text}\n{task4_text}"
        self.task_content_label.configure(text=all_tasks_text)
        self.task_tab_frame.pack_forget()
        self.task_tab_frame.pack(fill="x", pady=(0, 0), side="top")
        self.animate_task_tab(0, 10, duration=250)
        self.task_content_label.pack(fill="x", pady=(5, 10), anchor="w")
        self.close_btn.pack(side="right", padx=(4, 8), pady=0)

    def animate_task_tab(self, start, end, duration=250):
        # Animate pady from start to end over duration ms with ease-in
        steps = 20
        delta = end - start
        delay = duration // steps

        def ease_in(t):
            # Quadratic ease-in: t in [0,1]
            return t * t

        def step(i=0):
            t = i / steps
            value = int(start + delta * ease_in(t))
            self.task_tab_frame.pack_configure(pady=(value, 0))
            if i < steps:
                self.frame.after(delay, lambda: step(i + 1))
        step()
        
    def hide_task_content(self):
        # Animate shrink to bottom, just hide the label and close button (do not delete data)
        def after_hide():
            self.task_content_label.pack_forget()
            self.close_btn.pack_forget()
            self.task_tab_frame.pack_forget()
            self.task_tab_frame.pack(fill="x", pady=(0, 0), side="bottom")
        self.animate_task_tab(10, 0, duration=250)
        self.frame.after(250, after_hide)

    def show(self):
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
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

        self.show_task_content()
        
    def clear_table(self, tree):
        for row in tree.get_children():
            tree.delete(row)

    def create_table(self, parent, columns, height=5):
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=height)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center", width=120)
        tree.pack(fill="x", pady=5)
        return tree

    def update_table_height(self, tree, data_len, min_height=4, max_height=20):
        tree.config(height=max(min_height, min(data_len, max_height)))

    def set_treeview_style(self, mode):
        style = ttk.Style()
        if mode == "dark":
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


# --- Excel All Task Class (placeholder for your calculation display) ---
class ExcelAllTask:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ctk.CTkFrame(parent)
        ctk.CTkLabel(self.frame, text="Excel All Task - Calculation Display", font=("Arial", 18, "bold")).pack(pady=40)
        # Add your calculation widgets here
        self.frame.pack_forget()

    def show(self):
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()

# --- Main Content Area ---
main_content = ctk.CTkFrame(app)
main_content.pack(side="left", fill="both", expand=True)

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