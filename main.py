import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from PIL import Image, ImageTk, ImageDraw
import datetime
import threading
import subprocess
import sys
import os
import psutil
import win32gui
import win32con
from openpyxl import Workbook
from tkinter import ttk, messagebox


ctk.set_default_color_theme("Themes/extreme.json")
ctk.set_appearance_mode("dark")
 
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
hamburger_img = ctk.CTkImage(Image.open("Picture/hamburger.png"), size=(24, 24))
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
    fg_color="transparent",
    bg_color="transparent",
    corner_radius=10,
    border_width=2,
    
)
btn_report.pack(pady=10, padx=(0, 0))

btn_another = ctk.CTkButton(
    menu_frame,
    text="Excel All Task",
    width=150,
    height=40,
    fg_color="transparent",
    bg_color="transparent",
    font=("Arial", 15, "bold"),
    corner_radius=10,
    border_width=2,
)
btn_another.pack(pady=10, padx=(0, 0))

def is_integrate_running():
    """
    Check if a process running integrate.py exists.
    Returns the process if found, otherwise None.
    """
    for proc in psutil.process_iter(attrs=["cmdline"]):
        try:
            cmdline = proc.info["cmdline"]
            if cmdline and "integrate.py" in " ".join(cmdline):
                return proc
        except Exception:
            continue
    return None

def bring_integrate_to_front():
    """
    Brings the window with title containing 'Report Launcher' (adjust if needed)
    to the front.
    """
    def enum_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "Report Launcher" in title:
                results.append(hwnd)
    hwnds = []
    win32gui.EnumWindows(enum_callback, hwnds)
    if hwnds:
        h = hwnds[0]
        # Restore the window (if minimized) and bring to foreground
        win32gui.ShowWindow(h, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(h)

def back_to_main():
    script_path = os.path.join(os.getcwd(), "integrate.py")
    proc = is_integrate_running()
    if proc is None:
        subprocess.Popen([sys.executable, script_path])
    else:
        bring_integrate_to_front()

btn_main = ctk.CTkButton(
    menu_frame,
    text="🔙 Back to Main",
    width=150,
    height=40,
    command=back_to_main,  # Pass the function reference without parentheses
    fg_color="transparent",
    bg_color="transparent",
    font=("Arial", 15, "bold"),
    corner_radius=10,
    border_width=2,
)
btn_main.pack(pady=10, padx=(0, 0))

# --- Start with sidebar expanded ---
sidebar.configure(width=SIDEBAR_EXPANDED_WIDTH)
menu_frame.pack(fill="both", expand=True)

# At the very bottom of the sidebar, add a container for the toggle and settings buttons:
sidebar_bottom = ctk.CTkFrame(sidebar, fg_color="transparent")
sidebar_bottom.pack(side="bottom", fill="x", pady=10)

# Create an inner frame to hold both buttons side by side
toggle_setting_frame = ctk.CTkFrame(sidebar_bottom, fg_color="transparent")
toggle_setting_frame.pack(fill="x", pady=5)

# Load icons (if not already loaded)
try:
    dark_icon = ctk.CTkImage(Image.open("Picture/dark_mode_icon.png"), size=(24, 24))
    light_icon = ctk.CTkImage(Image.open("Picture/light_mode_icon.png"), size=(24, 24))
except Exception as e:
    print(f"Error loading toggle icons: {e}")
    dark_icon = None
    light_icon = None


# Set initial mode tracker
current_mode = {"mode": "dark"}

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
        
        
def toggle_sidebar_mode():
    # Toggle between dark and light modes using patina theme settings
    if current_mode["mode"] == "dark":
        ctk.set_appearance_mode("light")
        mode_toggle_btn.configure(image=light_icon)
        current_mode["mode"] = "light"
    else:
        ctk.set_appearance_mode("dark")
        mode_toggle_btn.configure(image=dark_icon)
        current_mode["mode"] = "dark"

# Create the dark/light mode toggle button
mode_toggle_btn = ctk.CTkButton(
    toggle_setting_frame,
    text="",
    image=dark_icon,  # initially dark mode icon
    width=40,
    height=40,
    fg_color="transparent",
    hover_color="#444",
    command=toggle_sidebar_mode
)
mode_toggle_btn.pack(side="left", expand=True, padx=5)


# Placeholder for button commands
def do_nothing():
    pass
 
# --- CCRIS Report Class ---
class CCRISReport:
    def __init__(self, parent):
        self.parent = parent
        
        
        # Set Treeview style to dark before creating any Treeview
        self.set_treeview_style("dark")
        self.set_treeview_style(ctk.get_appearance_mode())
        
        # --- Scrollable Frame Setup ---
        self.outer_frame = ctk.CTkFrame(parent)
        self.outer_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.task_tab = TaskTabBar(self.outer_frame)
        
        # --- Loading Overlay ---
        self.loading_gif = Image.open("Picture/loading.gif")
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

        # Header
        self.header = ctk.CTkFrame(self.frame)
        self.header.pack(fill="x", pady=10)
        logo_image = ctk.CTkImage(Image.open("Picture/bnm_logo.png"), size=(300, 56))
        alrajhi_logo_image = ctk.CTkImage(Image.open("Picture/alrajhi_logo.png"), size=(170, 60))
        self.header.columnconfigure((0, 1, 2), weight=1)
        ctk.CTkLabel(self.header, image=logo_image, text="").grid(row=0, column=1, padx=10, pady=5, sticky="nsew")
        ctk.CTkLabel(self.header, text="CREDIT REPORT", font=("Arial", 22, "bold")).grid(row=0, column=0, padx=(20, 10), sticky="w")
        ctk.CTkLabel(self.header, image=alrajhi_logo_image, text="").grid(row=0, column=2, padx=10, pady=5, sticky="nsew")
 
        # Replace your entire "Controls" section with the following grid-only layout:

        self.control_frame = ctk.CTkFrame(self.frame)
        self.control_frame.pack(fill="x", pady=5)  # use pack to place the frame in self.frame

        # Configure columns of the control_frame (you can adjust weights as needed)
        self.control_frame.grid_columnconfigure(0, weight=1)  # Import button column
        self.control_frame.grid_columnconfigure(1, weight=2)  # Navigation controls column
        self.control_frame.grid_columnconfigure(2, weight=1)  # (optional extra spacer)

        import_icon = ctk.CTkImage(Image.open("Picture/importing.png"), size=(24, 24))
        
        # Import button (placed at left)
        self.import_button = ctk.CTkButton(self.control_frame, text="Import CCRIS Excel", image=import_icon ,command=self.load_excel)
        self.import_button.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        # Create a navigation subframe to hold previous, combobox, and next buttons centered
        nav_frame = ctk.CTkFrame(self.control_frame)
        nav_frame.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Configure three columns inside nav_frame so controls can be centered
        nav_frame.grid_columnconfigure(0, weight=1)  # left spacer
        nav_frame.grid_columnconfigure(1, weight=0)  # center control column
        nav_frame.grid_columnconfigure(2, weight=1)  # right spacer

        # Load arrow icons
        left_arrow_icon = ctk.CTkImage(Image.open("Picture/left-arrow.png"), size=(24, 24))
        right_arrow_icon = ctk.CTkImage(Image.open("Picture/right-arrow.png"), size=(24, 24))

        # Previous Button in left column (aligned to right)
        self.prev_btn = ctk.CTkButton(
            nav_frame,
            text="",
            image=left_arrow_icon,
            fg_color="transparent",
            hover_color="#444",
            command=self.on_previous
        )
        self.prev_btn.grid(row=0, column=0, padx=10, pady=5, sticky="e")

        # Combobox in center column
        self.selected_pg_rqs = ctk.StringVar()
        style = ttk.Style()
        style.configure("Custom.TCombobox", font=("Consolas", 15, "bold"))
        self.pg_dropdown = ttk.Combobox(nav_frame,
                                        textvariable=self.selected_pg_rqs,
                                        width=25,
                                        style="Custom.TCombobox")
        self.pg_dropdown.grid(row=0, column=1, padx=10, pady=5)
        self.pg_dropdown.bind("<<ComboboxSelected>>", lambda event: self.load_pg_data())

        self.next_btn = ctk.CTkButton(
            nav_frame,
            text="",
            image=right_arrow_icon,
            fg_color="transparent",
            hover_color="#444",
            command=self.on_next
        )
        self.next_btn.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        
        self.arrears_label = ctk.CTkLabel(self.control_frame, text="Arrears in 12 Months:")
        self.arrears_label.grid(row=0, column=2, padx=10, pady=5, sticky="e")

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
        ctk.CTkLabel(self.table_section, text="Outstanding Credit", font=("Consolas", 14, "bold")).pack(anchor="w")
        self.outstanding_tree = self.create_table(self.table_section, self.outstanding_cols, height=6)
 
        # Special Attention
        ctk.CTkLabel(self.table_section, text="Special Attention Account", font=("Consolas", 14, "bold")).pack(anchor="w")
        self.attention_tree = self.create_table(self.table_section, self.outstanding_cols, height=4)
 
        # Application for Credit
        ctk.CTkLabel(self.table_section, text="Application for Credit", font=("Consolas", 14, "bold")).pack(anchor="w")
        self.application_tree = self.create_table(self.table_section, self.outstanding_cols, height=4)
 
        # Data
        self.excel_data = {}
        self.pg_list = []
 
        # Hide by default (will be shown by sidebar button)
        self.frame.pack_forget()
    
    # Function to handle mouse wheel events
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # Bind mouse wheel when cursor enters canvas,
        # unbind it when cursor leaves
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def on_previous(self):
        # Example: set the combobox to the previous page (if available)
        current_index = self.pg_dropdown.current()
        if current_index > 0:
            self.pg_dropdown.current(current_index - 1)
            self.load_pg_data()

    def on_next(self):
        # Example: set the combobox to the next page (if available)
        current_index = self.pg_dropdown.current()
        if current_index < len(self.pg_dropdown['values']) - 1:
            self.pg_dropdown.current(current_index + 1)
            self.load_pg_data()
        
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
        if self.task_tab.visible:
            self.task_tab.show_content(self.excel_data, pg)
        
       
    def clear_table(self, tree):
        for row in tree.get_children():
            tree.delete(row)
 
    def create_table(self, parent, columns, height=5):
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="both", expand=True, pady=5)

        # Create the Treeview with the "Treeview" style.
        tree = ttk.Treeview(
            frame,
            columns=columns,
            show="headings",
            height=height,
            style="Treeview",
            selectmode="browse"
        )
        tree.pack(side="left", fill="both", expand=True)

        # Update column widths when the frame resizes
        def adjust_columns(event):
            total_width = event.width
            # Subtract a few pixels for padding if needed
            col_width = total_width // len(columns)
            for col in columns:
                tree.column(col, anchor="center", width=col_width, stretch=True)

        frame.bind("<Configure>", adjust_columns)

        # Set headings (initial settings; widths will be adjusted on resize)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center", width=120, stretch=True)
        return tree
 
    def update_table_height(self, tree, data_len, min_height=4, max_height=20):
        tree.config(height=max(min_height, min(data_len, max_height)))
 
    def set_treeview_style(self, mode):
        style = ttk.Style()
        # Force the treeview area to fill the widget even if empty:
        style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        
        # Get theme settings from CustomTkinter's ThemeManager:
        theme = ctk.ThemeManager.theme
        is_dark = 1 if mode.lower() == "dark" else 0
        
        style.configure("Treeview",
                        rowheight=theme["Treeview"].get("rowheight", 25),
                        font=("Consolas", 13),
                        background=theme["Treeview"]["background"][is_dark],
                        fieldbackground=theme["Treeview"]["background"][is_dark],
                        foreground=theme["Treeview"]["foreground"][is_dark])
        
        style.configure("Treeview.Heading",
                        font=("Consolas", 13),
                        background=theme["Treeview"]["heading_background"][is_dark],
                        foreground="#000000")  # Set header text to black
        
        style.map("Treeview",
                background=[("selected", theme["Treeview"]["selected_background"][is_dark])],
                foreground=[("selected", theme["Treeview"]["selected_foreground"][is_dark])])

class TaskTabBar:
    def __init__(self, parent):
        self.parent = parent
        self.visible = False
        self.animating = False
        self.min_height = 40
        self.default_height = 320
        self.max_height = 600
        self.current_height = self.default_height
        self.last_height = self.default_height
        self.resizing = False
        self.start_y = 0

        # --- Main panel frame ---
        self.frame = ctk.CTkFrame(parent, fg_color="#23272e", corner_radius=10)
        self.frame.pack(side="bottom", fill="x")
        self.frame.configure(height=self.default_height)
        self.frame.pack_propagate(False)

        # --- Grip bar for resizing ---
        self.grip = ctk.CTkFrame(self.frame, height=8, fg_color="#444", cursor="sb_v_double_arrow")
        self.grip.pack(fill="x", side="top")
        self.grip.bind("<ButtonPress-1>", self.start_resize)
        self.grip.bind("<B1-Motion>", self.perform_resize)
        self.grip.bind("<Double-Button-1>", self.toggle_minimize)

        # --- Tab bar (always visible) ---
        self.tab_bar = ctk.CTkFrame(self.frame, fg_color="#23272e")
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

        close_icon = ctk.CTkImage(Image.open("Picture/close.png"), size=(24, 24))
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

        # --- Content area (hidden by default) ---
        self.content_frame = ctk.CTkFrame(self.frame, fg_color="#1e1e1e")
        self.content_label = ctk.CTkLabel(
            self.content_frame,
            text="",
            font=("Consolas", 13),
            anchor="nw",
            justify="left",
            text_color="#d4d4d4"
        )

        # For data
        self.excel_data = None
        self.pg = None

        # Set initial height
        self.set_panel_height(self.default_height)
        self.hide_content()

    def set_data(self, excel_data, pg):
        self.excel_data = excel_data
        self.pg = pg

    def _on_show(self):
        if self.excel_data is not None and self.pg is not None:
            self.show_content(self.excel_data, self.pg)
            self.animate_panel_height(self.current_height, self.max_height, duration=200)  # Animate open

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

        # Task 2: CRDTCARD Outstanding
        task2_result = "-"
        task2_text = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["FCY_TYPE"] = df_pg_part2["FCY_TYPE"].astype(str).str.strip().str.upper()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce").fillna(0)
            df_pg_part2["IM_LIM_AM"] = pd.to_numeric(df_pg_part2["IM_LIM_AM"], errors="coerce").fillna(0)
            df_pg_part2 = df_pg_part2.reset_index(drop=True)

            # Find all CRDTCARD rows (these have the outstanding)
            crdtcard_rows = df_pg_part2[df_pg_part2["FCY_TYPE"] == "CRDTCARD"]
            crdtcard_outstanding = crdtcard_rows["IM_AM"].sum()

            # For each CRDTCARD row, find the nearest previous row with approval date and take its limit
            used_approval_dates = set()
            total_limit = 0
            for idx in crdtcard_rows.index:
                found_limit = 0
                found_date = None
                for prev_idx in range(idx, -1, -1):
                    appr_date = str(df_pg_part2.loc[prev_idx, "DT_APPL"]).strip()
                    if appr_date not in ["-", "", "NaN", "nan"] and pd.notna(appr_date):
                        found_date = appr_date
                        found_limit = df_pg_part2.loc[prev_idx, "IM_LIM_AM"]
                        break
                # Only add the limit if this approval date hasn't been used yet
                if found_date and found_date not in used_approval_dates:
                    total_limit += found_limit
                    used_approval_dates.add(found_date)

            if total_limit > 0:
                ratio = crdtcard_outstanding / total_limit
                task2_result = f"{ratio:.2%}"
            else:
                task2_result = "No outstanding found"
            task2_text = f"Task 2:\nCRDTCARD Outstanding : {task2_result}\n"
        
        task3_result = "-"
        task3_text = "Task 3:\nCCRISCard Age : -\n"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            # Get the earliest (min) date
            if not df_pg_part2["DT_APPL"].isna().all():
                earliest_date = df_pg_part2["DT_APPL"].min()
                # Get report date from part_4 if available
                report_date = "-"
                if "part_4" in excel_data:
                    df_part4 = excel_data["part_4"]
                    df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg].copy()
                    if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                        latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                        if pd.notna(latest_report_date):
                            report_date = latest_report_date.strftime("%d-%m-%Y")
                task3_result = f"earliest date : {earliest_date.strftime('%d-%m-%Y')}\nreport date : {report_date}"
                task3_text = f"Task 3:\n-CCRISCard Age :\nearliest date : {earliest_date.strftime('%d-%m-%Y')}\nreport date : {report_date}\n"
            

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
        
        task6_result = "-"
        task6_text = "Task 6:\nThin Ccris : -\n"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            # a. Calculate months between earliest date and report date
            earliest_date = df_pg_part2["DT_APPL"].min() if not df_pg_part2["DT_APPL"].isna().all() else None
            report_date = None
            if "part_4" in excel_data:
                df_part4 = excel_data["part_4"]
                df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg].copy()
                if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                    latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                    if pd.notna(latest_report_date):
                        report_date = latest_report_date
            # Calculate months difference
            months_diff = "-"
            if earliest_date is not None and report_date is not None:
                months_diff = (report_date.year - earliest_date.year) * 12 + (report_date.month - earliest_date.month)
            # b. Only 1 facility?
            only_one_facility = "No"
            if df_pg_part2["FCY_TYPE"].nunique() == 1:
                only_one_facility = "Yes"
            task6_result = f"a. Months: {months_diff}\nb. Only 1 facility: {only_one_facility}"
            task6_text = f"Task 6:\nTHIN CCRIS :\na. Months: {months_diff}\nb. Only 1 facility: {only_one_facility}\n"

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

        all_tasks_text = f"{task1_text}\n{task2_text}\n{task3_text}\n{task4_text}\n{task5_text}\n{task6_text}\n{task7_text}\n{task8_text}"
        self.content_label.configure(text=all_tasks_text)

        # Show content area if not visible
        if not self.visible:
            self.content_frame.pack(fill="both", expand=True, padx=0, pady=(0, 0), side="top")
            self.content_label.pack(fill="both", expand=True, padx=16, pady=12, side="top")
            self.visible = True
            self.set_panel_height(self.current_height)

    def hide_content(self):
        self.content_label.pack_forget()
        self.content_frame.pack_forget()
        self.visible = False
        self.animating = False
        self.animate_panel_height(self.current_height, self.min_height, duration=200)  # Animate close

    def set_panel_height(self, height):
        height = max(self.min_height, min(self.max_height, int(height)))
        self.current_height = height
        self.frame.configure(height=height)
        self.frame.pack_propagate(False)

    def animate_panel_height(self, start, end, duration=200):
        if int(start) == int(end):
            self.set_panel_height(end)
            return
        steps = max(6, int(abs(end - start) / 30))
        delay = max(15, int(duration / steps))
        delta = (end - start) / steps
        self.animating = True

        def step(i=0):
            if i < steps:
                self.set_panel_height(start + delta * i)
                self.frame.update_idletasks()
                self.frame.after(delay, lambda: step(i + 1))
            else:
                self.set_panel_height(end)
                self.animating = False

        step()

    def start_resize(self, event):
        self.resizing = True
        self.start_y = event.y_root
        self.orig_height = self.current_height

    def perform_resize(self, event):
        if self.resizing:
            delta = self.start_y - event.y_root
            new_height = self.orig_height + delta
            if int(new_height) != int(self.current_height):
                self.set_panel_height(new_height)

    def toggle_minimize(self, event):
        if self.current_height > self.min_height + 10:
            self.last_height = self.current_height
            self.animate_panel_height(self.current_height, self.min_height, duration=200)
        else:
            self.animate_panel_height(self.current_height, self.last_height, duration=200)

class ExcelAllTask:
    def __init__(self, parent):
        self.parent = parent
        self.search_var = tk.StringVar()
        self.frame = ctk.CTkFrame(parent, corner_radius=12)
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Header (with logos) ---
        self.header = ctk.CTkFrame(self.frame, fg_color="transparent")
        self.header.pack(fill="x", pady=(10, 0))
        logo_image = ctk.CTkImage(Image.open("Picture/bnm_logo.png"), size=(220, 40))
        alrajhi_logo_image = ctk.CTkImage(Image.open("Picture/alrajhi_logo.png"), size=(120, 40))
        ctk.CTkLabel(self.header, image=logo_image, text="").pack(side="left", padx=(10, 0))
        ctk.CTkLabel(self.header, image=alrajhi_logo_image, text="").pack(side="right", padx=(0, 10))

        # --- Controls Row ---
        self.control_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        self.control_frame.pack(fill="x", pady=(10, 0), padx=10)

        # Search bar with icon
        search_icon = ctk.CTkImage(Image.open("Picture/search.png"), size=(20, 20))
        left_arrow_icon = ctk.CTkImage(Image.open("Picture/left-arrow.png"), size=(24, 24))
        right_arrow_icon = ctk.CTkImage(Image.open("Picture/right-arrow.png"), size=(24, 24))

        search_entry_frame = ctk.CTkFrame(self.control_frame, fg_color="#ffffff", corner_radius=8)
        search_entry_frame.pack(side="left", padx=(0, 10))

        # Previous button
        self.prev_btn = ctk.CTkButton(
            search_entry_frame,
            text="",
            image=left_arrow_icon,
            width=36,
            height=36,
            fg_color="transparent",
            hover_color="#eee",
            command=self.on_prev_match
        )
        self.prev_btn.pack(side="left", padx=(2, 2), pady=2)

        ctk.CTkLabel(search_entry_frame, image=search_icon, text="", fg_color="transparent").pack(side="left", padx=(2, 2))
        search_entry = ctk.CTkEntry(
            search_entry_frame, textvariable=self.search_var, width=180, fg_color="#222222", border_width=0, font=("Segoe UI", 14)
        )
        search_entry.pack(side="left", padx=(0, 2), pady=4)
        search_entry.bind("<Return>", self.on_search)
        search_entry.bind("<KeyRelease>", lambda e: None)

        # Next button
        self.next_btn = ctk.CTkButton(
            search_entry_frame,
            text="",
            image=right_arrow_icon,
            width=36,
            height=36,
            fg_color="transparent",
            hover_color="#eee",
            command=self.on_next_match
        )
        self.next_btn.pack(side="left", padx=(2, 2), pady=2)

        # Export button
        export_icon = ctk.CTkImage(Image.open("Picture/export.png"), size=(20, 20))
        self.export_button = ctk.CTkButton(
            self.control_frame,
            text="Export",
            image=export_icon,
            width=120,
            height=36,
            font=("Segoe UI", 14, "bold"),
            fg_color="#1976d2",
            hover_color="#1565c0",
            text_color="#fff",
            border_width=0,
            corner_radius=8,
            command=self.export_data
        )
        self.export_button.pack(side="right", padx=(10, 0))

        # --- Search navigation state ---
        self.matching_row_ids = []
        self.match_index = 0

        # --- Animated GIF Loading (rotating and small) ---
        self.loading_gif = Image.open("Picture/loading.gif")
        self.loading_frames = []
        size = (48, 48)
        for angle in range(0, 360, 30):
            frame = self.loading_gif.copy().resize(size, Image.LANCZOS).convert("RGBA")
            rotated = frame.rotate(angle)
            self.loading_frames.append(ctk.CTkImage(rotated, size=size))

        self.loading_label = ctk.CTkLabel(self.frame, text="")
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label.lower()
        self.loading_gif_running = False

        # --- Table Section ---
        self.columns = ["PG_RQS"] + [f"Task {i}" for i in range(1, 9)]
        table_frame = ctk.CTkFrame(self.frame, corner_radius=8)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        style = ttk.Style()
        style.configure("Modern.Treeview",
                        font=("Consolas", 13),
                        rowheight=28,
                        background="#23272e",
                        fieldbackground="#23272e",
                        foreground="#ffffff")
        style.configure("Modern.Treeview.Heading",
                        font=("Consolas", 13),
                        background="#1976d2",
                        foreground="#000000")  # <-- Set header text to black
        style.layout("Modern.Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

        self.tree = ttk.Treeview(
            table_frame,
            columns=self.columns,
            show="headings",
            height=18,
            style="Modern.Treeview",
            selectmode="browse"
        )
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=160, stretch=True)
        self.tree.pack(side="left", fill="both", expand=True)

        # Add vertical scrollbar
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        yscroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=yscroll.set)

    def export_data(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Export"
        )
        if not file_path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Tasks Export"
        ws.append(self.columns)
        for row_id in self.tree.get_children():
            row_data = self.tree.item(row_id)['values']
            ws.append(row_data)
        wb.save(file_path)
        messagebox.showinfo("Export", "Export completed successfully!")

    def on_search(self, event=None):
        query = self.search_var.get().lower()
        self.matching_row_ids = []
        self.match_index = 0
        if not query:
            self.tree.selection_remove(self.tree.selection())
            return
        # Find all matching rows
        for row_id in self.tree.get_children():
            row_data = self.tree.item(row_id)['values']
            if any(query in str(val).lower() for val in row_data):
                self.matching_row_ids.append(row_id)
        if self.matching_row_ids:
            self.highlight_match(0)
        else:
            self.tree.selection_remove(self.tree.selection())

    def highlight_match(self, idx):
        # Remove previous selection
        self.tree.selection_remove(self.tree.selection())
        if not self.matching_row_ids:
            return
        idx = idx % len(self.matching_row_ids)
        row_id = self.matching_row_ids[idx]
        self.tree.selection_set(row_id)
        self.tree.focus(row_id)
        self.tree.see(row_id)
        self.match_index = idx

    def on_next_match(self):
        if self.matching_row_ids:
            next_idx = (self.match_index + 1) % len(self.matching_row_ids)
            self.highlight_match(next_idx)

    def on_prev_match(self):
        if self.matching_row_ids:
            prev_idx = (self.match_index - 1) % len(self.matching_row_ids)
            self.highlight_match(prev_idx)

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
        threading.Thread(target=self.populate_grid, daemon=True).start()

    def hide(self):
        self.frame.pack_forget()

    def populate_grid(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        global ccris_report
        excel_data = ccris_report.excel_data
        if "part_1" not in excel_data:
            self.hide_loading()
            return
        pg_list = pd.Series(excel_data["part_1"]["PG_RQS"].unique()).dropna().tolist()
        batch_size = 100
        max_rows = 500
        for start in range(0, min(max_rows, len(pg_list)), batch_size):
            end = min(start + batch_size, len(pg_list))
            for i in range(start, end):
                pg = pg_list[i]
                task_summaries = self.get_task_summaries_for_pg(pg, excel_data)
                self.tree.insert("", "end", values=[pg] + task_summaries)
            self.frame.update_idletasks()
        self.hide_loading()

    def get_latest_report_date(self, df_pg_part4, df_pg_part2):
        if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
            date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], format="%d/%m/%Y", errors="coerce").max()
            if pd.notna(date):
                return date
        if not df_pg_part2.empty and "DT_APPL" in df_pg_part2.columns:
            date = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce").max()
            if pd.notna(date):
                return date
        return None

    def get_task_summaries_for_pg(self, pg, excel_data):
        df_part2 = excel_data.get("part_2", pd.DataFrame())
        df_part4 = excel_data.get("part_4", pd.DataFrame())
        df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy() if not df_part2.empty else pd.DataFrame()
        df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg].copy() if not df_part4.empty else pd.DataFrame()

        # Task 1: Pending applications in last One month
        task1 = "-"
        pending_count_last_month = 0
        pending_numbers_str = "-"
        report_date_str = "-"
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
                    if pending_numbers:
                        pending_numbers_str = ", ".join(str(n) for n in pending_numbers)
        task1 = f"{pending_count_last_month} (Rows: {pending_numbers_str}, Report: {report_date_str})"

        # Task 2: CRDTCARD Outstanding (new logic)
        task2 = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["FCY_TYPE"] = df_pg_part2["FCY_TYPE"].astype(str).str.strip().str.upper()
            df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce").fillna(0)
            df_pg_part2["IM_LIM_AM"] = pd.to_numeric(df_pg_part2["IM_LIM_AM"], errors="coerce").fillna(0)
            df_pg_part2 = df_pg_part2.reset_index(drop=True)

            # Find all CRDTCARD rows (these have the outstanding)
            crdtcard_rows = df_pg_part2[df_pg_part2["FCY_TYPE"] == "CRDTCARD"]
            crdtcard_outstanding = crdtcard_rows["IM_AM"].sum()

            # For each CRDTCARD row, find the nearest previous row with approval date and take its limit
            used_approval_dates = set()
            total_limit = 0
            for idx in crdtcard_rows.index:
                found_limit = 0
                found_date = None
                for prev_idx in range(idx, -1, -1):
                    appr_date = str(df_pg_part2.loc[prev_idx, "DT_APPL"]).strip()
                    if appr_date not in ["-", "", "NaN", "nan"] and pd.notna(appr_date):
                        found_date = appr_date
                        found_limit = df_pg_part2.loc[prev_idx, "IM_LIM_AM"]
                        break
                # Only add the limit if this approval date hasn't been used yet
                if found_date and found_date not in used_approval_dates:
                    total_limit += found_limit
                    used_approval_dates.add(found_date)

            if total_limit > 0:
                ratio = crdtcard_outstanding / total_limit
                task2 = f"{ratio:.2%}"
            else:
                task2 = "No outstanding found"
                
        # Task 3: Oldest approval date for this PG and report date for this PG (like TaskTabBar)
        task3 = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            if not df_pg_part2["DT_APPL"].isna().all():
                earliest_date = df_pg_part2["DT_APPL"].min()
                # Get report date from part_4 if available
                report_date = "-"
                if "part_4" in excel_data:
                    df_part4 = excel_data["part_4"]
                    df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg].copy()
                    if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                        latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                        if pd.notna(latest_report_date):
                            report_date = latest_report_date.strftime("%d-%m-%Y")
                task3 = f"earliest: {earliest_date.strftime('%d-%m-%Y')}, report: {report_date}"
        

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

        # Task 6: Thin CCRIS logic
        task6 = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["PG_RQS"] == pg].copy()
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            # a. Calculate months between earliest date and report date
            earliest_date = df_pg_part2["DT_APPL"].min() if not df_pg_part2["DT_APPL"].isna().all() else None
            # Get report date from part_4 if available, else use latest DT_APPL
            report_date = None
            if "part_4" in excel_data:
                df_part4 = excel_data["part_4"]
                df_pg_part4 = df_part4[df_part4["PG_RQS"] == pg].copy()
                if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                    latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                    if pd.notna(latest_report_date):
                        report_date = latest_report_date
            if report_date is None and not df_pg_part2["DT_APPL"].isna().all():
                report_date = df_pg_part2["DT_APPL"].max()
            # Calculate months difference
            months_diff = "-"
            if earliest_date is not None and report_date is not None:
                months_diff = (report_date.year - earliest_date.year) * 12 + (report_date.month - earliest_date.month)
            # b. Only 1 facility?
            only_one_facility = "No"
            if df_pg_part2["FCY_TYPE"].nunique() == 1:
                only_one_facility = "Yes"
            task6 = f"a. Months: {months_diff}, b. Only 1 facility: {only_one_facility}"

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
main_content.pack(side="left", fill="both", expand=True, padx=(18, 0))  # Adjust padding as needed

# Instantiate pages (now that main_content is defined)
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
