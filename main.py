import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from PIL import Image, ImageTk, ImageDraw, ImageSequence
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
    """
    Create a hamburger menu icon image.
    
    Args:
        size (int): Size of the icon in pixels (default: 32)
        color (str): Color of the menu lines (default: "#bbb")
        
    Returns:
        ctk.CTkImage: Custom tkinter image object for the menu icon
    """
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    x = size // 2
    for y in [size // 4, size // 2, 3 * size // 4]:
        draw.line([(x - 6, y), (x + 6, y)], fill=color, width=3)
    return ctk.CTkImage(img, size=(size, size))
 
menu_icon = create_menu_icon(32, "#bbb")

tab_ccris_icon = ctk.CTkImage(Image.open("Picture/tab_ccris.png"), size=(24, 24))
summary_icon = ctk.CTkImage(Image.open("Picture/summary.png"), size=(24, 24))
back_to_main_icon = ctk.CTkImage(Image.open("Picture/back_to_main.png"), size=(24, 24))

 
# --- Sidebar Setup ---
SIDEBAR_EXPANDED_WIDTH = 175
SIDEBAR_SHRUNK_WIDTH = 48

sidebar_container = ctk.CTkFrame(app, fg_color="transparent")
sidebar_container.pack(side="left", fill="y")

sidebar = ctk.CTkFrame(sidebar_container, width=SIDEBAR_EXPANDED_WIDTH)
sidebar.pack(side="left", fill="y")
sidebar.pack_propagate(False)

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

menu_label = ctk.CTkLabel(menu_frame, text="Menu", font=("Arial", 18, "bold"))
menu_label.pack(pady=(12, 10))

btn_report = ctk.CTkButton(
    menu_frame,
    text="CCRIS Report",
    width=150,
    height=40,
    font=("Arial", 15, "bold"),
    fg_color="transparent",
    bg_color="transparent",
    corner_radius=10,
    border_width=1,
    anchor="w"
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
    border_width=1,
    anchor="w"
)
btn_another.pack(pady=10, padx=(0, 0))

def is_integrate_running():
    """
    Check if the integrate.exe process is currently running.
    
    Returns:
        psutil.Process: The integrate.exe process if found, None otherwise
    """
    for proc in psutil.process_iter(attrs=["cmdline"]):
        try:
            cmdline = proc.info["cmdline"]
            if cmdline and "integrate.exe" in " ".join(cmdline):
                return proc
        except Exception:
            continue
    return None

def bring_integrate_to_front():
    """
    Brings the Report Launcher window to the foreground if it exists.
    
    Uses Windows API to find and activate the window with "Report Launcher" in its title.
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
    """
    Return to the main Report Launcher application.
    
    Launches integrate.exe if not running, otherwise brings the existing window to front.
    """
    exe_path = os.path.join(os.getcwd(), "integrate.exe")
    proc = is_integrate_running()
    if proc is None:
        subprocess.Popen([exe_path])
    else:
        bring_integrate_to_front()

btn_main = ctk.CTkButton(
    menu_frame,
    text="Back to Main",
    width=150,
    height=40,
    command=back_to_main,
    fg_color="transparent",
    bg_color="transparent",
    font=("Arial", 15, "bold"),
    corner_radius=10,
    border_width=1,
    anchor="w"
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
    """
    Toggle the sidebar between expanded and collapsed states.
    
    In collapsed state, shows icons only. In expanded state, shows text labels.
    Updates button appearance and sidebar width accordingly.
    """
    global sidebar_expanded
    if sidebar_expanded:
        sidebar.configure(width=SIDEBAR_SHRUNK_WIDTH)
        # Show only icons, hide text, keep positions
        btn_report.configure(text="", image=tab_ccris_icon, width=40, anchor="center", font=("Arial", 1))
        btn_another.configure(text="", image=summary_icon, width=40, anchor="center", font=("Arial", 1))
        btn_main.configure(text="", image=back_to_main_icon, width=40, anchor="center", font=("Arial", 1))
        menu_label.pack_forget()
        menu_label.pack(pady=(12, 10), before=btn_report)  # Always keep at the top
        sidebar_expanded = False
    else:
        sidebar.configure(width=SIDEBAR_EXPANDED_WIDTH)
        # Show only text, hide icons, keep positions
        btn_report.configure(text="CCRIS Report", image=None, width=150, anchor="w", font=("Arial", 15, "bold"))
        btn_another.configure(text="Excel All Task", image=None, width=150, anchor="w", font=("Arial", 15, "bold"))
        btn_main.configure(text="Back to Main", image=None, width=150, anchor="w", font=("Arial", 15, "bold"))
        menu_label.pack_forget()
        menu_label.pack(pady=(12, 10), before=btn_report)  # Always keep at the top
        sidebar_expanded = True
      
        
def toggle_sidebar_mode():
    """
    Toggle between dark and light appearance modes.
    
    Switches the entire application theme and updates the mode toggle button icon.
    """
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
    """Placeholder function for buttons that don't have specific actions yet."""
    pass
 
# --- CCRIS Report Class ---
class CCRISReport:
    """
    Main CCRIS Credit Report application interface.
    
    Handles Excel file import, data deduplication, customer navigation, 
    and displays credit information in tabular format with task calculations.
    
    Key Features:
    - Import and deduplicate CCRIS Excel data from 4 parts
    - Navigate between customers (NU_PTL values) 
    - Display Outstanding Credit, Special Attention, and Application tables
    - Show calculated task results in expandable panel
    - Preserve original NU_PTL values during processing
    """
    def __init__(self, parent):
        """
        Initialize the CCRIS Report interface.
        
        Args:
            parent: Parent widget to contain this interface
        """
        self.parent = parent
        self._repeat_job = None
        self._repeat_fast_job = None
        self._repeat_fast_timer = None
        
        # Set Treeview style to dark before creating any Treeview
        self.set_treeview_style("dark")
        self.set_treeview_style(ctk.get_appearance_mode())
        
        # --- Scrollable Frame Setup ---
        self.outer_frame = ctk.CTkFrame(parent)
        self.outer_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.loading_label = ctk.CTkLabel(self.outer_frame, text="", fg_color="#141414")
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label.lower()
        self.loading_gif_running = False
        
        self.task_tab = TaskTabBar(self.outer_frame)
        
        # --- Loading Overlay ---
        self.loading_gif = Image.open("Picture/loading.gif")
        self.loading_frames = []
        size = (150, 150)  # or (48, 48) for ExcelAllTask

        for frame in ImageSequence.Iterator(self.loading_gif):
            rgba = frame.convert("RGBA").resize(size, Image.LANCZOS)
            self.loading_frames.append(ctk.CTkImage(rgba, size=size))

        
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
        )
        self.prev_btn.grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.prev_btn.bind("<ButtonPress-1>", self._start_prev_repeat)
        self.prev_btn.bind("<ButtonRelease-1>", self._stop_prev_repeat)

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
        )
        self.next_btn.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.next_btn.bind("<ButtonPress-1>", self._start_next_repeat)
        self.next_btn.bind("<ButtonRelease-1>", self._stop_next_repeat)
        
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
    

        # Function to handle mouse wheel events for the canvas
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        # Use this function for all Treeviews:
        def bind_treeview_mousewheel_with_passthrough(tree, canvas):
            def _on_mousewheel(event):
                first, last = tree.yview()
                direction = -1 if event.delta > 0 else 1
                # If Treeview can scroll in the direction, scroll it
                if (direction == -1 and first > 0) or (direction == 1 and last < 1):
                    tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return "break"
                else:
                    # Otherwise, scroll the canvas
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return "break"
            def _bind(event):
                tree.bind_all("<MouseWheel>", _on_mousewheel)
            def _unbind(event):
                tree.unbind_all("<MouseWheel>")
            tree.bind("<Enter>", _bind)
            tree.bind("<Leave>", _unbind)

        # After creating the Treeviews:
        bind_treeview_mousewheel_with_passthrough(self.outstanding_tree, self.canvas)
        bind_treeview_mousewheel_with_passthrough(self.attention_tree, self.canvas)
        bind_treeview_mousewheel_with_passthrough(self.application_tree, self.canvas)
        
    def _start_next_repeat(self, event=None):
        self._stop_next_repeat()
        def slow_repeat():
            self.on_next()
            self._repeat_job = self.outer_frame.after(400, slow_repeat)
        slow_repeat()
        # After 5 seconds, switch to fast repeat
        self._repeat_fast_timer = self.outer_frame.after(3000, self._switch_to_fast_next_repeat)

    def _switch_to_fast_next_repeat(self):
        self._stop_next_repeat()
        def fast_repeat():
            self.on_next()
            self._repeat_fast_job = self.outer_frame.after(80, fast_repeat)
        fast_repeat()

    def _stop_next_repeat(self, event=None):
        if self._repeat_job:
            self.outer_frame.after_cancel(self._repeat_job)
            self._repeat_job = None
        if self._repeat_fast_job:
            self.outer_frame.after_cancel(self._repeat_fast_job)
            self._repeat_fast_job = None
        if self._repeat_fast_timer:
            self.outer_frame.after_cancel(self._repeat_fast_timer)
            self._repeat_fast_timer = None

    def _start_prev_repeat(self, event=None):
        self._stop_prev_repeat()
        def slow_repeat():
            self.on_previous()
            self._repeat_job = self.outer_frame.after(400, slow_repeat)
        slow_repeat()
        self._repeat_fast_timer = self.outer_frame.after(5000, self._switch_to_fast_prev_repeat)

    def _switch_to_fast_prev_repeat(self):
        self._stop_prev_repeat()
        def fast_repeat():
            self.on_previous()
            self._repeat_fast_job = self.outer_frame.after(80, fast_repeat)
        fast_repeat()

    def _stop_prev_repeat(self, event=None):
        if self._repeat_job:
            self.outer_frame.after_cancel(self._repeat_job)
            self._repeat_job = None
        if self._repeat_fast_job:
            self.outer_frame.after_cancel(self._repeat_fast_job)
            self._repeat_fast_job = None
        if self._repeat_fast_timer:
            self.outer_frame.after_cancel(self._repeat_fast_timer)
            self._repeat_fast_timer = None

    def on_previous(self):
        """Navigate to the previous customer in the dropdown list."""
        current_index = self.pg_dropdown.current()
        if current_index > 0:
            self.pg_dropdown.current(current_index - 1)
            self.load_pg_data()

    def on_next(self):
        """Navigate to the next customer in the dropdown list."""
        current_index = self.pg_dropdown.current()
        if current_index < len(self.pg_dropdown['values']) - 1:
            self.pg_dropdown.current(current_index + 1)
            self.load_pg_data()
        
    def show_loading(self):
        """Display animated loading GIF overlay during data processing."""
        self.loading_label.lift()
        self.loading_gif_running = True
        self.animate_loading_gif(0)

    def hide_loading(self):
        """Hide the loading GIF overlay."""
        self.loading_label.lower()
        self.loading_gif_running = False

    def animate_loading_gif(self, idx):
        """
        Animate the loading GIF by cycling through frames.
        
        Args:
            idx (int): Current frame index
        """
        if not self.loading_gif_running:
            return
        frame = self.loading_frames[idx]
        self.loading_label.configure(image=frame, text="")
        next_idx = (idx + 1) % len(self.loading_frames)
        self.outer_frame.after(60, lambda: self.animate_loading_gif(next_idx))

 
    def show(self):
        """Display the CCRIS Report interface."""
        self.outer_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def hide(self):
        """Hide the CCRIS Report interface."""
        self.outer_frame.pack_forget()
 
    def load_excel(self):
        """
        Import CCRIS Excel file and process the data.
        
        Opens file dialog to select Excel file, loads data from parts 1-4,
        applies deduplication logic, and populates the customer dropdown.
        Preserves original NU_PTL values throughout the process.
        """
        self.show_loading()
        def do_import():
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
            if not file_path:
                self.hide_loading()
                return
            xls = pd.ExcelFile(file_path)
            
            # Load raw data first
            raw_data = {}
            for sheet in ["part_1", "part_2", "part_3", "part_4"]:
                if sheet in xls.sheet_names:
                    df = xls.parse(sheet, dtype=str).fillna("NaN")
                    raw_data[sheet] = df

            # Apply deduplication logic
            self.excel_data = self.deduplicate_ccris_data(raw_data)

            if "part_1" in self.excel_data:
                print(f"CCRISReport - After deduplication: {len(self.excel_data['part_1'])} part_1 records")
                
                # Extract NU_PTL values directly without conversion
                deduplicated_part1 = self.excel_data["part_1"]
                raw_nu_ptl_values = deduplicated_part1["NU_PTL"].tolist()
                
                # Clean NU_PTL list while preserving original values
                self.pg_list = []
                for nu_ptl in raw_nu_ptl_values:
                    # Convert to string for validation but keep original if valid
                    nu_ptl_str = str(nu_ptl).strip()
                    
                    # Skip clearly invalid entries
                    if nu_ptl_str in ["", "NaN", "nan", "None", "none"] or pd.isna(nu_ptl):
                        continue
                    
                    # Use original value, not converted string
                    if nu_ptl not in self.pg_list:
                        self.pg_list.append(nu_ptl)
                
                # Sort numerically if possible, otherwise string sort
                try:
                    self.pg_list.sort(key=lambda x: float(str(x)) if str(x).replace('.','').isdigit() else float('inf'))
                except:
                    self.pg_list.sort(key=str)
                
                print(f"CCRISReport - Final NU_PTL list: {len(self.pg_list)} unique values")
                print(f"First 10 NU_PTL values: {[str(x) for x in self.pg_list[:10]]}")
                
                # Convert to strings for dropdown display
                self.pg_dropdown.configure(values=[str(x) for x in self.pg_list])
                if self.pg_list:
                    self.selected_pg_rqs.set(str(self.pg_list[0]))
                    self.load_pg_data()
            self.hide_loading()
        threading.Thread(target=do_import, daemon=True).start()

    def deduplicate_ccris_data(self, raw_data):
        """
        Deduplicate CCRIS data by keeping only the latest PG_RQS for each NU_PTL.
        
        Ensures proper REC_CTR sequencing for parts 2-4, while part 1 only needs
        latest PG_RQS per customer.
        
        Args:
            raw_data (dict): Raw data from Excel sheets by part name
            
        Returns:
            dict: Deduplicated data with proper sequencing
        """
        cleaned_data = {}
        
        for part_name, df in raw_data.items():
            if df.empty:
                cleaned_data[part_name] = df
                continue
                
            print(f"\n=== Processing {part_name} ===")
            print(f"Original data shape: {df.shape}")
            
            if part_name == "part_1":
                # Part 1: Just keep latest PG_RQS for each NU_PTL (no REC_CTR)
                cleaned_data[part_name] = self.deduplicate_part1(df)
            else:
                # Parts 2,3,4: Full deduplication with REC_CTR validation
                cleaned_data[part_name] = self.deduplicate_with_rec_ctr(df, part_name)
        
        return cleaned_data

    def deduplicate_part1(self, df):
        """
        Deduplicate part_1 by keeping latest PG_RQS for each NU_PTL.
        
        Preserves original NU_PTL values while ensuring only the most recent
        report data is kept for each customer.
        
        Args:
            df (pandas.DataFrame): Raw part_1 data
            
        Returns:
            pandas.DataFrame: Deduplicated data with original NU_PTL preserved
        """
        if "NU_PTL" not in df.columns or "PG_RQS" not in df.columns:
            print("Missing required columns for part_1 deduplication")
            return df
        
        print(f"Part 1 deduplication - Input: {len(df)} records")
        print(f"Part 1 unique NU_PTL before: {df['NU_PTL'].nunique()}")
        
        # Preserve original NU_PTL values
        df_clean = df.copy()
        
        # Clean NU_PTL - remove invalid entries but preserve original format
        df_clean = df_clean[df_clean['NU_PTL'].notna()]
        df_clean = df_clean[df_clean['NU_PTL'].astype(str).str.strip() != '']
        df_clean = df_clean[~df_clean['NU_PTL'].astype(str).str.lower().isin(['nan', 'none'])]
        
        # Convert PG_RQS to numeric for proper sorting only
        df_clean["PG_RQS_NUM"] = pd.to_numeric(df_clean["PG_RQS"], errors="coerce")
        df_clean = df_clean.dropna(subset=["PG_RQS_NUM"])
        
        print(f"Part 1 after cleaning invalid data: {len(df_clean)} records")
        
        # Sort by NU_PTL and PG_RQS (descending = latest first)
        df_clean = df_clean.sort_values(["NU_PTL", "PG_RQS_NUM"], ascending=[True, False])
        
        # Keep only the first (latest) record for each NU_PTL
        df_latest = df_clean.groupby("NU_PTL").first().reset_index()
        
        # Remove helper column but keep original NU_PTL
        df_latest = df_latest.drop("PG_RQS_NUM", axis=1)
        
        print(f"Part 1 - Original: {len(df)} → Latest: {len(df_latest)} records")
        print(f"Part 1 unique NU_PTL after: {df_latest['NU_PTL'].nunique()}")
        
        # Show actual NU_PTL values being kept
        if len(df_latest) > 0:
            print("Sample deduplicated NU_PTL values:")
            for i, nu_ptl in enumerate(df_latest['NU_PTL'].head(10)):
                pg_rqs = df_latest.iloc[i]['PG_RQS']
                print(f"  NU_PTL: {nu_ptl} (original), PG_RQS: {pg_rqs}")
        
        return df_latest

    def deduplicate_with_rec_ctr(self, df, part_name):
        """
        Deduplicate parts 2,3,4 with REC_CTR validation.
        
        Keeps only the latest PG_RQS for each NU_PTL and ensures REC_CTR
        sequences are valid (0,1,2,3...) for each customer.
        
        Args:
            df (pandas.DataFrame): Raw data for parts 2-4
            part_name (str): Name of the part being processed
            
        Returns:
            pandas.DataFrame: Deduplicated and validated data
        """
        required_cols = ["NU_PTL", "PG_RQS", "REC_CTR"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            print(f"Missing columns in {part_name}: {missing_cols}")
            return df
        
        df_clean = df.copy()
        
        # Convert to appropriate types
        df_clean["PG_RQS_NUM"] = pd.to_numeric(df_clean["PG_RQS"], errors="coerce")
        df_clean["REC_CTR_NUM"] = pd.to_numeric(df_clean["REC_CTR"], errors="coerce")
        
        # Remove invalid rows
        df_clean = df_clean.dropna(subset=["NU_PTL", "PG_RQS_NUM", "REC_CTR_NUM"])
        
        # Group by NU_PTL and find latest PG_RQS for each customer
        latest_pg_rqs = df_clean.groupby("NU_PTL")["PG_RQS_NUM"].max().reset_index()
        latest_pg_rqs.rename(columns={"PG_RQS_NUM": "LATEST_PG_RQS"}, inplace=True)
        
        # Merge to keep only latest PG_RQS records
        df_clean = df_clean.merge(latest_pg_rqs, on="NU_PTL")
        df_latest = df_clean[df_clean["PG_RQS_NUM"] == df_clean["LATEST_PG_RQS"]].copy()
        
        # Validate and fix REC_CTR sequencing
        df_validated = self.validate_rec_ctr_sequence(df_latest, part_name)
        
        # Clean up helper columns
        columns_to_drop = ["PG_RQS_NUM", "REC_CTR_NUM", "LATEST_PG_RQS"]
        df_validated = df_validated.drop([col for col in columns_to_drop if col in df_validated.columns], axis=1)
        
        print(f"{part_name} - Original: {len(df)} → Latest: {len(df_validated)} records")
        
        return df_validated

    def validate_rec_ctr_sequence(self, df, part_name):
        """
        Validate that REC_CTR is sequential (0,1,2,3...) for each NU_PTL.
        
        Fixes any sequence issues by reassigning REC_CTR values to ensure
        proper sequential order within each customer's records.
        
        Args:
            df (pandas.DataFrame): Data to validate
            part_name (str): Name of the part being validated
            
        Returns:
            pandas.DataFrame: Data with corrected REC_CTR sequences
        """
        validated_groups = []
        issues_found = 0
        
        for nu_ptl, group in df.groupby("NU_PTL"):
            # Sort by REC_CTR
            group_sorted = group.sort_values("REC_CTR_NUM").reset_index(drop=True)
            
            # Expected sequence: 0, 1, 2, 3, ...
            expected_sequence = list(range(len(group_sorted)))
            actual_sequence = group_sorted["REC_CTR_NUM"].astype(int).tolist()
            
            if actual_sequence != expected_sequence:
                issues_found += 1
                print(f"⚠️  {part_name} - NU_PTL {nu_ptl}: Expected {expected_sequence}, Got {actual_sequence}")
                
                # Fix the sequence by reassigning REC_CTR
                group_sorted["REC_CTR"] = expected_sequence
                group_sorted["REC_CTR_NUM"] = expected_sequence
                print(f"   ✅ Fixed sequence for NU_PTL {nu_ptl}")
            
            validated_groups.append(group_sorted)
        
        if issues_found == 0:
            print(f"✅ {part_name} - All REC_CTR sequences are valid")
        else:
            print(f"🔧 {part_name} - Fixed {issues_found} REC_CTR sequence issues")
        
        return pd.concat(validated_groups, ignore_index=True) if validated_groups else df

 
    def load_pg_data(self):
        """
        Load and display data for the selected customer (NU_PTL).
        
        Populates the three main tables (Outstanding Credit, Special Attention, 
        Application for Credit) with data for the currently selected customer.
        Also calculates and displays report date and arrears information.
        """
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
 
        # Calculate Report Date for part_4 (Application for Credit)
        report_date_str = "-"
        pending_count_last_month = 0
        if "part_4" in self.excel_data:
            df_part4 = self.excel_data["part_4"]
            df_pg_part4 = df_part4[df_part4["NU_PTL"] == pg]
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
 
        # Populate tables with formatted data
        for part, tree in zip(["part_2", "part_3", "part_4"], [self.outstanding_tree, self.attention_tree, self.application_tree]):
            df = self.excel_data.get(part, pd.DataFrame())
            df_pg = df[df["NU_PTL"] == pg]
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
 
        # Load MTH_C from part_1 for arrears display
        mth_c_value = "12-Month Arrears"
        if "part_1" in self.excel_data:
            arrears_df = self.excel_data["part_1"]
            mth_c = arrears_df.loc[arrears_df["NU_PTL"] == pg, "MTH_C"]
            if not mth_c.empty:
                mth_c_value = mth_c.iloc[0]
 
        # Update arrears label with report date
        self.arrears_label.configure(
            text=f"Arrears in 12 Months: {mth_c_value}    |    Report Date: {report_date_str}"
        )
        self.attention_tree.heading("12-Month Arrears", text=mth_c_value)
        self.application_tree.heading("12-Month Arrears", text=mth_c_value)
        self.outstanding_tree.heading("12-Month Arrears", text=mth_c_value)
 
        # Update task panel with current data
        self.task_tab.set_data(self.excel_data, pg)
        if self.task_tab.visible:
            self.task_tab.show_content(self.excel_data, pg)
        
       
    def clear_table(self, tree):
        """Remove all rows from a Treeview table."""
        for row in tree.get_children():
            tree.delete(row)
 
    def create_table(self, parent, columns, height=5):
        """
        Create a styled Treeview table with proper column configuration.
        
        Args:
            parent: Parent widget for the table
            columns (list): List of column names
            height (int): Initial table height in rows
            
        Returns:
            ttk.Treeview: Configured table widget
        """
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
        """
        Adjust table height based on data length within specified bounds.
        
        Args:
            tree: Treeview to update
            data_len (int): Number of data rows
            min_height (int): Minimum height in rows
            max_height (int): Maximum height in rows
        """
        tree.config(height=max(min_height, min(data_len, max_height)))
 
    def set_treeview_style(self, mode):
        """
        Configure Treeview styling for dark/light mode compatibility.
        
        Args:
            mode (str): "dark" or "light" appearance mode
        """
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
    """
    Resizable bottom panel that displays calculated task results.
    
    Shows 8 calculated tasks based on CCRIS business logic:
    1. Pending applications count
    2. Credit card utilization percentage  
    3. CCRIS age (earliest date)
    4-5. Unsecured facilities in 12/18 months
    6. Thin CCRIS analysis
    7-8. Secured/Unsecured financing counts and totals
    
    Features resizing, minimize/expand, and task-specific calculations.
    """
    def __init__(self, parent):
        """
        Initialize the resizable task panel.
        
        Args:
            parent: Parent widget to contain this panel
        """
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
        """
        Set the data for task calculations.
        
        Args:
            excel_data (dict): Processed CCRIS data by part
            pg (str): Customer identifier (NU_PTL)
        """
        self.excel_data = excel_data
        self.pg = pg

    def _on_show(self):
        """Show task content and animate panel expansion if data is available."""
        if self.excel_data is not None and self.pg is not None:
            self.show_content(self.excel_data, self.pg)
            self.animate_panel_height(self.current_height, self.max_height, duration=200)  # Animate open

    def show_content(self, excel_data, pg):
        """
        Calculate and display all 8 tasks for the specified customer.
        
        Performs the core CCRIS business logic calculations:
        - Task 1: Count pending applications in last month
        - Task 2: Calculate credit card utilization ratio
        - Task 3: Find earliest financing date  
        - Task 4-5: Count unsecured facilities in 12/18 months
        - Task 6: Analyze thin CCRIS (months + single facility check)
        - Task 7-8: Count and sum secured/unsecured financing by groups
        
        Args:
            excel_data (dict): Processed CCRIS data
            pg (str): Customer identifier
        """
        # Build calculation text
        task1_text = ""
        pending_numbers = []
        pending_count_last_month = 0
        report_date_str = "-"
        mth_c_value = "12-Month Arrears"
        if "part_4" in excel_data and pg:
            df_part4 = excel_data["part_4"]
            df_pg_part4 = df_part4[df_part4["NU_PTL"] == pg]
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
            mth_c = arrears_df.loc[arrears_df["NU_PTL"] == pg, "MTH_C"]
            if not mth_c.empty:
                mth_c_value = mth_c.iloc[0]
        if pending_numbers:
            pending_numbers_str = ", ".join(str(n) for n in pending_numbers)
        else:
            pending_numbers_str = "-"
        task1_text = (
            f"1. Number of pending applications in the last one month: {pending_count_last_month}\n"
            f"Row No : {pending_numbers_str}\n"
            f"Report Date: {report_date_str}\n"
        )

        # Task 2: Credit Card utilization calculation
        task2_result = "-"
        task2_text = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["NU_PTL"] == pg].copy()
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
            task2_text = f"2. Credit Card utilization:\nCRDTCARD Outstanding : {task2_result}\n"
        
        # Task 3: CCRIS Age calculation (earliest financing date)  
        task3_result = "-"
        task3_text = "Task 3:\nCCRISCard Age : -\n"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["NU_PTL"] == pg].copy()
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            # Get the earliest (min) date
            if not df_pg_part2["DT_APPL"].isna().all():
                earliest_date = df_pg_part2["DT_APPL"].min()
                # Get report date from part_4 if available
                report_date = "-"
                if "part_4" in excel_data:
                    df_part4 = excel_data["part_4"]
                    df_pg_part4 = df_part4[df_part4["NU_PTL"] == pg].copy()
                    if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                        latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                        if pd.notna(latest_report_date):
                            report_date = latest_report_date.strftime("%d-%m-%Y")
                task3_result = f"earliest date : {earliest_date.strftime('%d-%m-%Y')}\nreport date : {report_date}"
                task3_text = f"3. CCRIS Age:\nearliest date : {earliest_date.strftime('%d-%m-%Y')}\nreport date : {report_date}\n"
            

        # Task 4 & 5: Unsecured facilities in last 12/18 months
        # Task 4: Number of unsecured facilities in last 12 months
        task4_result = "-"
        task4_rows = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["NU_PTL"] == pg].copy()
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            df_pg_part2["TM_AGG_UTE"] = pd.to_datetime(df_pg_part2["TM_AGG_UTE"], errors="coerce")
            
            # Get report date from part_2 TM_AGG_UTE
            report_date = df_pg_part2["TM_AGG_UTE"].max()
            if pd.isna(report_date):
                report_date = df_pg_part2["DT_APPL"].max()
                
            if pd.notna(report_date):
                one_year_ago = report_date - pd.DateOffset(months=12)
                
                # Debug: Print data for troubleshooting
                print(f"DEBUG Task 4 - PG: {pg}")
                print(f"DEBUG - Report Date: {report_date}")
                print(f"DEBUG - One Year Ago: {one_year_ago}")
                print(f"DEBUG - Data shape: {df_pg_part2.shape}")
                if not df_pg_part2.empty:
                    print(f"DEBUG - COL_TYPE values: {df_pg_part2['COL_TYPE'].astype(str).str.strip().unique()}")
                    print(f"DEBUG - DT_APPL values: {df_pg_part2['DT_APPL'].tolist()}")
                    print(f"DEBUG - Sample rows:")
                    for idx, row in df_pg_part2.iterrows():
                        col_type = str(row['COL_TYPE']).strip()
                        dt_appl = row['DT_APPL']
                        is_unsecured = col_type in ["0", "-", "NaN", "nan"]
                        is_in_range = pd.notna(dt_appl) and dt_appl >= one_year_ago and dt_appl <= report_date
                        print(f"  Row {idx}: COL_TYPE='{col_type}', DT_APPL={dt_appl}, Unsecured={is_unsecured}, InRange={is_in_range}")
                
                # Modified mask to include 'NaN' and 'nan' as unsecured, and handle grouped data
                col_type_unsecured = df_pg_part2["COL_TYPE"].astype(str).str.strip().isin(["0", "-", "NaN", "nan"])
                date_in_range = (df_pg_part2["DT_APPL"] >= one_year_ago) & (df_pg_part2["DT_APPL"] <= report_date)
                
                # For cases where COL_TYPE and DT_APPL might be in different rows, 
                # check if any row in the customer data has both unsecured COL_TYPE and valid date
                unsecured_count = 0
                matching_rows = []
                
                # Group by REC_CTR to handle cases where data spans multiple rows
                if 'REC_CTR' in df_pg_part2.columns:
                    for rec_ctr, group in df_pg_part2.groupby('REC_CTR'):
                        has_unsecured = any(str(col).strip() in ["0", "-", "NaN", "nan"] for col in group['COL_TYPE'])
                        has_valid_date = any(pd.notna(dt) and dt >= one_year_ago and dt <= report_date for dt in group['DT_APPL'])
                        
                        if has_unsecured and has_valid_date:
                            unsecured_count += 1
                            matching_rows.append(str(rec_ctr))
                else:
                    # Fallback: use original row-based logic but include NaN
                    mask = col_type_unsecured & date_in_range
                    unsecured_rows = df_pg_part2[mask]
                    unsecured_count = len(unsecured_rows)
                    matching_rows = unsecured_rows["REC_CTR"].astype(str).tolist() if not unsecured_rows.empty else []
                
                task4_result = unsecured_count
                task4_rows = ", ".join(matching_rows) if matching_rows else "-"
                print(f"DEBUG - Task 4 Result: {task4_result}, Rows: {task4_rows}")
                print("=" * 50)
        task4_text = (
            f"4. Number of unsecured facilities in last 12 months: {task4_result}\n"
            f"Row No: {task4_rows}\n"
        )

        # Task 5: Number of unsecured facilities in last 18 months
        task5_result = "-"
        task5_rows = "-"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["NU_PTL"] == pg].copy()
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            df_pg_part2["TM_AGG_UTE"] = pd.to_datetime(df_pg_part2["TM_AGG_UTE"], errors="coerce")
            
            # Get report date from part_2 TM_AGG_UTE
            report_date = df_pg_part2["TM_AGG_UTE"].max()
            if pd.isna(report_date):
                report_date = df_pg_part2["DT_APPL"].max()
                
            if pd.notna(report_date):
                eighteen_months_ago = report_date - pd.DateOffset(months=18)
                
                # Debug: Print data for troubleshooting
                print(f"DEBUG Task 5 - PG: {pg}")
                print(f"DEBUG - Report Date: {report_date}")
                print(f"DEBUG - Eighteen Months Ago: {eighteen_months_ago}")
                
                # Modified mask to include 'NaN' and 'nan' as unsecured, and handle grouped data
                col_type_unsecured = df_pg_part2["COL_TYPE"].astype(str).str.strip().isin(["0", "-", "NaN", "nan"])
                date_in_range = (df_pg_part2["DT_APPL"] >= eighteen_months_ago) & (df_pg_part2["DT_APPL"] <= report_date)
                
                # For cases where COL_TYPE and DT_APPL might be in different rows, 
                # check if any row in the customer data has both unsecured COL_TYPE and valid date
                unsecured_count = 0
                matching_rows = []
                
                # Group by REC_CTR to handle cases where data spans multiple rows
                if 'REC_CTR' in df_pg_part2.columns:
                    for rec_ctr, group in df_pg_part2.groupby('REC_CTR'):
                        has_unsecured = any(str(col).strip() in ["0", "-", "NaN", "nan"] for col in group['COL_TYPE'])
                        has_valid_date = any(pd.notna(dt) and dt >= eighteen_months_ago and dt <= report_date for dt in group['DT_APPL'])
                        
                        if has_unsecured and has_valid_date:
                            unsecured_count += 1
                            matching_rows.append(str(rec_ctr))
                else:
                    # Fallback: use original row-based logic but include NaN
                    mask = col_type_unsecured & date_in_range
                    unsecured_rows = df_pg_part2[mask]
                    unsecured_count = len(unsecured_rows)
                    matching_rows = unsecured_rows["REC_CTR"].astype(str).tolist() if not unsecured_rows.empty else []
                
                task5_result = unsecured_count
                task5_rows = ", ".join(matching_rows) if matching_rows else "-"
                print(f"DEBUG - Task 5 Result: {task5_result}, Rows: {task5_rows}")
                print("=" * 50)
        task5_text = (
            f"5. Number of unsecured facilities in last 18 months: {task5_result}\n"
            f"Row No: {task5_rows}\n"
        )

        # Task 6: Thin CCRIS analysis (age and facility diversity)
        task6_result = "-"
        task6_text = "Task 6:\nThin Ccris : -\n"
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["NU_PTL"] == pg].copy()
            df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
            # a. Calculate months between earliest date and report date
            earliest_date = df_pg_part2["DT_APPL"].min() if not df_pg_part2["DT_APPL"].isna().all() else None
            report_date = None
            if "part_4" in excel_data:
                df_part4 = excel_data["part_4"]
                df_pg_part4 = df_part4[df_part4["NU_PTL"] == pg].copy()
                if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
                    latest_report_date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], errors="coerce").max()
                    if pd.notna(latest_report_date):
                        report_date = latest_report_date
            # Calculate months difference
            months_diff = "-"
            if earliest_date is not None and report_date is not None:
                months_diff = (report_date.year - earliest_date.year) * 12 + (report_date.month - earliest_date.month)
            # b. Only 1 facility?
            facilities = df_pg_part2["FCY_TYPE"].dropna().astype(str).str.strip()
            facilities = facilities[(facilities != "") & (facilities != "NaN") & (facilities != "-")]
            only_one_facility = "Yes" if facilities.nunique() == 1 else "No"
            task6_result = f"a. Months: {months_diff}\nb. Only 1 facility: {only_one_facility}"
            task6_text = f"6. Thin CCRIS:\na. Months: {months_diff}\nb. Only 1 facility: {only_one_facility}\n"

        # Task 7 & 8: Secured/Unsecured financing analysis by groups
        task7_count = 0
        task7_outstanding = 0.0
        task8_count = 0
        task8_outstanding = 0.0
        
        # Helper function to validate date values
        def is_valid_date(date_val):
            """Check if a value is a valid date (not '-', '', None, NaN)"""
            if pd.isna(date_val):
                return False
            date_str = str(date_val).strip()
            if date_str in ["-", "", "NaN", "nan", "None", "none"]:
                return False
            # Additional check: if it contains actual date characters
            if any(char.isdigit() for char in date_str) and ("-" in date_str or "/" in date_str):
                return True
            return False

        # Process secured/unsecured financing by groups
        if "part_2" in excel_data and pg:
            df_part2 = excel_data["part_2"]
            df_pg_part2 = df_part2[df_part2["NU_PTL"] == pg].copy()
            
            if not df_pg_part2.empty:
                df_pg_part2 = df_pg_part2.reset_index(drop=True)
                df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce").fillna(0)
                
                # Find date anchors to group related records
                date_mask = df_pg_part2["DT_APPL"].apply(is_valid_date)
                group_indices = df_pg_part2.index[date_mask].tolist()
                group_indices.append(len(df_pg_part2))  # sentinel for last group

                # Process each group between date anchors
                for i in range(len(group_indices) - 1):
                    start = group_indices[i]
                    end = group_indices[i + 1]
                    group = df_pg_part2.iloc[start:end]
                    
                    # Check if group has any secured collateral (not 0, -, NaN, etc.)
                    secured_col_types = [str(col) for col in group["COL_TYPE"] if str(col) not in ["0", "-", "NaN", "nan", "None", "none"]]
                    
                    if secured_col_types:
                        # Group contains secured financing
                        task7_count += 1
                        task7_outstanding += group["IM_AM"].sum()
                    else:
                        # Group contains only unsecured financing
                        task8_count += 1
                        task8_outstanding += group["IM_AM"].sum()

        task7_text = (f"7. Secured Financing: {task7_count} \n" f"(Outstanding: {task7_outstanding:,.2f})\n")
        task8_text = (f"8. Unsecured Financing: {task8_count} \n" f"(Outstanding: {task8_outstanding:,.2f})")
        

        # Display all calculated tasks
        all_tasks_text = f"{task1_text}\n{task2_text}\n{task3_text}\n{task4_text}\n{task5_text}\n{task6_text}\n{task7_text}\n{task8_text}"
        self.content_label.configure(text=all_tasks_text)

        # Show content area if not visible
        if not self.visible:
            self.content_frame.pack(fill="both", expand=True, padx=0, pady=(0, 0), side="top")
            self.content_label.pack(fill="both", expand=True, padx=16, pady=12, side="top")
            self.visible = True
            self.set_panel_height(self.current_height)

    def hide_content(self):
        """Hide the task content panel and animate to minimum height."""
        self.content_label.pack_forget()
        self.content_frame.pack_forget()
        self.visible = False
        self.animating = False
        self.animate_panel_height(self.current_height, self.min_height, duration=200)  # Animate close

    def set_panel_height(self, height):
        """
        Set the panel height within allowed bounds.
        
        Args:
            height (int): Desired height in pixels
        """
        height = max(self.min_height, min(self.max_height, int(height)))
        self.current_height = height
        self.frame.configure(height=height)
        self.frame.pack_propagate(False)

    def animate_panel_height(self, start, end, duration=200):
        """
        Smoothly animate panel height change.
        
        Args:
            start (int): Starting height
            end (int): Target height  
            duration (int): Animation duration in milliseconds
        """
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
        """
        Begin panel resize operation.
        
        Args:
            event: Mouse button press event
        """
        self.resizing = True
        self.start_y = event.y_root
        self.orig_height = self.current_height

    def perform_resize(self, event):
        """
        Handle panel resizing during mouse drag.
        
        Args:
            event: Mouse motion event
        """
        if self.resizing:
            delta = self.start_y - event.y_root
            new_height = self.orig_height + delta
            if int(new_height) != int(self.current_height):
                self.set_panel_height(new_height)

    def toggle_minimize(self, event):
        """
        Toggle between minimized and last expanded height.
        
        Args:
            event: Double-click event
        """
        if self.current_height > self.min_height + 10:
            self.last_height = self.current_height
            self.animate_panel_height(self.current_height, self.min_height, duration=200)
        else:
            self.animate_panel_height(self.current_height, self.last_height, duration=200)

class ExcelAllTask:
    """
    Excel export interface showing all customers and their calculated tasks.
    
    Displays a comprehensive table with all customers (NU_PTL) and their
    calculated task results in a searchable, exportable format. Uses optimized
    batch processing for performance with large datasets.
    
    Features:
    - Searchable table with column filtering
    - Export to Excel functionality  
    - Batch processing with progress indication
    - Navigation through search results
    - Identical business logic to TaskTabBar but optimized for bulk processing
    """
    def __init__(self, parent):
        """
        Initialize the Excel All Task interface.
        
        Args:
            parent: Parent widget to contain this interface
        """
        self.parent = parent
        self.search_var = tk.StringVar()
        self.frame = ctk.CTkFrame(parent, corner_radius=12)
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)
        self._repeat_job = None
        self._repeat_fast_job = None
        self._repeat_fast_timer = None

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
        
        # --- Table Section ---
        self.columns = [
            "NU_PTL",
            "pending applications last 1 month",
            "Credit Card utilization",
            "earliest date",
            "Unsecured Facilities Approved last 12 months",
            "Unsecured Facilities Approved last 18 months",
            "Date CCRIS pulled – Date earliest financing",
            "only 1 facility",
            "Secured financing",
            "Secured financing (Total outstanding)",
            "Unsecured financing",
            "Unsecured financing (Total outstanding)"
        ]
        table_frame = ctk.CTkFrame(self.frame, corner_radius=8)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)


        # Search bar with icon
        search_icon = ctk.CTkImage(Image.open("Picture/search.png"), size=(20, 20))
        left_arrow_icon = ctk.CTkImage(Image.open("Picture/left-arrow.png"), size=(24, 24))
        right_arrow_icon = ctk.CTkImage(Image.open("Picture/right-arrow.png"), size=(24, 24))

        search_entry_frame = ctk.CTkFrame(self.control_frame, corner_radius=8)
        search_entry_frame.pack(side="left", padx=(0, 10))
        
        # --- Column dropdown for search ---
        self.search_column_var = tk.StringVar(value="All")
        column_options = ["All"] + self.columns
        self.search_column_dropdown = ctk.CTkComboBox(
            search_entry_frame,
            values=column_options,
            variable=self.search_column_var,
            width=120,
            font=("Segoe UI", 13)
        )
        self.search_column_dropdown.pack(side="left", padx=(2, 2))

        # Previous button
        self.prev_btn = ctk.CTkButton(
            search_entry_frame,
            text="",
            image=left_arrow_icon,
            width=36,
            height=36,
            fg_color="transparent",
            hover_color="#eee",
        )
        self.prev_btn.pack(side="left", padx=(2, 2), pady=2)
        self.prev_btn.bind("<ButtonPress-1>", self._start_prev_repeat)
        self.prev_btn.bind("<ButtonRelease-1>", self._stop_prev_repeat)
        

        ctk.CTkLabel(search_entry_frame, image=search_icon, text="", fg_color="transparent").pack(side="left", padx=(2, 2))
        search_entry = ctk.CTkEntry(
            search_entry_frame, textvariable=self.search_var, width=180, fg_color="#222222", border_width=0, font=("Segoe UI", 14)
        )
        search_entry.pack(side="left", padx=(0, 2), pady=4)
        search_entry.bind("<Return>", self.on_search)
        search_entry.bind("<KeyRelease>", lambda e: None)
        
        # --- Add search counter label here ---
        self.search_counter_label = ctk.CTkLabel(
            search_entry_frame, text="", fg_color="transparent", font=("Segoe UI", 13)
        )
        self.search_counter_label.pack(side="left", padx=(4, 2))

        # Next button
        self.next_btn = ctk.CTkButton(
            search_entry_frame,
            text="",
            image=right_arrow_icon,
            width=36,
            height=36,
            fg_color="transparent",
            hover_color="#eee",
        )
        self.next_btn.pack(side="left", padx=(2, 2), pady=2)
        self.next_btn.bind("<ButtonPress-1>", self._start_next_repeat)
        self.next_btn.bind("<ButtonRelease-1>", self._stop_next_repeat)
        
        self.matching_row_ids = []
        self.match_index = 0
        self._repeat_job = None  # <-- Add this line

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
        size = (100, 100)
        for frame in ImageSequence.Iterator(self.loading_gif):
            rgba = frame.convert("RGBA").resize(size, Image.LANCZOS)
            self.loading_frames.append(ctk.CTkImage(rgba, size=size))
        
        # Create loading label and percentage label (placed over self.frame)
        self.loading_label = ctk.CTkLabel(self.frame, text="", fg_color="#141414")
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label.lower()
        self.loading_percentage_label = ctk.CTkLabel(self.frame, text="", fg_color="#141414", font=("Arial", 14))
        self.loading_percentage_label.place(relx=0.5, rely=0.65, anchor="center")
        self.loading_percentage_label.lower()
        
        self.loading_gif_running = False
        self.loading_progress = 0


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
        
        self.grid_populated = False
        
    def _start_next_repeat(self, event=None):
        self._stop_next_repeat()
        def slow_repeat():
            self.on_next_match()
            self._repeat_job = self.frame.after(400, slow_repeat)
        slow_repeat()
        # After 5 seconds, switch to fast repeat
        self._repeat_fast_timer = self.frame.after(3000, self._switch_to_fast_next_repeat)

    def _switch_to_fast_next_repeat(self):
        self._stop_next_repeat()
        def fast_repeat():
            self.on_next_match()
            self._repeat_fast_job = self.frame.after(80, fast_repeat)
        fast_repeat()

    def _stop_next_repeat(self, event=None):
        if self._repeat_job:
            self.frame.after_cancel(self._repeat_job)
            self._repeat_job = None
        if self._repeat_fast_job:
            self.frame.after_cancel(self._repeat_fast_job)
            self._repeat_fast_job = None
        if self._repeat_fast_timer:
            self.frame.after_cancel(self._repeat_fast_timer)
            self._repeat_fast_timer = None

    def _start_prev_repeat(self, event=None):
        self._stop_prev_repeat()
        def slow_repeat():
            self.on_prev_match()
            self._repeat_job = self.frame.after(400, slow_repeat)
        slow_repeat()
        self._repeat_fast_timer = self.frame.after(5000, self._switch_to_fast_prev_repeat)

    def _switch_to_fast_prev_repeat(self):
        self._stop_prev_repeat()
        def fast_repeat():
            self.on_prev_match()
            self._repeat_fast_job = self.frame.after(80, fast_repeat)
        fast_repeat()

    def _stop_prev_repeat(self, event=None):
        if self._repeat_job:
            self.frame.after_cancel(self._repeat_job)
            self._repeat_job = None
        if self._repeat_fast_job:
            self.frame.after_cancel(self._repeat_fast_job)
            self._repeat_fast_job = None
        if self._repeat_fast_timer:
            self.frame.after_cancel(self._repeat_fast_timer)
            self._repeat_fast_timer = None


    def export_data(self):
        """
        Export the current table data to an Excel file.
        
        Opens a file save dialog and exports all visible table rows
        to a new Excel workbook.
        """
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
        
    def update_search_counter(self):
        """Update the search result counter display."""
        if self.matching_row_ids:
            self.search_counter_label.configure(
                text=f"{self.match_index + 1} / {len(self.matching_row_ids)}"
            )
        else:
            self.search_counter_label.configure(text="0 / 0")

    def on_search(self, event=None):
        """
        Perform search across table data.
        
        Searches either all columns or a specific selected column
        for the entered query text (case-insensitive).
        
        Args:
            event: Optional key press event
        """
        query = self.search_var.get().lower()
        self.matching_row_ids = []
        self.match_index = 0
        selected_column = self.search_column_var.get()
        if not query:
            self.tree.selection_remove(self.tree.selection())
            self.update_search_counter()
            return
        for row_id in self.tree.get_children():
            row_data = self.tree.item(row_id)['values']
            if selected_column == "All":
                if any(query in str(val).lower() for val in row_data):
                    self.matching_row_ids.append(row_id)
            else:
                try:
                    col_idx = self.columns.index(selected_column)
                    if query in str(row_data[col_idx]).lower():
                        self.matching_row_ids.append(row_id)
                except Exception:
                    pass
        if self.matching_row_ids:
            self.highlight_match(0)
        else:
            self.tree.selection_remove(self.tree.selection())
            self.update_search_counter()

    def highlight_match(self, idx):
        """
        Highlight and scroll to a specific search match.
        
        Args:
            idx (int): Index of the match to highlight
        """
        # Remove previous selection
        self.tree.selection_remove(self.tree.selection())
        if not self.matching_row_ids:
            self.update_search_counter()
            return
        idx = idx % len(self.matching_row_ids)
        row_id = self.matching_row_ids[idx]
        self.tree.selection_set(row_id)
        self.tree.focus(row_id)
        self.tree.see(row_id)
        self.match_index = idx
        self.update_search_counter()

    def on_next_match(self):
        """Navigate to the next search match."""
        if self.matching_row_ids:
            next_idx = (self.match_index + 1) % len(self.matching_row_ids)
            self.highlight_match(next_idx)

    def on_prev_match(self):
        """Navigate to the previous search match."""
        if self.matching_row_ids:
            prev_idx = (self.match_index - 1) % len(self.matching_row_ids)
            self.highlight_match(prev_idx)

    def show_loading(self):
        """Display loading animation with progress tracking."""
        self.loading_label.lift()
        self.loading_percentage_label.lift()
        self.loading_gif_running = True
        self.loading_progress = 0  # reset progress
        self.animate_loading_gif(0)
        self.update_loading_progress()

    def hide_loading(self):
        """Hide loading animation and show completion."""
        self.loading_label.lower()
        self.loading_percentage_label.lower()
        self.loading_gif_running = False
        self.loading_percentage_label.configure(text="Loading: 100%")

    def animate_loading_gif(self, idx):
        """
        Animate the loading GIF frames.
        
        Args:
            idx (int): Current frame index
        """
        if not self.loading_gif_running:
            return
        frame = self.loading_frames[idx]
        self.loading_label.configure(image=frame, text="")
        next_idx = (idx + 1) % len(self.loading_frames)
        # Continue updating the gif every 60ms.
        self.frame.after(60, lambda: self.animate_loading_gif(next_idx))

    def update_loading_progress(self):
        """Update loading progress display (called periodically)."""
        if not self.loading_gif_running:
            return
        
        self.frame.after(100, self.update_loading_progress)


    def show(self):
        """Display the Excel All Task interface and populate data if needed."""
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)
        if not self.grid_populated:
            self.show_loading()
            threading.Thread(target=self.populate_grid, daemon=True).start()

    def hide(self):
        """Hide the Excel All Task interface."""
        self.frame.pack_forget()

    def populate_grid(self):
        """
        Populate the data grid with calculated tasks for all customers.
        
        Uses optimized batch processing to handle large datasets efficiently.
        Preserves original NU_PTL values and applies the same deduplication
        logic as CCRISReport.
        """
        self.tree.pack_forget()
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        raw_excel_data = ccris_report.excel_data
        if "part_1" not in raw_excel_data:
            self.hide_loading()
            return
        
        # Use deduplicated data
        if not hasattr(ccris_report, 'excel_data_deduplicated'):
            print("Applying deduplication to ExcelAllTask data...")
            ccris_report.excel_data_deduplicated = ccris_report.deduplicate_ccris_data(raw_excel_data)
        
        excel_data = ccris_report.excel_data_deduplicated
        
        # Extract NU_PTL values same way as CCRISReport
        deduplicated_part1 = excel_data["part_1"]
        raw_nu_ptl_values = deduplicated_part1["NU_PTL"].tolist()
        
        # Clean NU_PTL list while preserving original values
        pg_list = []
        for nu_ptl in raw_nu_ptl_values:
            nu_ptl_str = str(nu_ptl).strip()
            
            if nu_ptl_str in ["", "NaN", "nan", "None", "none"] or pd.isna(nu_ptl):
                continue
            
            if nu_ptl not in pg_list:
                pg_list.append(nu_ptl)
        
        # Sort numerically if possible
        try:
            pg_list.sort(key=lambda x: float(str(x)) if str(x).replace('.','').isdigit() else float('inf'))
        except:
            pg_list.sort(key=str)
        
        print(f"ExcelAllTask - Processing {len(pg_list)} deduplicated NU_PTL records")
        print(f"Sample NU_PTL values: {[str(x) for x in pg_list[:10]]}")
        
        # Pre-process data for optimal performance
        self.preprocess_data(excel_data)
        
        total_pgs = len(pg_list)
        batch_size = 50
        
        def process_batch(start_idx):
            end_idx = min(start_idx + batch_size, total_pgs)
            batch_rows = []
            
            for i in range(start_idx, end_idx):
                pg = str(pg_list[i])  # Convert to string for processing
                task_summaries = self.get_task_summaries_for_pg_optimized(pg)
                batch_rows.append([pg] + task_summaries)
            
            def insert_batch():
                for row in batch_rows:
                    self.tree.insert("", "end", values=row)
                
                progress = int((end_idx / total_pgs) * 100)
                self.loading_percentage_label.configure(text=f"Loading: {progress}%")
                
                if end_idx < total_pgs:
                    self.frame.after(10, lambda: threading.Thread(
                        target=lambda: process_batch(end_idx), daemon=True
                    ).start())
                else:
                    self.tree.pack(side="left", fill="both", expand=True)
                    self.hide_loading()
                    self.grid_populated = True
            
            self.frame.after(0, insert_batch)
        
        threading.Thread(target=lambda: process_batch(0), daemon=True).start()

    def preprocess_data(self, excel_data):
        """
        Pre-process data for optimal performance during bulk calculations.
        
        Converts data types once and groups by NU_PTL for faster lookup
        during task calculations.
        
        Args:
            excel_data (dict): Processed CCRIS data by part
        """
        self.processed_data = {}
        
        for part_name in ["part_2", "part_4"]:
            if part_name in excel_data and not excel_data[part_name].empty:
                df = excel_data[part_name].copy()
                
                # Convert date columns once
                if "DT_APPL" in df.columns:
                    df["DT_APPL"] = pd.to_datetime(df["DT_APPL"], errors="coerce")
                if "TM_AGG_UTE" in df.columns:
                    df["TM_AGG_UTE"] = pd.to_datetime(df["TM_AGG_UTE"], errors="coerce")
                
                # Convert numeric columns once
                numeric_cols = ["IM_AM", "IM_LIM_AM"]
                for col in numeric_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                
                # Only group if we have NU_PTL column and data
                if "NU_PTL" in df.columns and not df.empty:
                    try:
                        # Group by NU_PTL for faster lookup
                        grouped = df.groupby("NU_PTL")
                        self.processed_data[part_name] = grouped
                        print(f"Preprocessed {part_name}: {len(grouped.groups)} unique NU_PTL groups")
                    except Exception as e:
                        print(f"Error grouping {part_name}: {e}")
                        # Fallback: don't group, use original data
                        self.processed_data[part_name] = df
                else:
                    print(f"Skipping {part_name}: no NU_PTL column or empty data")
            else:
                print(f"Skipping {part_name}: not in excel_data or empty")

    def get_task_summaries_for_pg_optimized(self, pg):
        """
        Calculate optimized task summaries for a specific customer.
        
        Uses pre-processed grouped data for faster calculations.
        Implements identical business logic to TaskTabBar but optimized
        for bulk processing.
        
        Args:
            pg (str): Customer identifier (NU_PTL)
            
        Returns:
            list: Calculated values for all 11 task columns
        """
        try:
            # Get customer data from pre-processed groups
            df_pg_part2 = pd.DataFrame()
            df_pg_part4 = pd.DataFrame()
            
            # Check if part_2 exists and has groups
            if "part_2" in self.processed_data and hasattr(self.processed_data["part_2"], 'groups'):
                if pg in self.processed_data["part_2"].groups:
                    df_pg_part2 = self.processed_data["part_2"].get_group(pg).copy()
            
            # Check if part_4 exists and has groups
            if "part_4" in self.processed_data and hasattr(self.processed_data["part_4"], 'groups'):
                if pg in self.processed_data["part_4"].groups:
                    df_pg_part4 = self.processed_data["part_4"].get_group(pg).copy()
                    
        except (KeyError, AttributeError) as e:
            # No data for this NU_PTL or grouping issue
            print(f"Error getting data for NU_PTL {pg}: {e}")
            return ["-"] * 11
        
        # Calculate latest report date once for performance
        latest_report_date = None
        if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
            latest_report_date = df_pg_part4["TM_AGG_UTE"].max()
        elif not df_pg_part2.empty and "DT_APPL" in df_pg_part2.columns:
            latest_report_date = df_pg_part2["DT_APPL"].max()
        
        # Calculate all tasks using optimized methods
        task1 = self.calculate_task1_optimized(df_pg_part4, latest_report_date)
        task2 = self.calculate_task2_optimized(df_pg_part2)
        task3 = self.calculate_task3_optimized(df_pg_part2)
        task4, task5 = self.calculate_task4_5_optimized(df_pg_part2, latest_report_date)
        task6a, task6b = self.calculate_task6_optimized(df_pg_part2, latest_report_date)
        task7a, task7b, task8a, task8b = self.calculate_task7_8_optimized(df_pg_part2)
        
        return [
            task1, task2, task3, task4, task5,
            task6a, task6b,
            task7a, task7b, task8a, task8b
        ]

    def calculate_task1_optimized(self, df_pg_part4, latest_report_date):
        """
        Task 1: Count pending applications in the last month.
        
        Args:
            df_pg_part4: Part 4 data for customer
            latest_report_date: Latest report date
            
        Returns:
            str: Count of pending applications or "-"
        """
        if df_pg_part4.empty or latest_report_date is None:
            return "-"
        
        one_month_ago = latest_report_date - pd.DateOffset(months=1)
        mask = (
            (df_pg_part4["APPL_STS"] == "P") &
            (df_pg_part4["DT_APPL"] >= one_month_ago) &
            (df_pg_part4["DT_APPL"] <= latest_report_date)
        )
        return str(mask.sum())

    def calculate_task2_optimized(self, df_pg_part2):
        """
        Task 2: Calculate credit card utilization ratio.
        
        Finds CRDTCARD outstanding and divides by associated credit limits.
        Uses nearest previous approval date logic to match limits.
        
        Args:
            df_pg_part2: Part 2 data for customer
            
        Returns:
            str: Percentage utilization or "No outstanding found" or "-"
        """
        if df_pg_part2.empty:
            return "-"
        
        # Pre-filter CRDTCARD rows
        crdtcard_mask = df_pg_part2["FCY_TYPE"].astype(str).str.strip().str.upper() == "CRDTCARD"
        crdtcard_rows = df_pg_part2[crdtcard_mask]
        
        if crdtcard_rows.empty:
            return "-"
        
        crdtcard_outstanding = crdtcard_rows["IM_AM"].sum()
        
        # Optimized limit calculation
        used_dates = set()
        total_limit = 0
        
        for idx in crdtcard_rows.index:
            # Find nearest previous row with valid date
            prev_rows = df_pg_part2.loc[:idx]
            valid_dates = prev_rows[prev_rows["DT_APPL"].notna()]["DT_APPL"]
            
            if not valid_dates.empty:
                latest_date = valid_dates.iloc[-1]
                date_str = str(latest_date)
                
                if date_str not in used_dates:
                    # Get the limit from the row with this date
                    date_row = prev_rows[prev_rows["DT_APPL"] == latest_date].iloc[-1]
                    total_limit += date_row["IM_LIM_AM"]
                    used_dates.add(date_str)
        
        if total_limit > 0:
            ratio = crdtcard_outstanding / total_limit
            return f"{ratio:.2%}"
        return "No outstanding found"

    def calculate_task3_optimized(self, df_pg_part2):
        """
        Task 3: Find earliest financing date.
        
        Args:
            df_pg_part2: Part 2 data for customer
            
        Returns:
            str: Earliest date in dd-mm-yyyy format or "-"
        """
        if df_pg_part2.empty:
            return "-"
        
        earliest_date = df_pg_part2["DT_APPL"].min()
        if pd.notna(earliest_date):
            return earliest_date.strftime('%d-%m-%Y')
        return "-"

    def calculate_task4_5_optimized(self, df_pg_part2, _unused=None):
        """
        Task 4 & 5: Count unsecured facilities in 12 and 18 months.
        Uses TM_AGG_UTE from part_2 as report date if available.
        """
        if df_pg_part2.empty:
            return "-", "-"

        df_pg_part2 = df_pg_part2.copy()
        df_pg_part2["DT_APPL"] = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce")
        df_pg_part2["TM_AGG_UTE"] = pd.to_datetime(df_pg_part2["TM_AGG_UTE"], errors="coerce")

        # Get report date from part_2 TM_AGG_UTE
        report_date = df_pg_part2["TM_AGG_UTE"].max()
        if pd.isna(report_date):
            report_date = df_pg_part2["DT_APPL"].max()

        if pd.isna(report_date):
            return "-", "-"

        one_year_ago = report_date - pd.DateOffset(months=12)
        eighteen_months_ago = report_date - pd.DateOffset(months=18)

        # Helper function to count unsecured facilities by group
        def count_unsecured_in_period(start_date):
            unsecured_count = 0
            
            # Group by REC_CTR to handle cases where data spans multiple rows
            if 'REC_CTR' in df_pg_part2.columns:
                for rec_ctr, group in df_pg_part2.groupby('REC_CTR'):
                    has_unsecured = any(str(col).strip() in ["0", "-", "NaN", "nan"] for col in group['COL_TYPE'])
                    has_valid_date = any(pd.notna(dt) and dt >= start_date and dt <= report_date for dt in group['DT_APPL'])
                    
                    if has_unsecured and has_valid_date:
                        unsecured_count += 1
            else:
                # Fallback: use original row-based logic but include NaN
                col_type_unsecured = df_pg_part2["COL_TYPE"].astype(str).str.strip().isin(["0", "-", "NaN", "nan"])
                date_in_range = (df_pg_part2["DT_APPL"] >= start_date) & (df_pg_part2["DT_APPL"] <= report_date)
                mask = col_type_unsecured & date_in_range
                unsecured_count = mask.sum()
            
            return unsecured_count

        # Task 4: 12 months
        task4_result = count_unsecured_in_period(one_year_ago)

        # Task 5: 18 months
        task5_result = count_unsecured_in_period(eighteen_months_ago)

        return str(task4_result), str(task5_result)

    def calculate_task6_optimized(self, df_pg_part2, latest_report_date):
        """
        Task 6: Thin CCRIS analysis (months between dates + facility diversity).
        
        Args:
            df_pg_part2: Part 2 data for customer
            latest_report_date: Latest report date
            
        Returns:
            tuple: (months_difference, only_one_facility) as strings
        """
        if df_pg_part2.empty:
            return "-", "-"

        # Calculate months difference
        earliest_date = df_pg_part2["DT_APPL"].min()
        report_date = latest_report_date if latest_report_date else df_pg_part2["DT_APPL"].max()

        months_diff = "-"
        if pd.notna(earliest_date) and pd.notna(report_date):
            months_diff = (report_date.year - earliest_date.year) * 12 + (report_date.month - earliest_date.month)

        # Only count non-empty, non-NaN facilities
        facilities = df_pg_part2["FCY_TYPE"].dropna().astype(str).str.strip()
        facilities = facilities[(facilities != "") & (facilities != "NaN") & (facilities != "-")]
        only_one_facility = "Yes" if facilities.nunique() == 1 else "No"

        return str(months_diff), only_one_facility

    def calculate_task7_8_optimized(self, df_pg_part2):
        """Fixed Task 7 & 8 calculation - treat NaN as unsecured"""
        if df_pg_part2.empty:
            return "-", "-", "-", "-"
        
        df_pg_part2 = df_pg_part2.reset_index(drop=True)
        df_pg_part2["IM_AM"] = pd.to_numeric(df_pg_part2["IM_AM"], errors="coerce").fillna(0)
        
        # Find date anchors
        def is_valid_date(date_val):
            if pd.isna(date_val):
                return False
            date_str = str(date_val).strip()
            if date_str in ["-", "", "NaN", "nan", "None", "none"]:
                return False
            if any(char.isdigit() for char in date_str) and ("-" in date_str or "/" in date_str):
                return True
            return False
        
        date_mask = df_pg_part2["DT_APPL"].apply(is_valid_date)
        date_indices = df_pg_part2.index[date_mask].tolist()
        date_indices.append(len(df_pg_part2))  # Add sentinel
        
        task7_count = 0
        task7_outstanding = 0.0
        task8_count = 0
        task8_outstanding = 0.0
        
        # Process groups
        for i in range(len(date_indices) - 1):
            start_idx = date_indices[i]
            end_idx = date_indices[i + 1]
            group = df_pg_part2.iloc[start_idx:end_idx]
            
            # **FIXED: Treat "NaN", "-", "0" as unsecured**
            secured_items = group[~group["COL_TYPE"].astype(str).isin(["0", "-", "NaN", "nan", "None", "none"])]
            
            if not secured_items.empty:
                # At least one secured item in group
                task7_count += 1
                group_outstanding = group["IM_AM"].sum()
                task7_outstanding += group_outstanding
            else:
                # All items are unsecured
                task8_count += 1
                group_outstanding = group["IM_AM"].sum()
                task8_outstanding += group_outstanding
        
        # Format results
        task7a_str = str(task7_count) if task7_count > 0 else "-"
        task7b_str = f"{task7_outstanding:,.2f}" if task7_count > 0 else "-"
        task8a_str = str(task8_count) if task8_count > 0 else "-"
        task8b_str = f"{task8_outstanding:,.2f}" if task8_count > 0 else "-"
        
        return task7a_str, task7b_str, task8a_str, task8b_str


    def get_latest_report_date(self, df_pg_part4, df_pg_part2):
        """
        Get the latest report date from either part 4 or part 2 data.
        
        Args:
            df_pg_part4 (DataFrame): Customer data from part 4
            df_pg_part2 (DataFrame): Customer data from part 2
            
        Returns:
            datetime or None: Latest report date found, or None if not available
        """
        if not df_pg_part4.empty and "TM_AGG_UTE" in df_pg_part4.columns:
            date = pd.to_datetime(df_pg_part4["TM_AGG_UTE"], format="%d/%m/%Y", errors="coerce").max()
            if pd.notna(date):
                return date
        if not df_pg_part2.empty and "DT_APPL" in df_pg_part2.columns:
            date = pd.to_datetime(df_pg_part2["DT_APPL"], errors="coerce").max()
            if pd.notna(date):
                return date
        return None
 
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
