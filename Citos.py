import threading
import time
import subprocess
import sys
import os
import decimal
import math
import xlsxwriter
import datetime
import customtkinter
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import xml.dom.minidom
import customtkinter as ctk 
from tkinter import filedialog
from PIL import Image, ImageTk
from customtkinter import CTkTabview
from collections import defaultdict
from lxml import etree, html
from xml.sax.saxutils import escape
from xml.dom import minidom
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import psutil
import win32gui
import win32con

def extract_first_report(xml_str):
    """
    Returns the XML string up to and including the first </report> tag.
    If </report> is not found, returns the original string.
    """
    end_idx = xml_str.lower().find("</report>")
    if end_idx != -1:
        return xml_str[:end_idx + len("</report>")]
    return xml_str

def count_section_presence(dom, section_id):
    """
    Returns 1 if <section id="section_id"> exists, 0 otherwise.
    """
    for section in dom.getElementsByTagName("section"):
        if section.hasAttribute("id") and section.getAttribute("id").strip().upper() == section_id:
            return 1
    return 0

def count_records_in_section(dom, section_id):
    """
    Returns the number of <record> in <section id="section_id">.
    """
    count = 0
    for section in dom.getElementsByTagName("section"):
        if section.hasAttribute("id") and section.getAttribute("id").strip().upper() == section_id:
            count += len(section.getElementsByTagName("record"))
    return count

def is_integrate_running():
        """
        Check if a process running integrate.py exists.
        Returns the process if found, otherwise None.
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

def get_node_text(node):
    """Return concatenated text of all text node children; return '-' if empty."""
    parts = []
    for child in node.childNodes:
        if child.nodeType == child.TEXT_NODE:
            parts.append(child.nodeValue)
        elif child.nodeType == child.ELEMENT_NODE:
            parts.append(get_node_text(child))
    text = "".join(parts).strip()
    return text if text else "-"

def clean_malformed_xml(xml_str):
   
    import re

    if not isinstance(xml_str, str):
        return "<root></root>"

    # First, extract only the content up to the first </report> tag
    xml_str = extract_first_report(xml_str)

    # Remove invalid control characters
    xml_str = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", xml_str).strip()

    try:
        # Try parsing normally first
        etree.fromstring(xml_str.encode('utf-8'))
        return xml_str  # It's already valid XML
    except Exception:
        try:
            # Use HTML parser to auto-fix
            repaired_tree = html.fromstring(xml_str)
            repaired_xml = etree.tostring(repaired_tree, pretty_print=True, encoding="unicode")
            return f"<root>{repaired_xml}</root>"
        except Exception:
            # Fallback
            return f"<root>{escape(xml_str)}</root>"
        
class CTOSReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CTOS Report Generator")
        self.geometry("1800x900")
        
        self.current_theme = "dark" 
        
        customtkinter.set_default_color_theme("Themes/patina.json")
        
        # Shared data (Excel, parsed XML, etc.)
        self.shared_data = None
        
        # Header click counter
        self.header_click_count = 0

        # Sidebar settings
        self.SIDEBAR_EXPANDED_WIDTH = 200
        self.SIDEBAR_SHRUNK_WIDTH = 60
        self.sidebar_expanded = True

        # Sidebar container with transparent background
        self.sidebar_container = ctk.CTkFrame(self, fg_color="transparent")
        self.sidebar_container.pack(side="left", fill="y")

        # Sidebar frame with a subtle background color and rounded corners
        self.sidebar = ctk.CTkFrame(self.sidebar_container, width=self.SIDEBAR_EXPANDED_WIDTH, corner_radius=8)
        self.sidebar.pack(side="left", fill="y", padx=5, pady=5)
        self.sidebar.pack_propagate(False)

        # Hamburger Button (always visible)
        hamburger_img = ctk.CTkImage(Image.open("Picture/hamburger.png"), size=(24, 24))
        self.hamburger_btn = ctk.CTkButton(
            self.sidebar,
            text="",
            image=hamburger_img,
            width=40,
            height=40,
            fg_color="transparent",
            hover_color="#555",
            command=self.toggle_sidebar
        )
        self.hamburger_btn.pack(pady=(10, 0), padx=10, anchor="nw")
        
        self.importing_icon = ctk.CTkImage(Image.open("Picture/importing.png"), size=(32, 32))
        self.xml_icon = ctk.CTkImage(Image.open("Picture/xml.png"), size=(32, 32))  # Try bigger size
        self.tab_ccris_icon = ctk.CTkImage(Image.open("Picture/tab_ccris.png"), size=(32, 32))
        self.summary_icon = ctk.CTkImage(Image.open("Picture/summary.png"), size=(32, 32))
        self.back_to_main_icon = ctk.CTkImage(Image.open("Picture/back_to_main.png"), size=(32, 32))

        # Sidebar buttons with improved spacing & bigger size
        button_font = ctk.CTkFont(family="Segoe UI", size=16, weight="bold")  # increased font size
        self.import_button = ctk.CTkButton(
            self.sidebar,
            text="Import Excel",
            command=self.import_excel,
            font=button_font,
            width=150,
            height=50,
            image=None,
        )
        self.import_button.pack(pady=10, padx=0)
        
        self.xml_format_button = ctk.CTkButton(
            self.sidebar,
            text="XML Format",
            command=self.show_xml_format,
            font=button_font,
            width=150,
            height=50,
            image=None,
        )
        self.xml_format_button.pack(pady=10, padx=0)
        
        self.ctos_report_button = ctk.CTkButton(
            self.sidebar,
            text="CTOS Report",
            command=self.show_ctos_report,
            font=button_font,
            width=150,
            height=50,
            image=None,
        )
        self.ctos_report_button.pack(pady=10, padx=0)

        self.ctos_summary_button = ctk.CTkButton(
            self.sidebar,
            text="CTOS Summary",
            command=self.show_ctos_summary,
            font=button_font,
            width=150,
            height=50,
            image=None,
        )
        self.ctos_summary_button.pack(pady=10, padx=0)

        self.main_app_button = ctk.CTkButton(
            self.sidebar,
            text="Back to Main",
            command=self.show_main_app,
            font=button_font,
            width=150,
            height=50,
            image=None,
        )
        self.main_app_button.pack(pady=10, padx=0)

        self.sidebar_buttons = [
            self.import_button,
            self.xml_format_button,
            self.ctos_report_button,
            self.ctos_summary_button,
            self.main_app_button
        ]
        
        
        try:
            self.dark_icon = ctk.CTkImage(Image.open("Picture/dark_mode_icon.png"), size=(24, 24))
            self.system_icon = ctk.CTkImage(Image.open("Picture/light_mode_icon.png"), size=(24, 24))
        except Exception as e:
            print(f"Error loading icons: {e}")
            self.dark_icon = None
            self.system_icon = None

        # Default: if current mode is System, we show dark_icon (clicking will set dark)
        self.mode_toggle_btn = ctk.CTkButton(
            self.sidebar,
            text="",
            image=self.dark_icon,
            width=40,
            height=40,
            fg_color="transparent",
            hover_color="#444",
            command=self.toggle_mode
        )
        self.mode_toggle_btn.pack(side="bottom", pady=15)
     
        # Add a small sidebar spacer
        self.sidebar_spacer = ctk.CTkFrame(self, width=5, fg_color="transparent")
        self.sidebar_spacer.pack(side="left", fill="y")

        # Main Content Area with slight padding and no corner radius
        self.main_frame = ctk.CTkFrame(self, corner_radius=8)
        self.main_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        # Initialize Views (adjusted margins for better spacing)
        self.xml_format_view = XMLFormatView(self.main_frame, self)
        self.ctos_report_view = CTOSReportView(self.main_frame, self)
        self.ctos_summary_view = CTOSSummaryView(self.main_frame, self)
        # Show default view here; use pack_forget on hidden ones
        self.show_xml_format()
        
    
    # In toggle_sidebar:
    def toggle_sidebar(self):
        if self.sidebar_expanded:
            self.sidebar.configure(width=self.SIDEBAR_SHRUNK_WIDTH)
            # Show only icons, hide text, make buttons transparent and fit icon
            for btn, icon in [
                (self.import_button, self.importing_icon),
                (self.xml_format_button, self.xml_icon),
                (self.ctos_report_button, self.tab_ccris_icon),
                (self.ctos_summary_button, self.summary_icon),
                (self.main_app_button, self.back_to_main_icon)
            ]:
                btn.configure(
                    text="",
                    image=icon,
                    width=48,
                    height=48,
                    anchor="center",
                    font=("Arial", 1)
                )
            self.sidebar_expanded = False
        else:
            self.sidebar.configure(width=self.SIDEBAR_EXPANDED_WIDTH)
            # Show only text, hide icons, restore button width/height
            self.import_button.configure(
                text="Import Excel",
                image=None,
                width=150,
                height=50,
                anchor="w",
                font=("Segoe UI", 16, "bold")
            )
            self.xml_format_button.configure(
                text="XML Format",
                image=None,
                width=150,
                height=50,
                anchor="w",
                font=("Segoe UI", 16, "bold")
            )
            self.ctos_report_button.configure(
                text="CTOS Report",
                image=None,
                width=150,
                height=50,
                anchor="w",
                font=("Segoe UI", 16, "bold")
            )
            self.ctos_summary_button.configure(
                text="CTOS Summary",
                image=None,
                width=150,
                height=50,
                anchor="w",
                font=("Segoe UI", 16, "bold")
            )
            self.main_app_button.configure(
                text="Back to Main",
                image=None,
                width=150,
                height=50,
                anchor="w",
                font=("Segoe UI", 16, "bold")
            )
            self.sidebar_expanded = True
                     
    def toggle_mode(self):
        customtkinter.set_default_color_theme("Themes/patina.json")
        if self.current_theme == "dark":
            self.current_theme = "light"
            ctk.set_appearance_mode("light")
            if self.system_icon is not None:
                self.mode_toggle_btn.configure(image=self.system_icon)
            else:
                self.mode_toggle_btn.configure(text="Light")
        else:
            self.current_theme = "dark"
            ctk.set_appearance_mode("dark")
            if self.dark_icon is not None:
                self.mode_toggle_btn.configure(image=self.dark_icon)
            else:
                self.mode_toggle_btn.configure(text="Dark")
        self.update_idletasks()
        self.update_treeview_style()
        
    def update_treeview_style(self):
        style = ttk.Style()
        # Set default values for Treeview styling (ignoring the theme from patina.json)
        style.configure("Treeview",
                        rowheight=28,
                        font=("Segoe UI", 11),
                        background="#FFFFFF",      # White background for data
                        foreground="#000000",      # Black text for data
                        fieldbackground="#FFFFFF") # White background in cells
        style.configure("Treeview.Heading",
                        font=("Segoe UI", 11, "bold"),
                        background="#F0F0F0",      # Light gray header background
                        foreground="#000000")      # Black header text
        style.map("Treeview",
                background=[("selected", "#A67C5F")],
                foreground=[("selected", "#000000")])
                
    
    def show_progress_popup(self, title="Processing...", message="Please wait..."):
        self.progress_popup = ctk.CTkToplevel(self)
        self.progress_popup.title(title)
        self.progress_popup.geometry("300x100")
        self.progress_popup.resizable(False, False)
        self.progress_popup.grab_set()
        self.progress_popup.attributes("-topmost", True)
        ctk.CTkLabel(self.progress_popup, text=message).pack(pady=10)
        self.progress_bar = ctk.CTkProgressBar(self.progress_popup, mode="indeterminate")
        self.progress_bar.pack(padx=20, pady=10, fill="x")
        self.progress_bar.start()

    def destroy_progress_popup(self):
        if hasattr(self, "progress_popup") and self.progress_popup.winfo_exists():
            self.progress_bar.stop()
            self.progress_popup.destroy()

    def import_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        self.show_progress_popup(title="Importing Excel", message="Cleaning XML records...")
        def import_thread():
            try:
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip().str.upper()
                if "NU_PTL" not in df.columns or "XML" not in df.columns:
                    self.after(0, lambda: messagebox.showerror("Error", "Excel must contain columns: NU_PTL and XML"))
                    self.after(0, self.destroy_progress_popup)
                    return
                
                # Add ROW_ID column if it doesn't exist
                if "ROW_ID" not in df.columns:
                    df["ROW_ID"] = 0
                
                # Remove duplicates: for each NU_PTL and ROW_ID combination, keep only the last occurrence
                original_count = len(df)
                df = df.drop_duplicates(subset=["NU_PTL", "ROW_ID"], keep="last").reset_index(drop=True)
                
                # Sort by NU_PTL and ROW_ID to ensure proper ordering
                df = df.sort_values(["NU_PTL", "ROW_ID"]).reset_index(drop=True)
                
                # Create cleaned XML dict for XMLFormatView (combine XMLs per NU_PTL)
                cleaned_xml_dict = {}
                for nu_ptl, group in df.groupby("NU_PTL"):
                    # Sort by ROW_ID and combine all XML fragments for this NU_PTL
                    group_sorted = group.sort_values("ROW_ID")
                    combined_xml = ""
                    for _, row in group_sorted.iterrows():
                        raw_xml = str(row["XML"]) if pd.notna(row["XML"]) else ""
                        if raw_xml.strip():
                            combined_xml += clean_malformed_xml(raw_xml)
                    cleaned_xml_dict[str(nu_ptl)] = combined_xml
                    
                def update_data():
                    self.shared_data = df
                    self.xml_format_view.xml_data = cleaned_xml_dict
                    self.destroy_progress_popup()
                    removed_count = original_count - len(df)
                    messagebox.showinfo("Success", f"Excel imported successfully!\nProcessed {len(df)} unique records.\nRemoved {removed_count} duplicate entries.")
                self.after(0, update_data)
            except Exception as e:
                self.after(0, self.destroy_progress_popup)
                self.after(0, lambda: messagebox.showerror("Error", f"Error importing Excel file: {e}"))
        threading.Thread(target=import_thread, daemon=True).start()

    def show_ctos_report(self):
        self.ctos_report_view.pack(fill="both", expand=True)
        self.xml_format_view.pack_forget()
        self.ctos_summary_view.pack_forget()

    def show_xml_format(self):
        self.xml_format_view.pack(fill="both", expand=True)
        self.ctos_report_view.pack_forget()
        self.ctos_summary_view.pack_forget()
        
    def show_ctos_summary(self):
        self.ctos_summary_view.pack(fill="both", expand=True)
        self.ctos_report_view.pack_forget()
        self.xml_format_view.pack_forget()


    def show_main_app(self):
        proc = is_integrate_running()
        exe_path = os.path.join(os.getcwd(), "integrate.exe")
        if proc is None:
            subprocess.Popen([exe_path])
        else:
            bring_integrate_to_front()
        
        
class CTOSReportView(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app  # Reference to the main app to access shared data
        self.account_var = tk.StringVar()
        self.search_var = tk.StringVar()
        self.all_accounts = []
        self.current_index = 0
        self.filtered_data = None

        # --- Header Frame ---
        header_frame = ctk.CTkFrame(self)
        header_frame.pack(fill="x", pady=10)
        

        # CTOS logo in center
        try:
            ctos_img = Image.open("Picture/ctos.png")
            self.ctos_logo = ctk.CTkImage(light_image=ctos_img, size=(220, 50))
            ctos_logo_label = ctk.CTkLabel(header_frame, image=self.ctos_logo, text="")
            ctos_logo_label.pack(side="top", pady=5)
        except Exception:
            ctos_logo_label = ctk.CTkLabel(header_frame, text="CTOS")
            ctos_logo_label.pack(side="top", pady=5)


        # Al Rajhi logo on right
        try:
            alrajhi_img = Image.open("Picture/alrajhi_logo.png")
            self.alrajhi_logo = ctk.CTkImage(light_image=alrajhi_img, size=(220, 50))
            alrajhi_logo_label = ctk.CTkLabel(header_frame, image=self.alrajhi_logo, text="")
            alrajhi_logo_label.place(relx=1.0, rely=0.0, anchor="ne")
        except Exception:
            alrajhi_logo_label = ctk.CTkLabel(header_frame, text="Al Rajhi")
            alrajhi_logo_label.place(relx=1.0, rely=0.0, anchor="ne")

        # --- Control Frame ---
        self.control_frame = ctk.CTkFrame(self)
        self.control_frame.pack(fill="x", pady=5)
        self.control_frame.grid_columnconfigure(0, weight=1)
        self.control_frame.grid_columnconfigure(1, weight=0)
        self.control_frame.grid_columnconfigure(2, weight=1)

        left_arrow_icon = ctk.CTkImage(Image.open("Picture/left-arrow.png"), size=(24, 24))
        right_arrow_icon = ctk.CTkImage(Image.open("Picture/right-arrow.png"), size=(24, 24))

        self.prev_btn = ctk.CTkButton(
            self.control_frame,
            text="",
            image=left_arrow_icon,
            fg_color="transparent",
            hover_color="#444",
            command=self.go_to_previous
        )
        self.prev_btn.grid(row=0, column=1, padx=10, pady=5, sticky="e")

        self.ttk_style = ttk.Style()
        self.ttk_style.theme_use('clam')
        self.account_combobox = ttk.Combobox(
            self.control_frame, textvariable=self.account_var, values=[], width=25
        )
        self.account_combobox.grid(row=0, column=2, padx=10, pady=5)
        self.account_combobox.bind("<<ComboboxSelected>>", self.display_data)
        self.account_combobox.bind("<KeyRelease>", self.on_account_typing)

        self.next_btn = ctk.CTkButton(
            self.control_frame,
            text="",
            image=right_arrow_icon,
            fg_color="transparent",
            hover_color="#444",
            command=self.go_to_next
        )
        self.next_btn.grid(row=0, column=3, padx=10, pady=5, sticky="w")

        self.export_icon = ctk.CTkImage(Image.open("Picture/export.png"), size=(24, 24))
        self.convert_button = ctk.CTkButton(self.control_frame, text="Old Ctos", image=self.export_icon, command=self.convert_to_excel)
        self.convert_button.grid(row=0, column=0, padx=5)
        self.convert_new_button = ctk.CTkButton(
            self.control_frame,
            text="New CTOS",
            image=self.export_icon,
            command=self.convert_new_ctos_to_excel
        )
        self.convert_new_button.grid(row=0, column=5, padx=5)

        # --- Treeview for displaying parsed XML data ---
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.tree = ttk.Treeview(self.tree_frame, show="headings")
        self.tree.pack(fill="both", expand=True, side="left", padx=5, pady=5)
        self.tree["columns"] = ["Field", "Value"]
        self.tree.heading("Field", text="Field")
        self.tree.heading("Value", text="Value")
        self.tree.column("Field", anchor="center", width=300)
        self.tree.column("Value", anchor="center", width=400)

        self.tree.bind("<Button-3>", self.show_context_menu)
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Copy Row", command=self.copy_row)
        self.context_menu.add_command(label="Copy Cell", command=self.copy_cell)
        self._right_click_row = None
        self._right_click_col = None

        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        refresh_button = ctk.CTkButton(self, text="Refresh", command=self.refresh_data)
        refresh_button.pack(pady=10)


    def on_account_typing(self, event):
        typed = self.account_var.get().lower()
        filtered = [acct for acct in self.all_accounts if typed in acct.lower()]
        self.account_combobox['values'] = filtered

    def refresh_data(self):
        xml_format_view = self.app.xml_format_view
        if not xml_format_view.xml_data:
            return

        # --- Improved logic for picking best XML per account ---
        def is_new_ctos_xml(xml_str):
            return any(tag in xml_str for tag in ["<section_d2", "<section_d4", "<section_etr_plus", "<section_etl"])

        def tag_count(xml):
            import re
            return len(re.findall(r"<[a-zA-Z0-9_]+", xml))

        # Group by NU_PTL, collect all XMLs for each
        from collections import defaultdict
        nuptl_to_xmls = defaultdict(list)
        for nuptl, xml in xml_format_view.xml_data.items():
            if xml and isinstance(xml, str):
                nuptl_to_xmls[nuptl].append(xml)

        cleaned_data = {}
        for nuptl, xml_list in nuptl_to_xmls.items():
            # Remove trailing garbage after </report> using standardized function
            cleaned_xmls = [extract_first_report(xml) for xml in xml_list]
            # Prefer new CTOS if available
            new_ctos_xmls = [xml for xml in cleaned_xmls if is_new_ctos_xml(xml)]
            if new_ctos_xmls:
                best_xml = max(new_ctos_xmls, key=tag_count)
            else:
                best_xml = max(cleaned_xmls, key=tag_count)
            cleaned_data[nuptl] = best_xml

        self.filtered_data = pd.DataFrame.from_dict(cleaned_data, orient="index", columns=["XML"])
        self.filtered_data.reset_index(inplace=True)
        self.filtered_data.rename(columns={"index": "NU_PTL"}, inplace=True)
        self.all_accounts = self.filtered_data["NU_PTL"].tolist()
        self.account_combobox['values'] = self.all_accounts
        if self.all_accounts:
            self.account_combobox.current(self.current_index)
        self.display_data()

    def show_context_menu(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            row_id = self.tree.identify_row(event.y)
            col_id = self.tree.identify_column(event.x)
            self._right_click_row = row_id
            self._right_click_col = col_id
            self.context_menu.tk_popup(event.x_root, event.y_root)
        else:
            self._right_click_row = None
            self._right_click_col = None

    def copy_row(self):
        if self._right_click_row:
            values = self.tree.item(self._right_click_row, "values")
            text = "\t".join(str(v) for v in values)
            self.clipboard_clear()
            self.clipboard_append(text)

    def copy_cell(self):
        if self._right_click_row and self._right_click_col:
            col_index = int(self._right_click_col.replace("#", "")) - 1
            values = self.tree.item(self._right_click_row, "values")
            if 0 <= col_index < len(values):
                text = str(values[col_index])
                self.clipboard_clear()
                self.clipboard_append(text)

    def display_data(self, event=None):
        selected_account = self.account_var.get()
        if selected_account in self.all_accounts:
            self.current_index = self.all_accounts.index(selected_account)
        else:
            self.current_index = 0

        current_row = self.filtered_data.iloc[self.current_index]
        nu_ptl = current_row.get("NU_PTL", "")
        self.search_var.set(str(nu_ptl))
        xml_data = current_row.get("XML", "")
        if pd.isna(xml_data) or not xml_data.strip():
            xml_data = "<No XML Data>"

        # --- Add this line here ---
        xml_data = extract_first_report(xml_data)

        self.tree.delete(*self.tree.get_children())
        try:
            dom = xml.dom.minidom.parseString(xml_data)
            root = dom.documentElement
            self.parse_xml_to_treeview(root, "")
        except Exception as e:
            self.tree.insert("", "end", values=["Error", str(e)])


    def parse_xml_to_treeview(self, node, parent_path=""):
        
        # If the node is a known wrapper (broken XML outer tags), skip it and process its children
        if node.nodeType == xml.dom.minidom.Node.ELEMENT_NODE and node.tagName.lower() in ["root", "span"]:
            for child in node.childNodes:
                self.parse_xml_to_treeview(child, parent_path)
            return

        for child in node.childNodes:
            if child.nodeType != xml.dom.minidom.Node.ELEMENT_NODE:
                continue
            tag = child.tagName

            # --- Existing logic ---
            if tag == "enq_report" and child.hasAttribute("id"):
                field = "Report ID"
                value = child.getAttribute("id")
                self.tree.insert("", "end", values=[field, value])
                self.parse_xml_to_treeview(child, field)
                continue
            if tag == "header":
                has_nested_report = any(r for r in child.getElementsByTagName("report"))
                if has_nested_report:
                    for sub in child.childNodes:
                        self.parse_xml_to_treeview(sub, parent_path)
                    continue
                else:
                    for sub in child.childNodes:
                        if sub.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                            sub_tag = sub.tagName
                            value = sub.firstChild.nodeValue.strip() if (sub.firstChild and sub.firstChild.nodeValue) else "-"
                            self.tree.insert("", "end", values=[sub_tag, value])
                    continue
            if tag == "summary":
                has_field_sum = False
                for sub in child.getElementsByTagName("enq_sum"):
                    field_sum_nodes = sub.getElementsByTagName("field_sum")
                    if field_sum_nodes:
                        has_field_sum = True
                        for fs in field_sum_nodes:
                            if fs.nodeType == xml.dom.minidom.Node.ELEMENT_NODE and fs.hasAttribute("name"):
                                field = fs.getAttribute("name").strip()
                                value = fs.firstChild.nodeValue.strip() if (fs.firstChild and fs.firstChild.nodeValue) else "-"
                                self.tree.insert("", "end", values=[field, value])
                if not has_field_sum:
                    for sub in child.getElementsByTagName("enq_sum"):
                        for item in sub.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName.strip()
                                value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                                self.tree.insert("", "end", values=[field, value])
                continue
            if tag == "section" and child.hasAttribute("title"):
                title = child.getAttribute("title").strip()
                self.tree.insert("", "end", values=[title, "-"], tags=("section_bold",))
                self.parse_xml_to_treeview(child, title)
                continue
            if tag == "record" and child.hasAttribute("seq"):
                seq = child.getAttribute("seq").strip()
                self.tree.insert("", "end", values=["Record", seq])
                self.parse_xml_to_treeview(child, f"record_{seq}")
                continue
            if tag == "data":
                caption = child.getAttribute("caption").strip()
                name = child.getAttribute("name").strip()
                field = caption if caption else name
                # Special handling for <data name="age">
                if name.lower() == "age":
                    age_fields = ["30", "60", "90", "120", "150", "180", "210"]
                    found_ages = {af: False for af in age_fields}
                    for item in [i for i in child.childNodes if i.nodeType == i.ELEMENT_NODE and i.tagName == "item"]:
                        age_name = item.getAttribute("name").strip() if item.hasAttribute("name") else ""
                        age_value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                        self.tree.insert("", "end", values=[f"age_{age_name}", age_value])
                        found_ages[age_name] = True
                    for af in age_fields:
                        if not found_ages[af]:
                            self.tree.insert("", "end", values=[f"age_{af}", "-"])
                else:
                    value = child.firstChild.nodeValue.strip() if (child.firstChild and child.firstChild.nodeValue) else "-"
                    self.tree.insert("", "end", values=[field, value])
                continue
            
            # --- Trade Reference (TR) logic ---
            if tag == "tr_report":
                if child.hasAttribute("type") and child.getAttribute("type").strip().upper() == "TR":
                    parent_node = self.tree.insert(parent_path, "end", values=["TRADE REFERENCE", "-"], tags=("section_bold",))
                    # Header
                    header_nodes = [n for n in child.childNodes if n.nodeType == n.ELEMENT_NODE and n.tagName == "header"]
                    for header in header_nodes:
                        for sub in header.childNodes:
                            if sub.nodeType == sub.ELEMENT_NODE:
                                k = sub.tagName.lower()
                                v = sub.firstChild.nodeValue.strip() if (sub.firstChild and sub.firstChild.nodeValue) else "-"
                                self.tree.insert(parent_node, "end", values=[k, v])
                    # Enquiries
                    enquiry_nodes = [n for n in child.childNodes if n.nodeType == n.ELEMENT_NODE and n.tagName == "enquiry"]
                    for idx, enq in enumerate(enquiry_nodes, start=1):
                        account_no = enq.getAttribute("account_no") if enq.hasAttribute("account_no") else f"Account {idx}"
                        try:
                            if isinstance(account_no, str) and ("e+" in account_no.lower() or "E+" in account_no):
                                account_no = str(decimal.Decimal(account_no.strip()))
                        except Exception:
                            pass
                        enq_node = self.tree.insert(parent_node, "end", values=["Account No", account_no])
                        # For each section in enquiry
                        for section in [n for n in enq.childNodes if n.nodeType == n.ELEMENT_NODE and n.tagName == "section"]:
                            sec_id = section.getAttribute("id") if section.hasAttribute("id") else ""
                            status = section.getAttribute("status") if section.hasAttribute("status") else "-"
                            section_node = self.tree.insert(enq_node, "end", values=[sec_id, status])
                            for data in [n for n in section.childNodes if n.nodeType == n.ELEMENT_NODE and n.tagName == "data"]:
                                dname = data.getAttribute("name").strip()
                                # Special handling for age
                                if dname.lower() == "age":
                                    age_fields = ["30", "60", "90", "120", "150", "180", "210"]
                                    found_ages = {af: False for af in age_fields}
                                    for item in [i for i in data.childNodes if i.nodeType == i.ELEMENT_NODE and i.tagName == "item"]:
                                        age_name = item.getAttribute("name").strip() if item.hasAttribute("name") else ""
                                        age_value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                                        self.tree.insert(section_node, "end", values=[f"age_{age_name}", age_value])
                                        found_ages[age_name] = True
                                    for af in age_fields:
                                        if not found_ages[af]:
                                            self.tree.insert(section_node, "end", values=[f"age_{af}", "-"])
                                else:
                                    text_val = data.firstChild.nodeValue.strip() if (data.firstChild and data.firstChild.nodeValue) else "-"
                                    # Fix account_no and reference formatting
                                    if dname in ["account_no", "reference"]:
                                        try:
                                            if isinstance(text_val, str) and ("e+" in text_val.lower() or "E+" in text_val):
                                                text_val = str(decimal.Decimal(text_val.strip()))
                                        except Exception:
                                            pass
                                    self.tree.insert(section_node, "end", values=[dname, text_val])
                        # Blank row after each enquiry
                        self.tree.insert(parent_node, "end", values=["", ""],tags=("section_bold",))
                
            # --- New CTOS XML logic below ---

            # SECTION A (new format)
            if tag == "section_a":
                title = child.getAttribute("title") if child.hasAttribute("title") else "SECTION A"
                self.tree.insert("", "end", values=[title, "-"], tags=("section_bold",))
                for record in child.getElementsByTagName("record"):
                    seq = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                    self.tree.insert("", "end", values=["Record", seq])
                    for item in record.childNodes:
                        if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                            field = item.tagName
                            # For nested addr_breakdown
                            if field == "addr_breakdown":
                                for addr_item in item.childNodes:
                                    if addr_item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                        subfield = addr_item.tagName
                                        subvalue = addr_item.firstChild.nodeValue.strip() if (addr_item.firstChild and addr_item.firstChild.nodeValue) else "-"
                                        self.tree.insert("", "end", values=[subfield, subvalue])
                                continue
                            value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                            self.tree.insert("", "end", values=[field, value])
                continue

            # SECTION B (new format)
            if tag == "section_b":
                title = child.getAttribute("title") if child.hasAttribute("title") else "SECTION B"
                self.tree.insert("", "end", values=[title, "-"], tags=("section_bold",))
                # Handle <history> nodes
                for history in child.getElementsByTagName("history"):
                    year = history.getAttribute("year") if history.hasAttribute("year") else "-"
                    seq = history.getAttribute("seq") if history.hasAttribute("seq") else "-"
                    self.tree.insert("", "end", values=["history_year", year])
                    self.tree.insert("", "end", values=["history_seq", seq])
                    for period in history.getElementsByTagName("period"):
                        month = period.getAttribute("month") if period.hasAttribute("month") else "-"
                        self.tree.insert("", "end", values=["period_month", month])
                        for entity in period.getElementsByTagName("entity"):
                            etype = entity.getAttribute("type") if entity.hasAttribute("type") else "-"
                            value = entity.getAttribute("value") if entity.hasAttribute("value") else "-"
                            self.tree.insert("", "end", values=[f"entity_{etype}", value])
                # Handle <record> nodes (old style)
                for record in child.getElementsByTagName("record"):
                    seq = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                    self.tree.insert("", "end", values=["Record", seq])
                    for item in record.childNodes:
                        if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                            field = item.tagName
                            value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                            self.tree.insert("", "end", values=[field, value])
                continue

            # SECTION C (new format, including broken/nested)
            if tag == "section_c":
                title = child.getAttribute("title") if child.hasAttribute("title") else "SECTION C"
                self.tree.insert("", "end", values=[title, "-"], tags=("section_bold",))
                for record in child.getElementsByTagName("record"):
                    seq = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                    self.tree.insert("", "end", values=["Record", seq])
                    def flatten_record(node):
                        for item in node.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName
                                # If the node has element children, flatten recursively
                                if any(c.nodeType == xml.dom.minidom.Node.ELEMENT_NODE for c in item.childNodes):
                                    flatten_record(item)
                                else:
                                    value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                                    self.tree.insert("", "end", values=[field, value])
                    flatten_record(record)
                continue
            
                        # SECTION D (new format)
            if tag == "section_d":
                title = child.getAttribute("title") if child.hasAttribute("title") else "SECTION D"
                self.tree.insert("", "end", values=[title, "-"], tags=("section_bold",))
                for record in child.getElementsByTagName("record"):
                    seq = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                    self.tree.insert("", "end", values=["Record", seq])
                    for item in record.childNodes:
                        if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                            field = item.tagName
                            # Flatten <action>
                            if field == "action":
                                for subitem in item.childNodes:
                                    if subitem.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                        subfield = f"action_{subitem.tagName}"
                                        subvalue = subitem.firstChild.nodeValue.strip() if (subitem.firstChild and subitem.firstChild.nodeValue) else "-"
                                        self.tree.insert("", "end", values=[subfield, subvalue])
                                continue
                            # Flatten <settlement>
                            if field == "settlement":
                                for subitem in item.childNodes:
                                    if subitem.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                        subfield = f"settlement_{subitem.tagName}"
                                        subvalue = subitem.firstChild.nodeValue.strip() if (subitem.firstChild and subitem.firstChild.nodeValue) else "-"
                                        self.tree.insert("", "end", values=[subfield, subvalue])
                                continue
                            value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                            self.tree.insert("", "end", values=[field, value])
                continue
            
                        # SECTION D2 (new format)
            if tag == "section_d2":
                title = child.getAttribute("title") if child.hasAttribute("title") else "SECTION D2"
                # Insert section name as bold
                section_id = self.tree.insert("", "end", values=[title, "-"], tags=("section_bold",))
                for record in child.getElementsByTagName("record"):
                    seq = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                    self.tree.insert(section_id, "end", values=["Record", seq])
                    def flatten_record(node, prefix=""):
                        for item in node.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName
                                # Handle nested elements
                                if field in ["lawyer", "cedcon", "settlement", "latest_status", "other_defendants"]:
                                    if field == "other_defendants":
                                        for od in item.getElementsByTagName("other_defendant"):
                                            od_seq = od.getAttribute("seq") if od.hasAttribute("seq") else ""
                                            for od_item in od.childNodes:
                                                if od_item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                                    od_field = f"other_defendant_{od_seq}_{od_item.tagName}"
                                                    od_value = od_item.firstChild.nodeValue.strip() if (od_item.firstChild and od_item.firstChild.nodeValue) else "-"
                                                    self.tree.insert(section_id, "end", values=[od_field, od_value])
                                    else:
                                        for subitem in item.childNodes:
                                            if subitem.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                                subfield = f"{field}_{subitem.tagName}"
                                                subvalue = subitem.firstChild.nodeValue.strip() if (subitem.firstChild and subitem.firstChild.nodeValue) else "-"
                                                self.tree.insert(section_id, "end", values=[subfield, subvalue])
                                else:
                                    value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                                    self.tree.insert(section_id, "end", values=[field, value])
                    flatten_record(record)
                continue
            
            # SECTION E (new format)
            if tag == "section_e":
                title = child.getAttribute("title") if child.hasAttribute("title") else "SECTION E"
                self.tree.insert("", "end", values=[title, "-"], tags=("section_bold",))
                for enquiry in child.getElementsByTagName("enquiry"):
                    seq = enquiry.getAttribute("seq") if enquiry.hasAttribute("seq") else ""
                    account_no = enquiry.getAttribute("account_no") if enquiry.hasAttribute("account_no") else "-"
                    tref_id = enquiry.getAttribute("tref_id") if enquiry.hasAttribute("tref_id") else "-"
                    self.tree.insert("", "end", values=["Enquiry Seq", seq])
                    self.tree.insert("", "end", values=["Account No", account_no])
                    self.tree.insert("", "end", values=["Tref ID", tref_id])
                    # subject
                    for subject in enquiry.getElementsByTagName("subject"):
                        for item in subject.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName
                                value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                                self.tree.insert("", "end", values=[field, value])
                    # relationship
                    for rel in enquiry.getElementsByTagName("relationship"):
                        for item in rel.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName
                                value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                                self.tree.insert("", "end", values=[field, value])
                    # account_status
                    for acc in enquiry.getElementsByTagName("account_status"):
                        for item in acc.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName
                                # For <age> node, flatten children
                                if field == "age":
                                    for age_item in item.childNodes:
                                        if age_item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                            subfield = age_item.tagName
                                            subvalue = age_item.firstChild.nodeValue.strip() if (age_item.firstChild and age_item.firstChild.nodeValue) else "-"
                                            self.tree.insert("", "end", values=[subfield, subvalue])
                                    continue
                                value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                                self.tree.insert("", "end", values=[field, value])
                    # legal_action
                    for legal in enquiry.getElementsByTagName("legal_action"):
                        status = legal.getAttribute("status") if legal.hasAttribute("status") else "-"
                        self.tree.insert("", "end", values=["legal_action_status", status])
                        for item in legal.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName
                                # For reminder_letter, demand_letter_by_company, demand_letter_by_lawyer
                                for subitem in item.childNodes:
                                    if subitem.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                        subfield = f"{field}_{subitem.tagName}"
                                        subvalue = subitem.firstChild.nodeValue.strip() if (subitem.firstChild and subitem.firstChild.nodeValue) else "-"
                                        self.tree.insert("", "end", values=[subfield, subvalue])
                                # If the item has text directly
                                if item.firstChild and item.firstChild.nodeType == xml.dom.minidom.Node.TEXT_NODE:
                                    value = item.firstChild.nodeValue.strip()
                                    self.tree.insert("", "end", values=[field, value])
                    # referee_contact
                    for refc in enquiry.getElementsByTagName("referee_contact"):
                        for item in refc.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName
                                value = item.firstChild.nodeValue.strip() if (item.firstChild and item.firstChild.nodeValue) else "-"
                                self.tree.insert("", "end", values=[field, value])
                continue

            # --- fallback: process children ---
            self.parse_xml_to_treeview(child, parent_path)
             # At the end of parse_xml_to_treeview (after tree creation), add:
        self.tree.tag_configure("section_bold", font=("Segoe UI", 11, "bold"))

    def go_to_next(self):
        if self.all_accounts and self.current_index < len(self.all_accounts) - 1:
            self.current_index += 1
            self.account_var.set(self.all_accounts[self.current_index])
            self.display_data()

    def go_to_previous(self):
        if self.all_accounts and self.current_index > 0:
            self.current_index -= 1
            self.account_var.set(self.all_accounts[self.current_index])
            self.display_data()

    def search_nu_ptl(self, event=None):
        search_value = self.search_var.get().strip()
        if self.filtered_data is not None and search_value:
            matches = self.filtered_data[self.filtered_data["NU_PTL"].astype(str).str.contains(search_value)]
            if not matches.empty:
                self.filtered_data = matches.reset_index(drop=True)
                self.current_index = 0
                self.display_data()

    def convert_new_ctos_to_excel(self):
        # Ask for confirmation before starting the export
        if not messagebox.askyesno("Confirm Export", "Are you sure you want to download the New CTOS Excel?"):
            return
        self.is_converting = True
        self.show_progress_popup()
        threading.Thread(target=self.convert_new_ctos_to_excel_thread, daemon=True).start()


    def convert_new_ctos_to_excel_thread(self):
        
        # Helper function to detect new CTOS format
        def is_new_ctos_xml(xml_str):
            """Check if XML contains new CTOS format indicators"""
            new_ctos_tags = [
                "<section_a", "<section_b", "<section_c", 
                "<section_d", "<section_d2", "<section_d3", "<section_d4",
                "<section_e", "<section_etr_plus", "<section_etl",
                "<history", "<period", "<enquiry"  # New format specific elements
            ]
            return any(tag in xml_str.lower() for tag in new_ctos_tags)

        # New CTOS section columns with proper headers for each section
        new_section_columns = {
            "Header&Summary": ["NU_PTL", "USR", "CMP", "ACC", "TEL", "FAX", "EDT", "ETM", "EST", "NAME", "IC", "NIC", "IDX", "REF"],
            "Section-A": ["NU_PTL", "RID", "NAME", "IC", "NIC", "ADDR", "SRC", "BDT"],
            "Section-B1": ["NU_PTL", "RID", "CO", "ADREG", "LOC", "OBJ", "INC", "LST", "APP", "RSN", "NAME", "NIC", "ADDR", "POS", "CPO", "PD", "SH", "TSH", "RM"],
            "Section-B2": ["NU_PTL", "RID", "TTL", "NAME", "ALS", "IC", "NIC", "REF", "FIRM", "RM1", "RM2", "RM3", "AMT", "ENT"],
            "Section-C1": ["NU_PTL", "RID", "DETAILS"],
            "Section-D1": [
                "NU_PTL", "RID", "RPTTYPE", "STATUS", "TITLE", "SPECIAL_REMARK", "NAME", "NAME_MATCH", "ALIAS", "ADDR",
                "IC_LCNO", "NIC_BRNO", "NIC_BRNO_MATCH", "CASE_NO", "COURT_DETAIL", "FIRM", "PLAINTIFF",
                "ACTION_DATE", "ACTION_SOURCE_DETAIL", "HEAR_DATE", "AMOUNT", "REMARK", "LAWYER", "CEDCON",
                "SETTLEMENT_CODE", "SETTLEMENT_DATE", "SETTLEMENT_SOURCE", "SETTLEMENT_SOURCE_DATE",
                "LATEST_STATUS", "SUBJECT_CMT", "CRA_CMT"
            ],
            "Section-D2": [
                "NU_PTL", "RID", "TITLE", "SPECIAL_REMARK", "NAME", "ADDR", "CASE_NO", "COURT_DETAIL", "FIRM",
                "ACTION_DATE", "ACTION_SOURCE_DETAIL", "HEAR_DATE", "AMOUNT", "REMARK", "LAWYER_NAME", 
                "LAWYER_ADD1", "LAWYER_ADD2", "LAWYER_REF", "LATEST_STATUS_CODE", "OTHER_DEFENDANT_1_NAME", 
                "SUBJECT_CMT", "CRA_CMT"
            ],
            "Section-D3": ["NU_PTL", "RID", "DETAILS"],
            "Section-D4": ["NU_PTL", "RID", "DETAILS"],
            "Section-E1": ["NU_PTL", "RID", "DETAILS"],
            "Section-E2": [
                "NU_PTL", "RID", "ENQUIRY_TYPE", "REF_COM_NAME", "REF_COM_BUS", "PARTY_TYPE", "NIC_BRNO", "NAME", 
                "TREF_DATE", "REL_TYPE", "REL_STATUS", "ACCOUNT_NO", "REL_SYEAR", "REL_SMONTH", "REL_SDAY", 
                "REL_REMARK", "STATEMENT_DATE", "ACCOUNT_RATING", "ACCOUNT_TERM", "ACCOUNT_LIMIT", "ACCOUNT_STATUS", 
                "DEBTOR_NAME", "DEBTOR_NIC_BRNO", "DEBT_TYPE", "LAST_PAID_AMOUNT", "AGE_30", "AGE_60", "AGE_90", 
                "AGE_120", "AGE_150", "AGE_180", "AGE_OVER_180", "CONTACT_ADD", "CONTACT_TELNO", 
                "CONTACT_NATURE_OF_BUSINESS", "CONTACT_FAXNO", "CONTACT_TYPE"
            ]
        }

        new_sheets_data = {k: [] for k in new_section_columns}
        total = len(self.filtered_data)
        new_ctos_count = 0

        for index, (_, row) in enumerate(self.filtered_data.iterrows()):
            nu_ptl = row.get("NU_PTL", f"Row{index}")
            xml_data = clean_malformed_xml(row.get("XML", ""))
            if pd.isna(xml_data) or not str(xml_data).strip():
                continue
            
            # Ensure only up to first </report> is parsed
            xml_data = extract_first_report(xml_data)
            
            # Check if this is actually new CTOS format
            if not is_new_ctos_xml(xml_data):
                continue
                
            new_ctos_count += 1
            
            try:
                dom = xml.dom.minidom.parseString(xml_data)
                root = dom.documentElement

                # --- Header&Summary ---
                header_record = {col: "" for col in new_section_columns["Header&Summary"]}
                header_record["NU_PTL"] = nu_ptl
                for header in root.getElementsByTagName("header"):
                    for node in header.childNodes:
                        if node.nodeType == node.ELEMENT_NODE:
                            tag = node.tagName.strip().upper()
                            if tag == "USER":
                                header_record["USR"] = node.firstChild.nodeValue.strip() if node.firstChild else ""
                            elif tag == "COMPANY":
                                header_record["CMP"] = node.firstChild.nodeValue.strip() if node.firstChild else ""
                            elif tag == "ACCOUNT":
                                header_record["ACC"] = node.firstChild.nodeValue.strip() if node.firstChild else ""
                            elif tag == "TEL":
                                header_record["TEL"] = node.firstChild.nodeValue.strip() if node.firstChild else ""
                            elif tag == "FAX":
                                header_record["FAX"] = node.firstChild.nodeValue.strip() if node.firstChild else ""
                            elif tag == "ENQ_DATE":
                                header_record["EDT"] = node.firstChild.nodeValue.strip() if node.firstChild else ""
                            elif tag == "ENQ_TIME":
                                header_record["ETM"] = node.firstChild.nodeValue.strip() if node.firstChild else ""
                            elif tag == "ENQ_STATUS":
                                header_record["EST"] = node.firstChild.nodeValue.strip() if node.firstChild else ""
                for summary in root.getElementsByTagName("summary"):
                    for enq_sum in summary.getElementsByTagName("enq_sum"):
                        for item in enq_sum.childNodes:
                            if item.nodeType == item.ELEMENT_NODE:
                                tag = item.tagName.strip().upper()
                                if tag == "NAME":
                                    header_record["NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "IC_LCNO":
                                    header_record["IC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NIC_BRNO":
                                    header_record["NIC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "DD_INDEX":
                                    header_record["IDX"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REF_NO":
                                    header_record["REF"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                new_sheets_data["Header&Summary"].append(header_record)

                # --- Section-A ---
                for section_a in root.getElementsByTagName("section_a"):
                    for record in section_a.getElementsByTagName("record"):
                        rec = {col: "" for col in new_section_columns["Section-A"]}
                        rec["NU_PTL"] = nu_ptl
                        rec["RID"] = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                        for item in record.childNodes:
                            if item.nodeType == item.ELEMENT_NODE:
                                tag = item.tagName.strip().upper()
                                if tag == "NAME":
                                    rec["NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "IC_LCNO":
                                    rec["IC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NIC_BRNO":
                                    rec["NIC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ADDR":
                                    rec["ADDR"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "SOURCE":
                                    rec["SRC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "BIRTH_DATE":
                                    rec["BDT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                        new_sheets_data["Section-A"].append(rec)

                # --- Section-B1 (SECTION C - DIRECTORSHIPS maps to B1) ---
                for section_c in root.getElementsByTagName("section_c"):
                    for record in section_c.getElementsByTagName("record"):
                        rec = {col: "" for col in new_section_columns["Section-B1"]}
                        rec["NU_PTL"] = nu_ptl
                        rec["RID"] = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                        for item in record.childNodes:
                            if item.nodeType == item.ELEMENT_NODE:
                                tag = item.tagName.strip().upper()
                                if tag == "COMPANY_NAME":
                                    rec["CO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ADDITIONAL_REGISTRATION_NO":
                                    rec["ADREG"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "LOCAL":
                                    rec["LOC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "OBJECT":
                                    rec["OBJ"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "INCDATE":
                                    rec["INC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "EXPDATE":
                                    rec["LST"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NAME":
                                    rec["NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NIC_BRNO":
                                    rec["NIC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ADDR":
                                    rec["ADDR"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "POSITION":
                                    rec["POS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "CPO_DATE":
                                    rec["CPO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REMARK":
                                    rec["RM"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                        new_sheets_data["Section-B1"].append(rec)
                    
                # --- Section-B2 (Section B records - Internal List) ---
                for section_b in root.getElementsByTagName("section_b"):
                    for record in section_b.getElementsByTagName("record"):
                        rec = {col: "" for col in new_section_columns["Section-B2"]}
                        rec["NU_PTL"] = nu_ptl
                        rec["RID"] = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                        for item in record.childNodes:
                            if item.nodeType == item.ELEMENT_NODE:
                                tag = item.tagName.strip().upper()
                                if tag == "TITLE":
                                    rec["TTL"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NAME":
                                    rec["NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ALIAS":
                                    rec["ALS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "IC_LCNO":
                                    rec["IC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NIC_BRNO":
                                    rec["NIC"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REF":
                                    rec["REF"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "FIRM":
                                    rec["FIRM"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REMARK1":
                                    rec["RM1"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REMARK2":
                                    rec["RM2"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REMARK3":
                                    rec["RM3"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "AMOUNT":
                                    rec["AMT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ENTRY":
                                    rec["ENT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                        new_sheets_data["Section-B2"].append(rec)

                # --- Section-C1 (Empty in your example, but add placeholder) ---
                # Creating placeholder record for empty section
                rec = {col: "" for col in new_section_columns["Section-C1"]}
                rec["NU_PTL"] = nu_ptl
                rec["RID"] = "1"
                rec["DETAILS"] = "Empty section in new CTOS format"
                new_sheets_data["Section-C1"].append(rec)

                # --- Section-D1 ---
                for section_d in root.getElementsByTagName("section_d"):
                    for record in section_d.getElementsByTagName("record"):
                        rec = {col: "" for col in new_section_columns["Section-D1"]}
                        rec["NU_PTL"] = nu_ptl
                        rec["RID"] = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                        rec["RPTTYPE"] = record.getAttribute("rpttype") if record.hasAttribute("rpttype") else ""
                        rec["STATUS"] = record.getAttribute("status") if record.hasAttribute("status") else ""
                        for item in record.childNodes:
                            if item.nodeType == item.ELEMENT_NODE:
                                tag = item.tagName.strip().upper()
                                if tag == "TITLE":
                                    rec["TITLE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "SPECIAL_REMARK":
                                    rec["SPECIAL_REMARK"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NAME":
                                    rec["NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    rec["NAME_MATCH"] = item.getAttribute("match") if item.hasAttribute("match") else ""
                                elif tag == "ALIAS":
                                    rec["ALIAS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ADDR":
                                    rec["ADDR"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "IC_LCNO":
                                    rec["IC_LCNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NIC_BRNO":
                                    rec["NIC_BRNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    rec["NIC_BRNO_MATCH"] = item.getAttribute("match") if item.hasAttribute("match") else ""
                                elif tag == "CASE_NO":
                                    rec["CASE_NO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "COURT_DETAIL":
                                    rec["COURT_DETAIL"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "FIRM":
                                    rec["FIRM"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "PLAINTIFF":
                                    rec["PLAINTIFF"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ACTION":
                                    for subitem in item.childNodes:
                                        if subitem.nodeType == subitem.ELEMENT_NODE:
                                            subtag = subitem.tagName.strip().upper()
                                            if subtag == "DATE":
                                                rec["ACTION_DATE"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                            elif subtag == "SOURCE_DETAIL":
                                                rec["ACTION_SOURCE_DETAIL"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                elif tag == "HEAR_DATE":
                                    rec["HEAR_DATE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "AMOUNT":
                                    rec["AMOUNT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REMARK":
                                    rec["REMARK"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "LAWYER":
                                    rec["LAWYER"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "CEDCON":
                                    rec["CEDCON"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "SETTLEMENT":
                                    for subitem in item.childNodes:
                                        if subitem.nodeType == subitem.ELEMENT_NODE:
                                            subtag = subitem.tagName.strip().upper()
                                            if subtag == "CODE":
                                                rec["SETTLEMENT_CODE"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                            elif subtag == "DATE":
                                                rec["SETTLEMENT_DATE"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                            elif subtag == "SOURCE":
                                                rec["SETTLEMENT_SOURCE"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                            elif subtag == "SOURCE_DATE":
                                                rec["SETTLEMENT_SOURCE_DATE"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                elif tag == "LATEST_STATUS":
                                    rec["LATEST_STATUS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "SUBJECT_CMT":
                                    rec["SUBJECT_CMT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "CRA_CMT":
                                    rec["CRA_CMT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                        new_sheets_data["Section-D1"].append(rec)

                # --- Section-D2 ---
                for section_d2 in root.getElementsByTagName("section_d2"):
                    for record in section_d2.getElementsByTagName("record"):
                        rec = {col: "" for col in new_section_columns["Section-D2"]}
                        rec["NU_PTL"] = nu_ptl
                        rec["RID"] = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                        
                        for item in record.childNodes:
                            if item.nodeType == item.ELEMENT_NODE:
                                tag = item.tagName.strip().upper()
                                if tag == "TITLE":
                                    rec["TITLE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "SPECIAL_REMARK":
                                    rec["SPECIAL_REMARK"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "NAME":
                                    rec["NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ADDR":
                                    rec["ADDR"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "CASE_NO":
                                    rec["CASE_NO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "COURT_DETAIL":
                                    rec["COURT_DETAIL"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "FIRM":
                                    rec["FIRM"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "ACTION":
                                    for subitem in item.childNodes:
                                        if subitem.nodeType == subitem.ELEMENT_NODE:
                                            subtag = subitem.tagName.strip().upper()
                                            if subtag == "DATE":
                                                rec["ACTION_DATE"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                            elif subtag == "SOURCE_DETAIL":
                                                rec["ACTION_SOURCE_DETAIL"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                elif tag == "HEAR_DATE":
                                    rec["HEAR_DATE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "AMOUNT":
                                    rec["AMOUNT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REMARK":
                                    rec["REMARK"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "LAWYER":
                                    for subitem in item.childNodes:
                                        if subitem.nodeType == subitem.ELEMENT_NODE:
                                            subtag = subitem.tagName.strip().upper()
                                            if subtag == "NAME":
                                                rec["LAWYER_NAME"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                            elif subtag == "ADD1":
                                                rec["LAWYER_ADD1"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                            elif subtag == "ADD2":
                                                rec["LAWYER_ADD2"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                            elif subtag == "REF":
                                                rec["LAWYER_REF"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                elif tag == "LATEST_STATUS":
                                    for subitem in item.childNodes:
                                        if subitem.nodeType == subitem.ELEMENT_NODE:
                                            subtag = subitem.tagName.strip().upper()
                                            if subtag == "CODE":
                                                rec["LATEST_STATUS_CODE"] = subitem.firstChild.nodeValue.strip() if subitem.firstChild else ""
                                elif tag == "OTHER_DEFENDANTS":
                                    for other_def in item.getElementsByTagName("other_defendant"):
                                        if other_def.getAttribute("seq") == "1":
                                            for od_item in other_def.childNodes:
                                                if od_item.nodeType == od_item.ELEMENT_NODE and od_item.tagName.strip().upper() == "NAME":
                                                    rec["OTHER_DEFENDANT_1_NAME"] = od_item.firstChild.nodeValue.strip() if od_item.firstChild else ""
                                elif tag == "SUBJECT_CMT":
                                    rec["SUBJECT_CMT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "CRA_CMT":
                                    rec["CRA_CMT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                        
                        new_sheets_data["Section-D2"].append(rec)

                # --- Section-D3 & D4 (usually empty) ---
                for section_tag, sheet_name in [("section_d3", "Section-D3"), ("section_d4", "Section-D4")]:
                    for section in root.getElementsByTagName(section_tag):
                        for record in section.getElementsByTagName("record"):
                            rec = {col: "" for col in new_section_columns[sheet_name]}
                            rec["NU_PTL"] = nu_ptl
                            rec["RID"] = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                            rec["DETAILS"] = "Empty section"
                            new_sheets_data[sheet_name].append(rec)

                # --- Section-E1 (Empty in your example, but add placeholder) ---
                # Creating placeholder record for empty section
                rec = {col: "" for col in new_section_columns["Section-E1"]}
                rec["NU_PTL"] = nu_ptl
                rec["RID"] = "1"
                rec["DETAILS"] = "Empty section in new CTOS format"
                new_sheets_data["Section-E1"].append(rec)

                # --- Section-E2 (SECTION E - TRADE REFEREES) ---
                for section_e in root.getElementsByTagName("section_e"):
                    for enquiry in section_e.getElementsByTagName("enquiry"):
                        rec = {col: "" for col in new_section_columns["Section-E2"]}
                        rec["NU_PTL"] = nu_ptl
                        rec["RID"] = enquiry.getAttribute("seq") if enquiry.hasAttribute("seq") else ""
                        rec["ENQUIRY_TYPE"] = "Trade Referee"
                        
                        # Subject details
                        for subject in enquiry.getElementsByTagName("subject"):
                            for item in subject.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    value = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    if tag == "REF_COM_NAME":
                                        rec["REF_COM_NAME"] = value
                                    elif tag == "REF_COM_BUS":
                                        rec["REF_COM_BUS"] = value
                                    elif tag == "PARTY_TYPE":
                                        rec["PARTY_TYPE"] = value
                                    elif tag == "NIC_BRNO":
                                        rec["NIC_BRNO"] = value
                                    elif tag == "NAME":
                                        rec["NAME"] = value
                                    elif tag == "TREF_DATE":
                                        rec["TREF_DATE"] = value
                        
                        # Relationship details
                        for relationship in enquiry.getElementsByTagName("relationship"):
                            for item in relationship.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    value = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    if tag == "REL_TYPE":
                                        rec["REL_TYPE"] = value
                                    elif tag == "REL_STATUS":
                                        rec["REL_STATUS"] = value
                                    elif tag == "ACCOUNT_NO":
                                        rec["ACCOUNT_NO"] = value
                                    elif tag == "REL_SYEAR":
                                        rec["REL_SYEAR"] = value
                                    elif tag == "REL_SMONTH":
                                        rec["REL_SMONTH"] = value
                                    elif tag == "REL_SDAY":
                                        rec["REL_SDAY"] = value
                                    elif tag == "REMARK":
                                        rec["REL_REMARK"] = value
                        
                        # Account status details
                        for account_status in enquiry.getElementsByTagName("account_status"):
                            for item in account_status.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    if tag == "STATEMENT_DATE":
                                        rec["STATEMENT_DATE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ACCOUNT_RATING":
                                        rec["ACCOUNT_RATING"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ACCOUNT_TERM":
                                        rec["ACCOUNT_TERM"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ACCOUNT_LIMIT":
                                        rec["ACCOUNT_LIMIT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ACCOUNT_STATUS":
                                        rec["ACCOUNT_STATUS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "DEBTOR_NAME":
                                        rec["DEBTOR_NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "DEBTOR_NIC_BRNO":
                                        rec["DEBTOR_NIC_BRNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "DEBT_TYPE":
                                        rec["DEBT_TYPE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "LAST_PAID_AMOUNT":
                                        rec["LAST_PAID_AMOUNT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "AGE":
                                        # Handle age sub-elements
                                        for age_item in item.childNodes:
                                            if age_item.nodeType == age_item.ELEMENT_NODE:
                                                age_tag = age_item.tagName.strip().upper()
                                                age_value = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""
                                                if age_tag == "AGE_30":
                                                    rec["AGE_30"] = age_value
                                                elif age_tag == "AGE_60":
                                                    rec["AGE_60"] = age_value
                                                elif age_tag == "AGE_90":
                                                    rec["AGE_90"] = age_value
                                                elif age_tag == "AGE_120":
                                                    rec["AGE_120"] = age_value
                                                elif age_tag == "AGE_150":
                                                    rec["AGE_150"] = age_value
                                                elif age_tag == "AGE_180":
                                                    rec["AGE_180"] = age_value
                                                elif age_tag == "AGE_OVER_180":
                                                    rec["AGE_OVER_180"] = age_value
                        
                        # Referee contact details
                        for referee_contact in enquiry.getElementsByTagName("referee_contact"):
                            for item in referee_contact.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    value = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    if tag == "CONTACT_ADD":
                                        rec["CONTACT_ADD"] = value
                                    elif tag == "CONTACT_TELNO":
                                        rec["CONTACT_TELNO"] = value
                                    elif tag == "CONTACT_NATURE_OF_BUSINESS":
                                        rec["CONTACT_NATURE_OF_BUSINESS"] = value
                                    elif tag == "CONTACT_FAXNO":
                                        rec["CONTACT_FAXNO"] = value
                                    elif tag == "CONTACT_TYPE":
                                        rec["CONTACT_TYPE"] = value
                        
                        new_sheets_data["Section-E2"].append(rec)

            except Exception as e:
                msg = f"Error parsing XML for NU_PTL {nu_ptl}: {str(e)}"
                self.after(0, self.append_error, msg)
                continue

            if index % 10 == 0 or index + 1 == total:
                progress = (index + 1) / total
                self.after(0, self.update_progress, progress, index + 1, total)

        # Ensure all sheets exist and have at least header row
        for sheet in ["Section-A", "Section-B1", "Section-B2", "Section-C1", "Section-D1", "Section-D2", "Section-D3", "Section-D4", "Section-E1", "Section-E2"]:
            if sheet not in new_sheets_data:
                new_sheets_data[sheet] = []
            if not new_sheets_data[sheet]:
                new_sheets_data[sheet].append({col: "" for col in new_section_columns[sheet]})

        # Export to Excel
        self.after(0, self.update_status, "Writing to Excel...")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        new_save_path = os.path.join(downloads_folder, f"new_ctos_report_{timestamp}.xlsx")
        if any(len(records) > 0 for records in new_sheets_data.values()):
            with pd.ExcelWriter(new_save_path, engine="openpyxl") as writer:
                for sheet_name, records in new_sheets_data.items():
                    if records:
                        df = pd.DataFrame(records)
                        df = df.reindex(columns=new_section_columns[sheet_name])
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
        self.after(0, self.update_status, "Export successful!")
        self.after(0, self.destroy_popup)

    def convert_to_excel(self):
        # Ask for confirmation before starting the export
        if not messagebox.askyesno("Confirm Export", "Are you sure you want to download the Old CTOS Excel?"):
            return
        self.is_converting = True
        self.show_progress_popup()
        threading.Thread(target=self.convert_to_excel_thread, daemon=True).start()

    def update_progress(self, progress, index, total):
        self.progress_bar.set(progress)
        self.status_label.configure(text=f"Processing {index} of {total}")
        # Removed: self.popup.update()  # <-- Avoid explicit update call here
        
    def convert_to_excel_thread(self):
        
        # Helper function to detect old CTOS format
        def is_old_ctos_xml(xml_str):
            """Check if XML contains old CTOS format indicators"""
            old_ctos_tags = [
                "<section title=", "caption=", "<record seq=", 
                "<tr_report type=\"TR\"", "<data caption=", "<data name="
            ]
            # Old format usually doesn't have new format tags
            new_format_indicators = ["<section_a", "<section_d2", "<section_d4", "history", "period"]
            
            has_old_indicators = any(tag in xml_str for tag in old_ctos_tags)
            has_new_indicators = any(tag in xml_str.lower() for tag in new_format_indicators)
            
            return has_old_indicators and not has_new_indicators
        # Define columns for old CTOS sections
        old_section_columns = {
            "Header&Summary": ["NU_PTL", "user", "company", "account", "tel", "fax", "enq_date", 
                                "enq_time", "enq_status", "IC_LCNO", "NIC_BRNO", "NAME", "ALIAS", "STAT", "REF"],
            "Section-A": ["NU_PTL", "Record_ID", "ICNO", "MATCH", "NEWIC", "MATCH1", "NAME", "MATCH2", "ADDR", "ADDR1", "REMARK"],
            "Section-B": ["NU_PTL", "Record_ID", "CODE", "NAME", "MATCH", "ALIAS", "IC_LCNO", "NIC_BRNO", 
                        "REF", "CONUM", "CONAME", "REMARK", "REMARK2", "REMARK3", "AMOUNT", "ENTRY"],
            "Section-C": ["NU_PTL", "Record_ID", "EXTRANAME", "EXTRALOCAL", "EXTRALOCALA", "OBJECT", 
                        "INCORPORATION", "COMPANY PAIDUP", "SEARCH DATED", "EXTRALASTDOC", "NAME", "IC_LCNO", 
                        "NIC_BRNO", "STATUS", "SHARES", "REAMARK", "APPOINTED", "RESIGNED", "ADDRESS"],
            "Section-D": ["NU_PTL", "Record_ID", "ZTITLE", "ZSPECIAL", "NAME", "MATCH", "ALIAS", "I/C NO", 
                        "NEW IC", "REMARK", "ADDRESS", "FIRM", "PLAINTIFF", "CASE NO", "ZCOURT", "ACTION DATE", 
                        "ZNTPAP", "HEARING DATE", "AMOUNT", "SOLCTR", "LAWADD1", "TEL", "LAWADD2", "REF", 
                        "LAWADD3", "PLAINTIFF CONTACT", "CEDCONADD1", "CEDCONADD2", "CEDCONADD3"],
            "Section-E": ["NU_PTL", "Record_ID", "REFEREE", "INCORPORATION DATE", "NATURE OF BUSINESS", "ADDRESS", "TR_URL"],
            "Trade Reference": [
                "NU_PTL", "Row ID", "Date", "Req Name", "Req Com Name", "Req Com Addr", "Ref Com Name", "Ref Com Bus",
                "Report No", "IC LCNO", "NIC BRNO", "Name", "Enquiry Account No", 
                "Rel Type", "Rel Status", "Rel SYear", "Rel SMonth", "Rel SDay", 
                "Acc Account No", "Acc Statement Date", "Acc Rating", "Acc Term", "Acc Limit", "Acc Status",
                "Acc Debtor Name", "Acc Debtor IC LCNO", "Acc Debtor NIC BRNO", "Acc Address", "Acc Debt Type",
                "Acc Last Paid Amount", "Acc Age 30", "Acc Age 60", "Acc Age 90", "Acc Age 120", "Acc Age 150", 
                "Acc Age 180", "Acc Age Over 180", "Legal Action Status", "Contact Ref", "Contact Name", 
                "Contact Add", "Contact Telno", "Contact Nature Of Business", "Contact Faxno", "Contact Email", "Contact Type"
            ]
        }

        old_sheets_data = {k: [] for k in old_section_columns}
        trade_reference_data = []
        total = len(self.filtered_data)
        old_ctos_count = 0

        for index, (_, row) in enumerate(self.filtered_data.iterrows()):
            nu_ptl = row.get("NU_PTL", f"Row{index}")
            xml_data = clean_malformed_xml(row.get("XML", ""))
            if pd.isna(xml_data) or not str(xml_data).strip():
                continue
                
            # Ensure only up to first </report> is parsed
            xml_data = extract_first_report(xml_data)
            
            # Check if this is actually old CTOS format
            if not is_old_ctos_xml(xml_data):
                continue
                
            old_ctos_count += 1
            
            try:
                dom = xml.dom.minidom.parseString(xml_data)
                root = dom.documentElement

                # --- Header & Summary ---
                header_record = {col: "-" for col in old_section_columns["Header&Summary"]}
                header_record["NU_PTL"] = nu_ptl
                for header in root.getElementsByTagName("header"):
                    for node in header.childNodes:
                        if node.nodeType == node.ELEMENT_NODE:
                            tag = node.tagName.strip()
                            if tag in old_section_columns["Header&Summary"]:
                                header_record[tag] = get_node_text(node)
                for summary in root.getElementsByTagName("summary"):
                    for enq_sum in summary.getElementsByTagName("enq_sum"):
                        for fs in enq_sum.getElementsByTagName("field_sum"):
                            if fs.hasAttribute("name"):
                                field = fs.getAttribute("name").strip()
                                if field in old_section_columns["Header&Summary"]:
                                    header_record[field] = get_node_text(fs)
                old_sheets_data["Header&Summary"].append(header_record)

                # --- Sections A to E ---
                for section in root.getElementsByTagName("section"):
                    sec_id = section.getAttribute("id").strip().upper()
                    section_key = f"Section-{sec_id}"
                    if section_key not in old_section_columns:
                        continue
                    for rec in section.getElementsByTagName("record"):
                        record = {col: "-" for col in old_section_columns[section_key]}
                        record["NU_PTL"] = nu_ptl
                        record_id = rec.getAttribute("seq").strip() if rec.hasAttribute("seq") else "-"
                        record["Record_ID"] = record_id
                        for data in rec.getElementsByTagName("data"):
                            name = data.getAttribute("name").strip()
                            caption = data.getAttribute("caption").strip()
                            possible_keys = []
                            if caption:
                                possible_keys.append(caption)
                            if name:
                                possible_keys.append(name)
                            matched_field = None
                            for key in possible_keys:
                                for expected in old_section_columns[section_key]:
                                    if expected.upper() == key.upper():
                                        matched_field = expected
                                        break
                                if matched_field:
                                    break
                            if matched_field:
                                record[matched_field] = get_node_text(data)
                        old_sheets_data[section_key].append(record)

                # --- Trade Reference Records ---
                for tr_report in root.getElementsByTagName("tr_report"):
                    if tr_report.hasAttribute("type") and tr_report.getAttribute("type").strip().upper() == "TR":
                        header_info = {}
                        for header in tr_report.getElementsByTagName("header"):
                            for node in header.childNodes:
                                if node.nodeType == node.ELEMENT_NODE:
                                    tag = node.tagName.strip().lower()
                                    header_info[tag] = get_node_text(node)
                        enquiries = tr_report.getElementsByTagName("enquiry")
                        row_id_counter = 1
                        if enquiries:
                            for enq in enquiries:
                                trade_record = {col: "-" for col in old_section_columns["Trade Reference"]}
                                trade_record["NU_PTL"] = nu_ptl
                                trade_record["Row ID"] = str(row_id_counter)
                                row_id_counter += 1
                                trade_record["Date"] = header_info.get("date", "-")
                                trade_record["Req Name"] = header_info.get("req_name", "-")
                                trade_record["Req Com Name"] = header_info.get("req_com_name", "-")
                                trade_record["Req Com Addr"] = header_info.get("req_com_addr", "-")
                                trade_record["Ref Com Name"] = header_info.get("ref_com_name", "-")
                                trade_record["Ref Com Bus"] = header_info.get("ref_com_bus", "-")
                                trade_record["Report No"] = header_info.get("report_no", "-")
                                trade_record["IC LCNO"] = header_info.get("ic_lcno", "-")
                                trade_record["NIC BRNO"] = header_info.get("nic_brno", "-")
                                trade_record["Name"] = header_info.get("name", "-")
                                if enq.hasAttribute("account_no"):
                                    trade_record["Enquiry Account No"] = enq.getAttribute("account_no").strip() or "-"

                                # --- Relationship Section ---
                                relationship_section = None
                                for section in enq.getElementsByTagName("section"):
                                    if section.getAttribute("id").strip().lower() == "relationship":
                                        relationship_section = section
                                        break
                                if relationship_section:
                                    for data in relationship_section.getElementsByTagName("data"):
                                        dname = data.getAttribute("name").strip().lower()
                                        text_val = get_node_text(data)
                                        if dname == "rel_type":
                                            trade_record["Rel Type"] = text_val
                                        elif dname == "rel_status":
                                            trade_record["Rel Status"] = text_val
                                        elif dname == "rel_syear":
                                            trade_record["Rel SYear"] = text_val
                                        elif dname == "rel_smonth":
                                            trade_record["Rel SMonth"] = text_val
                                        elif dname == "rel_sday":
                                            trade_record["Rel SDay"] = text_val

                                # --- Account Status Section ---
                                account_section = None
                                for section in enq.getElementsByTagName("section"):
                                    if section.getAttribute("id").strip().lower() == "account_status":
                                        account_section = section
                                        break
                                if account_section:
                                    for data in account_section.getElementsByTagName("data"):
                                        dname = data.getAttribute("name").strip().lower()
                                        text_val = get_node_text(data)
                                        if dname == "statement_date":
                                            trade_record["Statement Date"] = text_val
                                        elif dname == "account_rating":
                                            trade_record["Account Rating"] = text_val
                                        elif dname == "account_term":
                                            trade_record["Account Term"] = text_val
                                        elif dname == "account_limit":
                                            trade_record["Account Limit"] = text_val
                                        elif dname == "account_status":
                                            trade_record["Account Status"] = text_val
                                        elif dname == "debtor_name":
                                            trade_record["Debtor Name"] = text_val
                                        elif dname == "debtor_ic_lcno":
                                            trade_record["Debtor IC LCNO"] = text_val
                                        elif dname == "debtor_nic_brno":
                                            trade_record["Debtor NIC BRNO"] = text_val
                                        elif dname == "address":
                                            trade_record["Address"] = text_val
                                        elif dname == "debt_type":
                                            trade_record["Debt Type"] = text_val
                                    # --- Robust Age Handling ---
                                    age_fields = ["30", "60", "90", "120", "150", "180", "210"]
                                    age_values = {af: "-" for af in age_fields}

                                    # Find the <data> node whose name attribute is "age"
                                    age_data = None
                                    for data in account_section.getElementsByTagName("data"):
                                        if data.getAttribute("name").strip().lower() == "age":
                                            age_data = data
                                            break

                                    if age_data:
                                        for age_item in age_data.childNodes:
                                            if age_item.nodeType == age_item.ELEMENT_NODE and age_item.tagName.lower() == "item":
                                                age_tag = age_item.getAttribute("name").strip()
                                                val = get_node_text(age_item)
                                                if age_tag in ["30", "60", "90", "120", "150", "180", "210"]:
                                                    age_values[f"Age {age_tag}"] = val
                                    for af in age_fields:
                                        trade_record[af] = age_values[af]

                                # --- Contact Section ---
                                contact_section = None
                                for section in enq.getElementsByTagName("section"):
                                    if section.getAttribute("id").strip().lower() == "contact":
                                        contact_section = section
                                        break
                                if contact_section:
                                    for data in contact_section.getElementsByTagName("data"):
                                        dname = data.getAttribute("name").strip().lower()
                                        text_val = get_node_text(data)
                                        if dname == "reference":
                                            trade_record["Contact Reference"] = text_val
                                        elif dname == "name":
                                            trade_record["Contact Name"] = text_val
                                        elif dname == "address":
                                            trade_record["Contact Address"] = text_val
                                        elif dname == "tel_no":
                                            trade_record["Contact Tel No"] = text_val
                                        elif dname == "fax_no":
                                            trade_record["Contact Fax No"] = text_val
                                        elif dname == "email":
                                            trade_record["Contact Email"] = text_val
                                        elif dname == "type":
                                            trade_record["Contact Type"] = text_val
                                        elif dname == "type_code":
                                            trade_record["Contact Type Code"] = text_val
                                trade_reference_data.append(trade_record)
                        else:
                            # If no enquiry is found, create a single record from header info
                            trade_record = {col: "-" for col in old_section_columns["Trade Reference"]}
                            trade_record["NU_PTL"] = nu_ptl
                            trade_record["Row ID"] = "1"
                            trade_record["Date"] = header_info.get("date", "-")
                            trade_record["Req Name"] = header_info.get("req_name", "-")
                            trade_record["Req Com Name"] = header_info.get("req_com_name", "-")
                            trade_record["Req Com Addr"] = header_info.get("req_com_addr", "-")
                            trade_record["Ref Com Name"] = header_info.get("ref_com_name", "-")
                            trade_record["Ref Com Bus"] = header_info.get("ref_com_bus", "-")
                            trade_record["Report No"] = header_info.get("report_no", "-")
                            trade_record["IC LCNO"] = header_info.get("ic_lcno", "-")
                            trade_record["NIC BRNO"] = header_info.get("nic_brno", "-")
                            trade_record["Name"] = header_info.get("name", "-")
                            for af in ["Age 30", "Age 60", "Age 90", "Age 120", "Age 150", "Age 180", "Age 210"]:
                                trade_record[af] = "-"
                            trade_reference_data.append(trade_record)

            except Exception as e:
                msg = f"Error parsing XML for NU_PTL {nu_ptl}: {str(e)}"
                self.after(0, self.append_error, msg)
                continue

            if index % 10 == 0 or index + 1 == total:
                progress = (index + 1) / total
                self.after(0, self.update_progress, progress, index + 1, total)

        # --- Export to Excel ---
        self.after(0, self.update_status, "Writing to Excel...")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        old_save_path = os.path.join(downloads_folder, f"old_ctos_report_{timestamp}.xlsx")

        with pd.ExcelWriter(old_save_path, engine="openpyxl") as writer:
            for sheet_name, records in old_sheets_data.items():
                if records:
                    df = pd.DataFrame(records)
                    df = df.reindex(columns=old_section_columns[sheet_name])
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            if trade_reference_data:
                df_tr = pd.DataFrame(trade_reference_data)
                df_tr = df_tr.reindex(columns=old_section_columns["Trade Reference"])
                df_tr.to_excel(writer, sheet_name="Trade Reference", index=False)

        self.after(0, self.update_status, "Export successful!")
        self.after(0, self.destroy_popup)

    def append_error(self, msg):
        if hasattr(self, "error_textbox"):
            self.error_textbox.configure(state="normal")
            self.error_textbox.insert("end", msg + "\n")
            self.error_textbox.see("end")
            self.error_textbox.configure(state="disabled")

    def update_status(self, text):
        self.status_label.configure(text=text)
        
    def destroy_popup(self):
        self.is_converting = False
        self.popup.destroy()


    def show_progress_popup(self):
        self.popup = ctk.CTkToplevel(self)
        self.popup.title("Export Progress")
        self.popup.geometry("600x250")
        self.popup.grab_set()  # Makes the popup modal
        self.popup.attributes("-topmost", True)

        self.progress_bar = ctk.CTkProgressBar(master=self.popup)
        self.progress_bar.pack(pady=10, fill='x', padx=20)
        self.progress_bar.set(0)
 

        self.status_label = ctk.CTkLabel(master=self.popup, text="Starting export...")
        self.status_label.pack(pady=5)

        self.error_textbox = ctk.CTkTextbox(master=self.popup, height=120, width=560)
        self.error_textbox.pack(pady=5, padx=10)


class XMLFormatView(ctk.CTkFrame):
    def __init__(self, parent, app):  # Add 'app' as a second argument
        super().__init__(parent)
        self.app = app  # Reference to the main app to access shared data
        self.xml_data = {}
        self.account_var = tk.StringVar()
        self.all_accounts = []

        # --- Header Frame ---
        header_frame = ctk.CTkFrame(self)
        header_frame.pack(fill="x", pady=10)
        

        # CTOS logo in center
        try:
            ctos_img = Image.open("Picture/ctos.png")
            self.ctos_logo = ctk.CTkImage(light_image=ctos_img, size=(220, 50))
            ctos_logo_label = ctk.CTkLabel(header_frame, image=self.ctos_logo, text="")
            ctos_logo_label.pack(side="top", pady=5)
        except Exception as e:
            ctos_logo_label = ctk.CTkLabel(header_frame, text="CTOS")
            ctos_logo_label.pack(side="top", pady=5)


        # Al Rajhi logo on right
        try:
            alrajhi_img = Image.open("Picture/alrajhi_logo.png")
            self.alrajhi_logo = ctk.CTkImage(light_image=alrajhi_img, size=(220, 50))
            alrajhi_logo_label = ctk.CTkLabel(header_frame, image=self.alrajhi_logo, text="")
            alrajhi_logo_label.place(relx=1.0, rely=0.0, anchor="ne")
        except Exception as e:
            alrajhi_logo_label = ctk.CTkLabel(header_frame, text="Al Rajhi")
            alrajhi_logo_label.place(relx=1.0, rely=0.0, anchor="ne")

        # --- Control Frame (Import + Combobox + Navigation) ---
        control_frame = ctk.CTkFrame(self)
        control_frame.pack(fill="x", pady=5)

        # Configure 3 columns to center the widgets
        control_frame.grid_columnconfigure(0, weight=1)  # Left spacer
        control_frame.grid_columnconfigure(1, weight=0)  # Buttons and Combobox
        control_frame.grid_columnconfigure(2, weight=1)  # Right spacer

        # Load arrow icons
        left_arrow_icon = ctk.CTkImage(Image.open("Picture/left-arrow.png"), size=(24, 24))
        right_arrow_icon = ctk.CTkImage(Image.open("Picture/right-arrow.png"), size=(24, 24))

        self.prev_btn = ctk.CTkButton(
            control_frame,
            text="",
            image=left_arrow_icon,
            fg_color="transparent",
            hover_color="#444",
            command=self.go_to_previous
        )
        self.prev_btn.grid(row=0, column=0, padx=10, pady=5, sticky="e")

        # ttk Combobox
        self.ttk_style = ttk.Style()
        self.ttk_style.theme_use('clam')
        self.account_combobox = ttk.Combobox(
            control_frame, textvariable=self.account_var, values=[], width=25
        )
        self.account_combobox.grid(row=0, column=1, padx=10, pady=5)
        self.account_combobox.bind("<<ComboboxSelected>>", self.display_xml_data)
        self.account_combobox.bind("<KeyRelease>", self.on_account_typing)

        self.next_btn = ctk.CTkButton(
            control_frame,
            text="",
            image=right_arrow_icon,
            fg_color="transparent",
            hover_color="#444",
            command=self.go_to_next
        )
        self.next_btn.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # --- XML Display ---
        self.xml_display = ctk.CTkTextbox(self, width=600, height=300)
        self.xml_display.pack(pady=10, fill="both", expand=True)

        # Add right-click context menu for copying
        self.xml_display.bind("<Button-3>", self.show_context_menu)
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Copy Selection", command=self.copy_selection)
        self.context_menu.add_command(label="Copy All", command=self.copy_all)

        # Refresh Button
        self.refresh_button = ctk.CTkButton(self, text="Refresh Data", command=self.refresh_data)
        self.refresh_button.pack(pady=10)
        
    
    def on_account_typing(self, event):
        typed = self.account_var.get().lower()
        
        # Filter values that contain the typed substring
        filtered = [acct for acct in self.all_accounts if typed in acct.lower()]
        
        # Update combobox values dynamically
        self.account_combobox['values'] = filtered

            
    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def copy_selection(self):
        try:
            selected = self.xml_display.get("sel.first", "sel.last")
        except tk.TclError:
            selected = ""
        if selected:
            self.clipboard_clear()
            self.clipboard_append(selected)

    def copy_all(self):
        all_text = self.xml_display.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(all_text)

    def refresh_data(self):
        # Get the shared data from the main app
        data = self.app.shared_data
        if data is not None:
            self.process_data(data)

    def process_data(self, data):
        from collections import defaultdict
        import re

        self.xml_data = {}
        self.all_accounts = []

        # Group by base NU_PTL (remove _0, _1, _2, etc.)
        nuptl_to_xmls = defaultdict(list)
        for nu_ptl, group in data.groupby("NU_PTL"):
            # Extract base NU_PTL (before underscore)
            base_nuptl = str(nu_ptl).split("_")[0]
            group = group.sort_values("ROW_ID").reset_index(drop=True)
            rowid_counters = defaultdict(int)
            set_indices = []
            for idx, row in group.iterrows():
                rid = row["ROW_ID"]
                set_indices.append(rowid_counters[rid])
                rowid_counters[rid] += 1
            sets = defaultdict(list)
            for idx, row in enumerate(group.itertuples()):
                set_idx = set_indices[idx]
                sets[set_idx].append(str(row.XML))
            for set_idx, xmls in sets.items():
                # Apply extract_first_report to the combined XML before storing
                combined_xml = "".join(xmls)
                nuptl_to_xmls[base_nuptl].append(extract_first_report(combined_xml))

        def is_perfect_xml(xml):
            xml = xml.strip()
            import re
            # Remove XML declaration if present
            if xml.startswith("<?xml"):
                xml = re.sub(r"<\?xml[^>]*\?>", "", xml).strip()
            # Exclude if starts with <root>, <span>, or <div>
            for tag in ("<root", "<span", "<div"):
                if xml.startswith(tag):
                    return False
            # Must contain <enq_report or <report (not just wrapper)
            if "<enq_report" in xml or "<report" in xml:
                return True
            return False

        for base_nuptl, xml_list in nuptl_to_xmls.items():
            # Collect all perfect XMLs
            perfect_xmls = [xml for xml in xml_list if is_perfect_xml(xml)]
            if perfect_xmls:
                perfect = perfect_xmls[0]
            else:
                perfect = xml_list[0]
            self.xml_data[base_nuptl] = perfect
            self.all_accounts.append(base_nuptl)

        self.account_combobox['values'] = self.all_accounts
        if self.all_accounts:
            self.account_var.set(self.all_accounts[0])
            self.display_xml_data()

    def display_xml_data(self, event=None):
        selected_account = self.account_var.get()
        if selected_account in self.xml_data:
            self.current_index = self.all_accounts.index(selected_account)
            raw_xml = self.xml_data[selected_account]

            # Extract only the content up to the first </report> tag before processing
            raw_xml = extract_first_report(raw_xml)

            try:
                cleaned_xml = clean_malformed_xml(raw_xml)
                dom = xml.dom.minidom.parseString(cleaned_xml)
                pretty_xml = dom.toprettyxml(indent="    ")
            except Exception as e:
                print(f"Error cleaning XML: {e}")
                pretty_xml = "<Error: Invalid or malformed XML>"

            self.xml_display.delete("1.0", "end")
            self.xml_display.insert("1.0", pretty_xml)
        else:
            self.xml_display.delete("1.0", "end")
            self.xml_display.insert("1.0", "No data available for the selected account.")

    def go_to_next(self):
        if self.all_accounts and self.current_index < len(self.all_accounts) - 1:
            self.current_index += 1
            self.account_var.set(self.all_accounts[self.current_index])
            self.display_xml_data()

    def go_to_previous(self):
        if self.all_accounts and self.current_index > 0:
            self.current_index -= 1
            self.account_var.set(self.all_accounts[self.current_index])
            self.display_xml_data()
                
    def update_navigation_buttons(self):
        if self.filtered_data is None or self.filtered_data.empty:
            self.prev_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
            return

        self.prev_button.configure(state="normal" if self.current_index > 0 else "disabled")
        self.next_button.configure(state="normal" if self.current_index < len(self.filtered_data) - 1 else "disabled")


class CTOSSummaryView(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.search_var_new = tk.StringVar()
        self.search_var_old = tk.StringVar()
        self.new_columns = [
            "Section A", "Section B1", "Section B2", "Section C", "Section D1",
            "Section D2", "Section D3", "Section D4", "Section E1", "Section E2", "DD_INDEX"
        ]
        self.old_columns = [
            "Section A", "Section B", "Section C", "Section D", "Section E", "Trade Reference"
        ]
        self.create_main_layout()

    def create_main_layout(self):

        # --- Tabview for New/Old CTOS ---
        self.tabview = CTkTabview(self, width=1600, height=700)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        self.new_tab = self.tabview.add("New CTOS Summary")
        self.old_tab = self.tabview.add("Old CTOS Summary")

        # --- New CTOS Summary Tab ---
        self._create_new_ctos_tab(self.new_tab)

        # --- Old CTOS Summary Tab ---
        self._create_old_ctos_tab(self.old_tab)

    # ----------- NEW CTOS TAB -----------
    def _create_new_ctos_tab(self, parent):
        # --- Header Frame with click counter ---
        header_frame = ctk.CTkFrame(parent)
        header_frame.pack(fill="x", pady=5)
        header_frame.bind("<Button-1>", self.on_header_click_new)

        # CTOS logo in center
        try:
            ctos_img = Image.open("Picture/ctos.png")
            self.ctos_logo_new = ctk.CTkImage(light_image=ctos_img, size=(220, 50))
            ctos_logo_label = ctk.CTkLabel(header_frame, image=self.ctos_logo_new, text="")
            ctos_logo_label.pack(side="top", pady=5)
            ctos_logo_label.bind("<Button-1>", self.on_header_click_new)
        except Exception:
            ctos_logo_label = ctk.CTkLabel(header_frame, text="NEW CTOS SUMMARY")
            ctos_logo_label.pack(side="top", pady=5)
            ctos_logo_label.bind("<Button-1>", self.on_header_click_new)

        # Header click counter display
        self.click_counter_label_new = ctk.CTkLabel(header_frame, text="Clicks: 0", font=ctk.CTkFont(size=12))
        self.click_counter_label_new.place(relx=0.0, rely=0.0, anchor="nw")

        search_frame = ctk.CTkFrame(parent)
        search_frame.pack(pady=5)
        ctk.CTkLabel(search_frame, text="Search NU_PTL:").pack(side="left", padx=(5, 2))
        search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var_new, width=200)
        search_entry.pack(side="left", padx=(0, 5))
        search_entry.bind("<Return>", self.search_summary_new)
        search_btn = ctk.CTkButton(search_frame, text="Search", command=self.search_summary_new)
        search_btn.pack(side="left")

        refresh_button = ctk.CTkButton(
            parent, text="Refresh Summary", command=self.refresh_summary_new, width=150, height=30
        )
        refresh_button.pack(pady=5)
        self.progress_bar_new = ctk.CTkProgressBar(parent, mode="determinate")
        self.progress_bar_new.pack(fill="x", padx=10, pady=(5,10))
        self.progress_bar_new.set(0)
        self.table_frame_new = ctk.CTkFrame(parent)
        self.table_frame_new.pack(fill="both", expand=True, padx=10, pady=10)
        self.create_summary_table_new({})

    def refresh_summary_new(self):
        data = self.app.shared_data
        if data is None or data.empty:
            messagebox.showerror("Error", "No shared data available for summary!")
            return

        self.progress_bar_new.set(0)
        def background_task():
            summary = self.calculate_new_ctos_summary(data)
            self.summary_data_new = summary  # Store for searching
            self.after(0, lambda: self.create_summary_table_new(summary))
        threading.Thread(target=background_task, daemon=True).start()
        
    def search_summary_new(self, event=None):
        search_value = self.search_var_new.get().strip().lower()
        if not hasattr(self, "summary_data_new") or not self.summary_data_new:
            return
        if not search_value:
            self.create_summary_table_new(self.summary_data_new)
            return
        filtered = {k: v for k, v in self.summary_data_new.items() if search_value in str(k).lower()}
        self.create_summary_table_new(filtered)

    def on_header_click_new(self, event):
        """Handle header clicks and update counter for new CTOS summary"""
        if not hasattr(self.app, 'header_click_count_new'):
            self.app.header_click_count_new = 0
        self.app.header_click_count_new += 1
        self.click_counter_label_new.configure(text=f"Clicks: {self.app.header_click_count_new}")

    def calculate_new_ctos_summary(self, records):
        import xml.dom.minidom

        summary = {}
        groups = list(records.groupby("NU_PTL"))
        total_groups = len(groups)
        count = 0

        for nu_ptl, group in groups:
            xml_fragments = group["XML"].dropna().astype(str).tolist()
            # Apply extract_first_report to each fragment before combining
            cleaned_fragments = [extract_first_report(fragment) for fragment in xml_fragments]
            combined_xml = "<root>" + "".join(cleaned_fragments) + "</root>"

            # --- Only process if it's new CTOS ---
            if not any(tag in combined_xml for tag in ["<section_a", "<section_d2", "<section_d4", "<section_etr_plus", "<dd_index"]):
                continue

            row = {
                "Section A": 0, "Section B1": 0, "Section B2": 0, "Section C": "-",
                "Section D1": 0, "Section D2": 0, "Section D3": "-", "Section D4": 0,
                "Section E1": 0, "Section E2": 0, "DD_INDEX": "-"
            }
            try:
                dom = xml.dom.minidom.parseString(combined_xml)
                # Section A
                row["Section A"] = sum(len(sec.getElementsByTagName("record")) for sec in dom.getElementsByTagName("section_a"))
                # Section B1 = section_c
                row["Section B1"] = sum(len(sec.getElementsByTagName("record")) for sec in dom.getElementsByTagName("section_c"))
                # Section B2 = section_b, only <record rpttype="Ib"> after <history>
                b2_count = 0
                for section_b in dom.getElementsByTagName("section_b"):
                    found_history = False
                    for node in section_b.childNodes:
                        if node.nodeType == node.ELEMENT_NODE and node.tagName == "history":
                            found_history = True
                        if found_history and node.nodeType == node.ELEMENT_NODE and node.tagName == "record":
                            if node.getAttribute("rpttype") == "Ib":
                                b2_count += 1
                row["Section B2"] = b2_count
                # Section D1 = section_d
                row["Section D1"] = sum(len(sec.getElementsByTagName("record")) for sec in dom.getElementsByTagName("section_d"))
                # Section D2 = section_d2
                row["Section D2"] = sum(len(sec.getElementsByTagName("record")) for sec in dom.getElementsByTagName("section_d2"))
                # Section D4 = section_d4
                row["Section D4"] = sum(len(sec.getElementsByTagName("record")) for sec in dom.getElementsByTagName("section_d4"))
                # Section E1 = section_etr_plus
                row["Section E1"] = sum(len(sec.getElementsByTagName("record")) for sec in dom.getElementsByTagName("section_etr_plus"))
                # Section E2 = section_e (count <enquiry>)
                row["Section E2"] = sum(len(sec.getElementsByTagName("enquiry")) for sec in dom.getElementsByTagName("section_e"))
                # DD_INDEX
                dd_index_nodes = dom.getElementsByTagName("dd_index")
                if dd_index_nodes and dd_index_nodes[0].firstChild and dd_index_nodes[0].firstChild.nodeValue.strip():
                    row["DD_INDEX"] = dd_index_nodes[0].firstChild.nodeValue.strip()
            except Exception:
                pass
            summary[nu_ptl] = row
            count += 1
            self.after(0, self.progress_bar_new.set, count / total_groups)
        return summary

    def create_summary_table_new(self, summary):
        for widget in self.table_frame_new.winfo_children():
            widget.destroy()
        container = ctk.CTkFrame(self.table_frame_new)
        container.pack(fill="both", expand=True)
        all_columns = ["NU_PTL"] + self.new_columns
        tree = ttk.Treeview(container, columns=all_columns, show="headings")
        tree.heading("NU_PTL", text="NU_PTL")
        tree.column("NU_PTL", width=120, anchor="center")
        for col in self.new_columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="center")
        nu_ptl_list = sorted(summary.keys())
        for nu_ptl in nu_ptl_list:
            counts = summary[nu_ptl]
            row_values = [str(nu_ptl)] + [str(counts.get(col, "")) for col in self.new_columns]
            tree.insert("", "end", values=row_values)
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        self.summary_tree_new = tree
        
        # Add right-click context menu for NU_PTL
        tree.bind("<Button-3>", self.show_nuptl_context_menu_new)
        self.nuptl_context_menu_new = tk.Menu(self, tearoff=0)
        self.nuptl_context_menu_new.add_command(label="View XML", command=lambda: self.navigate_to_view("xml"))
        self.nuptl_context_menu_new.add_command(label="View Report", command=lambda: self.navigate_to_view("report"))
        self.selected_nuptl_new = None
        
        convert_btn = ctk.CTkButton(self.table_frame_new, text="Convert", command=self.convert_summary_to_excel_new, width=150, height=30)
        convert_btn.pack(pady=10)

    def convert_summary_to_excel_new(self):
        rows = []
        for child in self.summary_tree_new.get_children():
            rows.append(self.summary_tree_new.item(child)["values"])
        import pandas as pd
        df = pd.DataFrame(rows, columns=["NU_PTL"] + self.new_columns)
        from tkinter import filedialog
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel Files", "*.xlsx")],
                                                title="Save New CTOS Summary to Excel")
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Export", "New CTOS summary exported successfully!")

    # ----------- OLD CTOS TAB -----------
    def _create_old_ctos_tab(self, parent):
        # --- Header Frame with click counter ---
        header_frame = ctk.CTkFrame(parent)
        header_frame.pack(fill="x", pady=5)
        header_frame.bind("<Button-1>", self.on_header_click_old)

        # CTOS logo in center
        try:
            ctos_img = Image.open("Picture/ctos.png")
            self.ctos_logo_old = ctk.CTkImage(light_image=ctos_img, size=(220, 50))
            ctos_logo_label = ctk.CTkLabel(header_frame, image=self.ctos_logo_old, text="")
            ctos_logo_label.pack(side="top", pady=5)
            ctos_logo_label.bind("<Button-1>", self.on_header_click_old)
        except Exception:
            ctos_logo_label = ctk.CTkLabel(header_frame, text="OLD CTOS SUMMARY")
            ctos_logo_label.pack(side="top", pady=5)
            ctos_logo_label.bind("<Button-1>", self.on_header_click_old)

        # Header click counter display
        self.click_counter_label_old = ctk.CTkLabel(header_frame, text="Clicks: 0", font=ctk.CTkFont(size=12))
        self.click_counter_label_old.place(relx=0.0, rely=0.0, anchor="nw")

        search_frame = ctk.CTkFrame(parent)
        search_frame.pack(pady=5)
        ctk.CTkLabel(search_frame, text="Search NU_PTL:").pack(side="left", padx=(5, 2))
        search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var_old, width=200)
        search_entry.pack(side="left", padx=(0, 5))
        search_entry.bind("<Return>", self.search_summary_old)
        search_btn = ctk.CTkButton(search_frame, text="Search", command=self.search_summary_old)
        search_btn.pack(side="left")

        refresh_button = ctk.CTkButton(
            parent, text="Refresh Summary", command=self.refresh_summary_old, width=150, height=30
        )
        refresh_button.pack(pady=5)
        self.progress_bar_old = ctk.CTkProgressBar(parent, mode="determinate")
        self.progress_bar_old.pack(fill="x", padx=10, pady=(5,10))
        self.progress_bar_old.set(0)
        self.table_frame_old = ctk.CTkFrame(parent)
        self.table_frame_old.pack(fill="both", expand=True, padx=10, pady=10)
        self.create_summary_table_old({})

    def refresh_summary_old(self):
        data = self.app.shared_data
        if data is None or data.empty:
            messagebox.showerror("Error", "No shared data available for summary!")
            return

        self.progress_bar_old.set(0)
        def background_task():
            summary = self.calculate_old_ctos_summary(data)
            self.summary_data_old = summary  # Store for searching
            self.after(0, lambda: self.create_summary_table_old(summary))
        threading.Thread(target=background_task, daemon=True).start()
        
    def search_summary_old(self, event=None):
        search_value = self.search_var_old.get().strip().lower()
        if not hasattr(self, "summary_data_old") or not self.summary_data_old:
            return
        if not search_value:
            self.create_summary_table_old(self.summary_data_old)
            return
        filtered = {k: v for k, v in self.summary_data_old.items() if search_value in str(k).lower()}
        self.create_summary_table_old(filtered)

    def on_header_click_old(self, event):
        """Handle header clicks and update counter for old CTOS summary"""
        if not hasattr(self.app, 'header_click_count_old'):
            self.app.header_click_count_old = 0
        self.app.header_click_count_old += 1
        self.click_counter_label_old.configure(text=f"Clicks: {self.app.header_click_count_old}")

    def calculate_old_ctos_summary(self, records):
        import xml.dom.minidom

        summary = {}
        groups = list(records.groupby("NU_PTL"))
        total_groups = len(groups)
        count = 0

        for nu_ptl, group in groups:
            xml_fragments = group["XML"].dropna().astype(str).tolist()
            # Apply extract_first_report to each fragment before combining
            cleaned_fragments = [extract_first_report(fragment) for fragment in xml_fragments]
            combined_xml = "<root>" + "".join(cleaned_fragments) + "</root>"

            # --- Only process if it's old CTOS ---
            # Old CTOS: no <section_a>, <section_d2>, <section_d4>, <section_etr_plus>, <dd_index>
            if any(tag in combined_xml for tag in ["<section_a", "<section_d2", "<section_d4", "<section_etr_plus", "<dd_index"]):
                continue

            row = {
                "Section A": 0, "Section B": 0, "Section C": 0, "Section D": 0, "Section E": 0, "Trade Reference": 0
            }
            try:
                dom = xml.dom.minidom.parseString(combined_xml)
                # Section A-E
                for sec_id, col in zip(["A", "B", "C", "D", "E"], ["Section A", "Section B", "Section C", "Section D", "Section E"]):
                    row[col] = sum(len(sec.getElementsByTagName("record")) for sec in dom.getElementsByTagName("section") if sec.getAttribute("id").strip().upper() == sec_id)
                # Trade Reference
                row["Trade Reference"] = len(dom.getElementsByTagName("tr_report"))
            except Exception:
                pass
            summary[nu_ptl] = row
            count += 1
            self.after(0, self.progress_bar_old.set, count / total_groups)
        return summary

    def create_summary_table_old(self, summary):
        for widget in self.table_frame_old.winfo_children():
            widget.destroy()
        container = ctk.CTkFrame(self.table_frame_old)
        container.pack(fill="both", expand=True)
        all_columns = ["NU_PTL"] + self.old_columns
        tree = ttk.Treeview(container, columns=all_columns, show="headings")
        tree.heading("NU_PTL", text="NU_PTL")
        tree.column("NU_PTL", width=120, anchor="center")
        for col in self.old_columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="center")
        nu_ptl_list = sorted(summary.keys())
        for nu_ptl in nu_ptl_list:
            counts = summary[nu_ptl]
            row_values = [str(nu_ptl)] + [str(counts.get(col, "")) for col in self.old_columns]
            tree.insert("", "end", values=row_values)
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        self.summary_tree_old = tree
        
        # Add right-click context menu for NU_PTL
        tree.bind("<Button-3>", self.show_nuptl_context_menu_old)
        self.nuptl_context_menu_old = tk.Menu(self, tearoff=0)
        self.nuptl_context_menu_old.add_command(label="View XML", command=lambda: self.navigate_to_view("xml"))
        self.nuptl_context_menu_old.add_command(label="View Report", command=lambda: self.navigate_to_view("report"))
        self.selected_nuptl_old = None
        
        convert_btn = ctk.CTkButton(self.table_frame_old, text="Convert", command=self.convert_summary_to_excel_old, width=150, height=30)
        convert_btn.pack(pady=10)

    def convert_summary_to_excel_old(self):
        rows = []
        for child in self.summary_tree_old.get_children():
            rows.append(self.summary_tree_old.item(child)["values"])
        import pandas as pd
        df = pd.DataFrame(rows, columns=["NU_PTL"] + self.old_columns)
        from tkinter import filedialog
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel Files", "*.xlsx")],
                                                title="Save Old CTOS Summary to Excel")
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Export", "Old CTOS summary exported successfully!")

    def show_nuptl_context_menu_new(self, event):
        """Show context menu for NU_PTL in new CTOS summary"""
        try:
            item = self.summary_tree_new.selection()[0] if self.summary_tree_new.selection() else None
            if not item:
                item = self.summary_tree_new.identify_row(event.y)
            if item:
                values = self.summary_tree_new.item(item, "values")
                if values and len(values) > 0:
                    self.selected_nuptl_new = str(values[0])  # First column is NU_PTL
                    self.nuptl_context_menu_new.tk_popup(event.x_root, event.y_root)
        except Exception:
            pass
        finally:
            try:
                self.nuptl_context_menu_new.grab_release()
            except:
                pass

    def show_nuptl_context_menu_old(self, event):
        """Show context menu for NU_PTL in old CTOS summary"""
        try:
            item = self.summary_tree_old.selection()[0] if self.summary_tree_old.selection() else None
            if not item:
                item = self.summary_tree_old.identify_row(event.y)
            if item:
                values = self.summary_tree_old.item(item, "values")
                if values and len(values) > 0:
                    self.selected_nuptl_old = str(values[0])  # First column is NU_PTL
                    self.nuptl_context_menu_old.tk_popup(event.x_root, event.y_root)
        except Exception:
            pass
        finally:
            try:
                self.nuptl_context_menu_old.grab_release()
            except:
                pass

    def navigate_to_view(self, view_type):
        """Navigate to XML or Report view with selected NU_PTL"""
        # Get the selected NU_PTL from either new or old summary
        nuptl = None
        if hasattr(self, 'selected_nuptl_new') and self.selected_nuptl_new:
            nuptl = self.selected_nuptl_new
        elif hasattr(self, 'selected_nuptl_old') and self.selected_nuptl_old:
            nuptl = self.selected_nuptl_old
        
        if not nuptl:
            messagebox.showwarning("Warning", "No NU_PTL selected")
            return
        
        if view_type == "xml":
            # Navigate to XML Format view
            self.app.show_xml_format()
            # Set the NU_PTL in the XML view
            if nuptl in self.app.xml_format_view.all_accounts:
                self.app.xml_format_view.account_var.set(nuptl)
                self.app.xml_format_view.display_xml_data()
            else:
                messagebox.showinfo("Info", f"NU_PTL {nuptl} not found in XML data")
        
        elif view_type == "report":
            # Navigate to CTOS Report view
            self.app.show_ctos_report()
            # Refresh data first if needed
            self.app.ctos_report_view.refresh_data()
            # Set the NU_PTL in the Report view
            if nuptl in self.app.ctos_report_view.all_accounts:
                self.app.ctos_report_view.account_var.set(nuptl)
                self.app.ctos_report_view.current_index = self.app.ctos_report_view.all_accounts.index(nuptl)
                self.app.ctos_report_view.display_data()
            else:
                messagebox.showinfo("Info", f"NU_PTL {nuptl} not found in report data")

if __name__ == "__main__":
    app = CTOSReportApp()
    app.mainloop()