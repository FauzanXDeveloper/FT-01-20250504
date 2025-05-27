import threading
import time
import subprocess
import sys
import os
import math
import xlsxwriter
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

def clean_malformed_xml(xml_str):
   
    import re

    if not isinstance(xml_str, str):
        return "<root></root>"

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
        self.iconbitmap("ctos.ico")  # Set the application icon
        self.title("CTOS Report Generator")
        self.geometry("1800x900")
        
        self.current_theme = "system" 
        
        customtkinter.set_default_color_theme("Themes/patina.json")
        
        # Shared data (Excel, parsed XML, etc.)
        self.shared_data = None

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
        hamburger_img = ctk.CTkImage(Image.open("hamburger.png"), size=(24, 24))
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

        # Sidebar buttons with improved spacing & bigger size
        button_font = ctk.CTkFont(family="Segoe UI", size=16, weight="bold")  # increased font size
        self.import_button = ctk.CTkButton(
            self.sidebar,
            text="📂 Import Excel",
            command=self.import_excel,
            font=button_font,
            width=150,      # increased width
            height=50       # increased height
        )
        self.import_button.pack(pady=20, padx=20)
        
        self.xml_format_button = ctk.CTkButton(
            self.sidebar,
            text="XML Format",
            command=self.show_xml_format,
            font=button_font,
            width=150,
            height=50
        )
        self.xml_format_button.pack(pady=20, padx=20)
        
        self.ctos_report_button = ctk.CTkButton(
            self.sidebar,
            text="CTOS Report",
            command=self.show_ctos_report,
            font=button_font,
            width=150,
            height=50
        )
        self.ctos_report_button.pack(pady=20, padx=20)
        
        self.ctos_summary_button = ctk.CTkButton(
            self.sidebar,
            text="CTOS Summary",
            command=self.show_ctos_summary,
            font=button_font,
            width=150,
            height=50
        )
        self.ctos_summary_button.pack(pady=20, padx=20)
        
        self.main_app_button = ctk.CTkButton(
            self.sidebar,
            text="🔙 Back to Main",
            command=self.show_main_app,
            font=button_font,
            width=150,
            height=50
        )
        self.main_app_button.pack(pady=20, padx=20)
        
        self.sidebar_buttons = [
            self.import_button,
            self.xml_format_button,
            self.ctos_report_button,
            self.ctos_summary_button,
            self.main_app_button
        ]
        
        
        try:
            self.dark_icon = ctk.CTkImage(Image.open("dark_mode_icon.png"), size=(24, 24))
            self.system_icon = ctk.CTkImage(Image.open("light_mode_icon.png"), size=(24, 24))
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
        
    
    def toggle_sidebar(self):
        if self.sidebar_expanded:
            for btn in self.sidebar_buttons:
                btn.pack_forget()
            # Also hide mode toggle? Optionally keep it visible.
            self.sidebar.configure(width=self.SIDEBAR_SHRUNK_WIDTH)
            self.sidebar_expanded = False
        else:
            for btn in self.sidebar_buttons:
                btn.pack(pady=15, padx=20)
            self.sidebar.configure(width=self.SIDEBAR_EXPANDED_WIDTH)
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
                cleaned_xml_dict = {}
                for _, row in df.iterrows():
                    nu_ptl = str(row["NU_PTL"])
                    raw_xml = row["XML"]
                    cleaned_xml = clean_malformed_xml(raw_xml)
                    cleaned_xml_dict[nu_ptl] = cleaned_xml
                def update_data():
                    self.shared_data = df
                    self.xml_format_view.xml_data = cleaned_xml_dict
                    self.destroy_progress_popup()
                    messagebox.showinfo("Success", "Excel imported and XML cleaned successfully!")
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
        script_path = os.path.join(os.getcwd(), "integrate.py")
        if proc is None:
            subprocess.Popen([sys.executable, script_path])
        else:
            bring_integrate_to_front()
        
        
class CTOSReportView(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app  # Reference to the main app to access shared data
        self.account_var = tk.StringVar()
        self.search_var = tk.StringVar()
        self.all_accounts = []
        self.current_index = 0  # Track the current NU_PTL index
        self.filtered_data = None  # Store filtered data for navigation
        
    
        # --- Header Frame ---
        header_frame = ctk.CTkFrame(self)
        header_frame.pack(fill="x", pady=10)

        # CTOS logo in center
        try:
            ctos_img = Image.open("ctos.png")
            self.ctos_logo = ctk.CTkImage(light_image=ctos_img, size=(220, 50))
            ctos_logo_label = ctk.CTkLabel(header_frame, image=self.ctos_logo, text="")
            ctos_logo_label.pack(side="top", pady=5)
        except Exception as e:
            ctos_logo_label = ctk.CTkLabel(header_frame, text="CTOS")
            ctos_logo_label.pack(side="top", pady=5)

        # Al Rajhi logo on right
        try:
            alrajhi_img = Image.open("alrajhi_logo.png")
            self.alrajhi_logo = ctk.CTkImage(light_image=alrajhi_img, size=(220, 50))
            alrajhi_logo_label = ctk.CTkLabel(header_frame, image=self.alrajhi_logo, text="")
            alrajhi_logo_label.place(relx=1.0, rely=0.0, anchor="ne")
        except Exception as e:
            alrajhi_logo_label = ctk.CTkLabel(header_frame, text="Al Rajhi")
            alrajhi_logo_label.place(relx=1.0, rely=0.0, anchor="ne")

        # --- Control Frame (Import + Combobox + Navigation) ---
        self.control_frame = ctk.CTkFrame(self)
        self.control_frame.pack(fill="x", pady=5)

        # Configure 3 columns to center the widgets
        self.control_frame.grid_columnconfigure(0, weight=1)  # Left spacer
        self.control_frame.grid_columnconfigure(1, weight=0)  # Buttons and Combobox
        self.control_frame.grid_columnconfigure(2, weight=1)  # Right spacer

        # Load arrow icons
        left_arrow_icon = ctk.CTkImage(Image.open("left-arrow.png"), size=(24, 24))
        right_arrow_icon = ctk.CTkImage(Image.open("right-arrow.png"), size=(24, 24))
        
        # Previous Button
        self.prev_btn = ctk.CTkButton(
            self.control_frame,
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
            self.control_frame, textvariable=self.account_var, values=[], width=25
        )
        self.account_combobox.grid(row=0, column=1, padx=10, pady=5)
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
        self.next_btn.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        self.export_icon = ctk.CTkImage(Image.open("export.png"), size=(24, 24))
        
        # Convert to Excel Button
        self.convert_button = ctk.CTkButton(self.control_frame, text="Convert to Excel", image=self.export_icon,command=self.convert_to_excel, )
        self.convert_button.grid(row=0, column=4, padx=5)
        

        # --- Treeview for displaying parsed XML data ---
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(self.tree_frame, show="headings")
        self.tree.pack(fill="both", expand=True, side="left", padx=5, pady=5)
       
        # Set up two columns with customized widths and centered text.
        self.tree["columns"] = ["Field", "Value"]
        self.tree.heading("Field", text="Field")
        self.tree.heading("Value", text="Value")
        self.tree.column("Field", anchor="center", width=300)
        self.tree.column("Value", anchor="center", width=400)

        # Add context menu for copying
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Copy Row", command=self.copy_row)
        self.context_menu.add_command(label="Copy Cell", command=self.copy_cell)
        self._right_click_row = None
        self._right_click_col = None

        # Add a scrollbar
        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        # Create a refresh button
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
        cleaned_data = {
            key: clean_malformed_xml(value)
            for key, value in xml_format_view.xml_data.items()
        }
        self.filtered_data = pd.DataFrame.from_dict(cleaned_data, orient="index", columns=["XML"])
        self.filtered_data.reset_index(inplace=True)
        self.filtered_data.rename(columns={"index": "NU_PTL"}, inplace=True)
        self.all_accounts = self.filtered_data["NU_PTL"].tolist()
        self.account_combobox['values'] = self.all_accounts
        if self.all_accounts:
            self.account_combobox.current(self.current_index)
        self.display_data()

    def show_context_menu(self, event):
        # Identify row and column under mouse
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
            # Example: handle <enq_report>
            if tag == "enq_report" and child.hasAttribute("id"):
                field = "Report ID"
                value = child.getAttribute("id")
                self.tree.insert("", "end", values=[field, value])
                self.parse_xml_to_treeview(child, field)
                continue
            if tag == "header":
                # Check if this header is broken (contains a nested <report> element)
                has_nested_report = any(r for r in child.getElementsByTagName("report"))
                if has_nested_report:
                    # Instead of skipping completely, recursively process its children
                    for sub in child.childNodes:
                        self.parse_xml_to_treeview(sub, parent_path)
                    continue
                else:
                    # Process a good header normally
                    for sub in child.childNodes:
                        if sub.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                            sub_tag = sub.tagName
                            value = sub.firstChild.nodeValue.strip() if (sub.firstChild and sub.firstChild.nodeValue) else "-"
                            self.tree.insert("", "end", values=[sub_tag, value])
                    continue
            # New summary logic
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
                self.tree.insert("", "end", values=[title, "-"])
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
                value = child.firstChild.nodeValue.strip() if (child.firstChild and child.firstChild.nodeValue) else ""
                self.tree.insert("", "end", values=[field, value])
                continue
            self.parse_xml_to_treeview(child, parent_path)

    def go_to_next(self):
        if self.all_accounts and self.current_index < len(self.all_accounts) - 1:
            self.current_index += 1
            self.account_var.set(self.all_accounts[self.current_index])
            self.display_data()  # <- FIXED


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

    def convert_to_excel(self):
        self.is_converting = True
        self.show_progress_popup()
        threading.Thread(target=self.convert_to_excel_thread, daemon=True).start()

    def update_progress(self, progress, index, total):
        self.progress_bar.set(progress)
        self.status_label.configure(text=f"Processing {index} of {total}")
        # Removed: self.popup.update()  # <-- Avoid explicit update call here
        
    def convert_to_excel_thread(self):
        from collections import defaultdict

        # Your fixed columns per sheet
        section_columns = {
            "Header&Summary": ["NU_PTL", "user", "company", "account", "tel", "fax", "enq_date", "enq_time", "enq_status", "IC_LCNO", "NIC_BRNO", "NAME", "ALIAS", "STAT", "REF"],
            "Section-A": ["NU_PTL", "Record_ID", "ICNO", "MATCH", "NEWIC", "MATCH1", "NAME", "MATCH2", "ADDR", "ADDR1", "REMARK"],
            "Section-B": ["NU_PTL", "Record_ID", "CODE", "NAME", "MATCH", "ALIAS", "IC_LCNO", "NIC_BRNO", "REF", "CONUM", "CONAME", "REMARK", "REMARK2", "REMARK3", "AMOUNT", "ENTRY"],
            "Section-C": ["NU_PTL", "Record_ID", "EXTRANAME", "EXTRALOCAL", "EXTRALOCALA", "OBJECT", "INCORPRATION", "COMPANY PAIDUP", "SEARCH DATE", "EXTRALASTDOC", "NAME", "IC_LCNO", "NIC_BRNO", "STATUS", "SHARES", "REAMARK", "APPOINTED", "RESIGNED", "ADDRESS"],
            "Section-D": ["NU_PTL", "Record_ID", "ZTITLE", "ZSPECIAL", "NAME", "MATCH", "ALIAS", "I/C NO", "NEW IC", "REMARK", "ADDRESS", "FIRM", "PLAINTIFF", "CASE NO", "ZCOURT", "ACTION DATE", "ZNTPAP", "HEARING DATE", "AMOUNT", "SOLCTR", "LAWADD1", "TEL", "LAWADD2", "REF", "LAWADD3", "PLAINTIFF CONTACT", "CEDCONADD1", "CEDCONADD2", "CEDCONADD3"],
            "Section-E": ["NU_PTL", "Record_ID", "REFEREE", "INCORPORATION DATE", "NATURE OF BUSINESS", "ADDRESS", "TR_URL"]
        }

        try:
            if self.filtered_data is None or self.filtered_data.empty:
                self.after(0, self.update_status, "No data to convert.")
                return

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Save Converted Excel"
            )
            if not save_path:
                self.after(0, self.update_status, "Save cancelled.")
                return

            self.after(0, self.update_status, "Extracting data...")
            self.after(0, lambda: self.error_textbox.delete("1.0", "end"))
            self.after(0, self.progress_bar.set, 0)

            # Initialize data container
            sheets_data = {
                "Header&Summary": [],
                "Section-A": [],
                "Section-B": [],
                "Section-C": [],
                "Section-D": [],
                "Section-E": []
            }

            total = len(self.filtered_data)

            for index, (_, row) in enumerate(self.filtered_data.iterrows()):
                if not self.is_converting:
                    break

                nu_ptl = row.get("NU_PTL", f"Row{index}")
                xml_data = clean_malformed_xml(row.get("XML", ""))

                if pd.isna(xml_data) or not str(xml_data).strip():
                    continue

                try:
                    dom = xml.dom.minidom.parseString(xml_data)
                    root = dom.documentElement

                    # Header&Summary
                    # Collect header info into a dict with all keys preset to ""
                    header_record = {col: "" for col in section_columns["Header&Summary"]}
                    header_record["NU_PTL"] = nu_ptl

                    for header in root.getElementsByTagName("header"):
                        for node in header.childNodes:
                            if node.nodeType == node.ELEMENT_NODE:
                                tag = node.tagName.strip()
                                if tag in section_columns["Header&Summary"]:
                                    header_record[tag] = node.firstChild.nodeValue.strip() if node.firstChild else "-"
                    # Collect summary/enq_sum fields (field_sum nodes)
                    for summary in root.getElementsByTagName("summary"):
                        for enq_sum in summary.getElementsByTagName("enq_sum"):
                            for fs in enq_sum.getElementsByTagName("field_sum"):
                                if fs.hasAttribute("name"):
                                    field = fs.getAttribute("name").strip()
                                    if field in section_columns["Header&Summary"]:
                                        value = fs.firstChild.nodeValue.strip() if fs.firstChild else "-"
                                        header_record[field] = value
                    sheets_data["Header&Summary"].append(header_record)

                    # Sections A-E
                    for section in root.getElementsByTagName("section"):
                        section_id = section.getAttribute("id").strip()
                        section_key = f"Section-{section_id}"

                        if section_key not in sheets_data:
                            continue

                        for rec in section.getElementsByTagName("record"):
                            record = {col: "" for col in section_columns[section_key]}
                            record["NU_PTL"] = nu_ptl
                            record_id = rec.getAttribute("seq").strip() if rec.hasAttribute("seq") else ""
                            record["Record_ID"] = record_id

                            for data in rec.getElementsByTagName("data"):
                                name = data.getAttribute("name").strip().upper()
                                caption = data.getAttribute("caption").strip().upper()

                                possible_keys = []
                                if caption:
                                    possible_keys.append(caption)
                                if name:
                                    possible_keys.append(name)

                                matched_field = None
                                for key in possible_keys:
                                    for expected in section_columns.get(section_key, []):
                                        if expected.upper() == key:
                                            matched_field = expected
                                            break
                                    if matched_field:
                                        break

                                if matched_field:
                                    value = data.firstChild.nodeValue.strip() if data.firstChild else ""
                                    record[matched_field] = value

                            sheets_data[section_key].append(record)

                except Exception as e:
                    msg = f"Error parsing XML for NU_PTL {nu_ptl}: {str(e)}"
                    self.after(0, self.append_error, msg)
                    continue

                if index % 10 == 0 or index + 1 == total:
                    progress = (index + 1) / total
                    self.after(0, self.update_progress, progress, index + 1, total)

            # Export to Excel
            self.after(0, self.update_status, "Writing to Excel...")

            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                for sheet_name, records in sheets_data.items():
                    if records:
                        df = pd.DataFrame(records)
                        # Ensure columns are ordered exactly as defined
                        df = df.reindex(columns=section_columns[sheet_name])
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

            self.after(0, self.update_status, "Export successful!")
            self.after(0, self.destroy_popup)

        except Exception as e:
            self.after(0, self.update_status, f"Error: {str(e)}")
            self.after(0, self.append_error, f"Fatal error: {str(e)}")
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
            ctos_img = Image.open("ctos.png")
            self.ctos_logo = ctk.CTkImage(light_image=ctos_img, size=(220, 50))
            ctos_logo_label = ctk.CTkLabel(header_frame, image=self.ctos_logo, text="")
            ctos_logo_label.pack(side="top", pady=5)
        except Exception as e:
            ctos_logo_label = ctk.CTkLabel(header_frame, text="CTOS")
            ctos_logo_label.pack(side="top", pady=5)
        
        # Al Rajhi logo on right
        try:
            alrajhi_img = Image.open("alrajhi_logo.png")
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
        left_arrow_icon = ctk.CTkImage(Image.open("left-arrow.png"), size=(24, 24))
        right_arrow_icon = ctk.CTkImage(Image.open("right-arrow.png"), size=(24, 24))
        
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
        self.xml_data = {}
        grouped = data.groupby("NU_PTL")
        for nu_ptl, group in grouped:
            group = group.sort_values("ROW_ID")
            combined_xml = "".join(str(x) for x in group["XML"].tolist())
            self.xml_data[str(nu_ptl)] = combined_xml
        
        # Store all NU_PTL values for reference
        self.all_accounts = list(self.xml_data.keys())

        self.all_accounts = list(self.xml_data.keys())
        self.account_combobox['values'] = self.all_accounts
        if self.all_accounts:
            self.account_var.set(self.all_accounts[0])
            self.display_xml_data()

    def display_xml_data(self, event=None):
        selected_account = self.account_var.get()
        if selected_account in self.xml_data:
            self.current_index = self.all_accounts.index(selected_account)
            raw_xml = self.xml_data[selected_account]

            try:
                cleaned_xml = clean_malformed_xml(raw_xml)
                # Beautify the XML
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
        self.headers = ["", "Total", "MIA>=4", ">=4%", "AKPK", "AKPK %", "Woff", "Woff %"]
        self.sections = ["A", "B", "C", "D", "E"]
        self.rows = [""] + self.sections  # Blank (Total), then A–E
        self.create_main_layout()

    def create_main_layout(self):
        self.header_label = ctk.CTkLabel(
            self, text="CTOS Summary", font=ctk.CTkFont(size=16, weight="bold")
        )
        self.header_label.pack(pady=(10, 5))
        self.control_frame = ctk.CTkFrame(self)
        self.control_frame.pack(fill="x", padx=10, pady=5)
        self.table_frame = ctk.CTkFrame(self)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.refresh_button = ctk.CTkButton(
            self.control_frame, text="Refresh Summary", command=self.refresh_summary, anchor="center", width=150, height=30
        )
        self.refresh_button.pack(side="left")
        self.create_summary_table({})  # Initial blank table

    def refresh_summary(self):
        records = self.app.shared_data  # DataFrame with NU_PTL and XML columns
        summary = self.calculate_summary(records)
        self.create_summary_table(summary)

    def calculate_summary(self, records):
        from collections import defaultdict
        import pandas as pd
        import xml.dom.minidom

        # Combine all XML fragments for each NU_PTL
        combined_xml_per_nuptl = {}
        grouped = records.groupby("NU_PTL")
        for nu_ptl, group in grouped:
            xml_fragments = [str(x) for x in group["XML"].tolist() if pd.notna(x) and str(x).strip()]
            combined_xml = "".join(xml_fragments)
            combined_xml_per_nuptl[nu_ptl] = combined_xml

        unique_nu_ptls = set(combined_xml_per_nuptl.keys())
        section_nu_ptls = {sec: set() for sec in self.sections}
        section_record_count = defaultdict(lambda: {"total": 0, "mia": 0, "akpk": 0, "woff": 0})

        for nu_ptl, xml_data in combined_xml_per_nuptl.items():
            if not xml_data.strip():
                continue
            # Ensure XML is wrapped in a root tag
            if not xml_data.strip().startswith("<root>"):
                xml_data = f"<root>{xml_data}</root>"
            try:
                dom = xml.dom.minidom.parseString(xml_data)
                for section in dom.getElementsByTagName("section"):
                    sec_id = section.getAttribute("id").strip().upper()
                    if sec_id in self.sections:
                        records_in_section = section.getElementsByTagName("record")
                        if records_in_section:
                            section_nu_ptls[sec_id].add(nu_ptl)
                        for rec in records_in_section:
                            section_record_count[sec_id]["total"] += 1
                            for data in rec.getElementsByTagName("data"):
                                name = data.getAttribute("name").strip().upper()
                                value = data.firstChild.nodeValue.strip() if data.firstChild else ""
                                if name == "MIA":
                                    try:
                                        if int(value) >= 4:
                                            section_record_count[sec_id]["mia"] += 1
                                    except Exception:
                                        pass
                                if name == "AKPK":
                                    try:
                                        if int(value) == 1:
                                            section_record_count[sec_id]["akpk"] += 1
                                    except Exception:
                                        pass
                                if name == "WOFF":
                                    try:
                                        if int(value) == 1:
                                            section_record_count[sec_id]["woff"] += 1
                                    except Exception:
                                        pass
            except Exception as e:
                # print(f"Error parsing XML for NU_PTL {nu_ptl}: {e}")
                continue

        summary = {}
        # Overall totals (the Total row)
        total_accounts = len(unique_nu_ptls)
        overall_mia = sum(sec["mia"] for sec in section_record_count.values())
        overall_akpk = sum(sec["akpk"] for sec in section_record_count.values())
        overall_woff = sum(sec["woff"] for sec in section_record_count.values())
        summary[""] = {
            "Total": total_accounts,
            "MIA>=4": overall_mia,
            ">=4%": "",  # Not applicable for total row
            "AKPK": overall_akpk,
            "AKPK %": "",
            "Woff": overall_woff,
            "Woff %": "",
        }
        for sec in self.sections:
            total_nu_ptl_in_section = len(section_nu_ptls[sec])  # NU_PTLs with at least 1 record in this section
            mia = section_record_count[sec]["mia"]
            akpk = section_record_count[sec]["akpk"]
            woff = section_record_count[sec]["woff"]
            total = section_record_count[sec]["total"]  # total records in this section
            summary[sec] = {
                "Total": total_nu_ptl_in_section,
                "MIA>=4": mia,
                ">=4%": f"{(mia/total*100):.1f}%" if total > 0 else "-",
                "AKPK": akpk,
                "AKPK %": f"{(akpk/total*100):.1f}%" if total > 0 else "-",
                "Woff": woff,
                "Woff %": f"{(woff/total*100):.1f}%" if total > 0 else "-",
            }
        return summary

    def create_summary_table(self, summary):
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        border_color = "#888888"  # Border color
        border_width = 1

        # Create header row
        for col, header in enumerate(self.headers):
            cell_frame = tk.Frame(
                self.table_frame,
                background=border_color,
                highlightthickness=0
            )
            cell_frame.grid(row=0, column=col, sticky="nsew", padx=0, pady=0)
            label = ctk.CTkLabel(
                cell_frame,
                text=header,
                font=ctk.CTkFont(weight="bold"),
                fg_color="#f5f5f5",
                corner_radius=0
            )
            label.pack(fill="both", expand=True, padx=border_width, pady=border_width)

        # Create rows for overall Total and each section
        for row_idx, row_label in enumerate(self.rows):
            for col_idx, header in enumerate(self.headers):
                if col_idx == 0:
                    text = "Total" if row_label == "" else row_label
                else:
                    text = summary.get(row_label, {}).get(header, "")
                cell_frame = tk.Frame(
                    self.table_frame,
                    background=border_color,
                    highlightthickness=0
                )
                cell_frame.grid(row=row_idx + 1, column=col_idx, sticky="nsew", padx=0, pady=0)
                label = ctk.CTkLabel(
                    cell_frame,
                    text=str(text),
                    fg_color="#fff",
                    corner_radius=0
                )
                label.pack(fill="both", expand=True, padx=border_width, pady=border_width)

        # Make columns expand equally
        for i in range(len(self.headers)):
            self.table_frame.grid_columnconfigure(i, weight=1)

if __name__ == "__main__":
    app = CTOSReportApp()
    app.mainloop()
