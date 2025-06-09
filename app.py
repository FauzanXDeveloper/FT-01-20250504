import threading
import time
import subprocess
import sys
import os
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
        self.prev_btn.grid(row=0, column=0, padx=10, pady=5, sticky="e")

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

        self.export_icon = ctk.CTkImage(Image.open("Picture/export.png"), size=(24, 24))
        self.convert_button = ctk.CTkButton(self.control_frame, text="Old Ctos", image=self.export_icon, command=self.convert_to_excel)
        self.convert_button.grid(row=0, column=4, padx=5)
        self.convert_new_button = ctk.CTkButton(
            self.control_frame,
            text="New CTOS",
            image=self.export_icon,
            command=self.convert_new_ctos_to_excel
        )
        self.convert_new_button.grid(row=0, column=5, padx=5)

        self.dd_index_var = tk.StringVar(value="-")
        self.dd_index_label = ctk.CTkLabel(
            self,
            textvariable=self.dd_index_var,
            width=48,
            height=48,
            fg_color="#1976d2",
            text_color="#fff",
            font=ctk.CTkFont(size=18, weight="bold"),
            corner_radius=12,
            anchor="center"
        )
        self.dd_index_label.place(x=20, y=80)

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

        def strip_after_report(xml_str):
            import re
            match = re.search(r"(.*?</report>)", xml_str, re.DOTALL | re.IGNORECASE)
            return match.group(1) if match else xml_str

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
            # Remove trailing garbage after </report>
            cleaned_xmls = [strip_after_report(xml) for xml in xml_list]
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

        # Extract dd_index
        dd_index_val = "-"
        try:
            dom = xml.dom.minidom.parseString(clean_malformed_xml(xml_data))
            dd_index_nodes = dom.getElementsByTagName("dd_index")
            if dd_index_nodes and dd_index_nodes[0].firstChild:
                dd_index_val = dd_index_nodes[0].firstChild.nodeValue.strip()
        except Exception:
            pass
        self.dd_index_var.set(dd_index_val)

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
            if tag == "tr_report":
                if child.hasAttribute("type") and child.getAttribute("type").strip().upper() == "TR":
                    nu_ptl = self.account_var.get() if hasattr(self, "account_var") else ""
                    # Gather header info from <header> inside tr_report
                    header_info = {}
                    for header in child.getElementsByTagName("header"):
                        for sub in header.childNodes:
                            if sub.nodeType == sub.ELEMENT_NODE:
                                header_info[sub.tagName.lower()] = sub.firstChild.nodeValue.strip() if sub.firstChild else ""
                    # Process each enquiry within this trade reference
                    enquiries = child.getElementsByTagName("enquiry")
                    for idx, enq in enumerate(enquiries, start=1):
                        # Build a flat record for this enquiry
                        record = {}
                        record["NU_PTL"] = nu_ptl
                        record["Row ID"] = str(idx)
                        record["Date"] = header_info.get("date", "")
                        record["Req Name"] = header_info.get("req_name", "")
                        record["Req Com Name"] = header_info.get("req_com_name", "")
                        record["Req Com Addr"] = header_info.get("req_com_addr", "")
                        record["Ref Com Name"] = header_info.get("ref_com_name", "")
                        record["Ref Com Bus"] = header_info.get("ref_com_bus", "")
                        record["Report No"] = header_info.get("report_no", "")
                        record["IC LCNO"] = header_info.get("ic_lcno", "")
                        record["NIC BRNO"] = header_info.get("nic_brno", "")
                        record["Name"] = header_info.get("name", "")
                        record["Enquiry Account No"] = enq.getAttribute("account_no") if enq.hasAttribute("account_no") else ""
                        # Relationship
                        rels = enq.getElementsByTagName("relationship")
                        if rels:
                            rel = rels[0]
                            for data in rel.getElementsByTagName("data"):
                                dname = data.getAttribute("name").strip().lower()
                                text_val = data.firstChild.nodeValue.strip() if data.firstChild else ""
                                if dname == "rel_type":
                                    record["Rel Type"] = text_val
                                elif dname == "rel_status":
                                    record["Rel Status"] = text_val
                                elif dname == "rel_syear":
                                    record["Rel SYear"] = text_val
                                elif dname == "rel_smonth":
                                    record["Rel SMonth"] = text_val
                                elif dname == "rel_sday":
                                    record["Rel SDay"] = text_val
                        # Account Status
                        accs = enq.getElementsByTagName("account_status")
                        if accs:
                            acc = accs[0]
                            for data in acc.getElementsByTagName("data"):
                                dname = data.getAttribute("name").strip().lower()
                                text_val = data.firstChild.nodeValue.strip() if data.firstChild else ""
                                if dname == "statement_date":
                                    record["Statement Date"] = text_val
                                elif dname == "account_rating":
                                    record["Account Rating"] = text_val
                                elif dname == "account_term":
                                    record["Account Term"] = text_val
                                elif dname == "account_limit":
                                    record["Account Limit"] = text_val
                                elif dname == "account_status":
                                    record["Account Status"] = text_val
                                elif dname == "debtor_name":
                                    record["Debtor Name"] = text_val
                                elif dname == "debtor_ic_lcno":
                                    record["Debtor IC LCNO"] = text_val
                                elif dname == "debtor_nic_brno":
                                    record["Debtor NIC BRNO"] = text_val
                                elif dname == "address":
                                    record["Address"] = text_val
                                elif dname == "debt_type":
                                    record["Debt Type"] = text_val
                            # Age breakdown
                            age_elements = acc.getElementsByTagName("age")
                            if age_elements:
                                age_elem = age_elements[0]
                                for item in age_elem.getElementsByTagName("item"):
                                    age_name = item.getAttribute("name") if item.hasAttribute("name") else ""
                                    age_value = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    if age_name:
                                        record[f"Age {age_name}"] = age_value
                        # Contact
                        contacts = enq.getElementsByTagName("contact")
                        if contacts:
                            contact = contacts[0]
                            for data in contact.getElementsByTagName("data"):
                                dname = data.getAttribute("name").strip().lower()
                                text_val = data.firstChild.nodeValue.strip() if data.firstChild else ""
                                if dname == "reference":
                                    record["Contact Reference"] = text_val
                                elif dname == "name":
                                    record["Contact Name"] = text_val
                                elif dname == "address":
                                    record["Contact Address"] = text_val
                                elif dname == "tel_no":
                                    record["Contact Tel No"] = text_val
                                elif dname == "fax_no":
                                    record["Contact Fax No"] = text_val
                                elif dname == "email":
                                    record["Contact Email"] = text_val
                                elif dname == "type":
                                    record["Contact Type"] = text_val
                                elif dname == "type_code":
                                    record["Contact Type Code"] = text_val
                        # Insert as a flat field/value list under a parent node
                        parent_label = f"Trade Reference: {record.get('Name','')} | {record.get('Date','')}"
                        parent_node = self.tree.insert(parent_path, "end", values=[parent_label, ""])
                        for field in [
                            "NU_PTL", "Row ID", "Date", "Req Name", "Req Com Name", "Req Com Addr", "Ref Com Name", "Ref Com Bus",
                            "Report No", "IC LCNO", "NIC BRNO", "Name", "Enquiry Account No", "Rel Type", "Rel Status", "Rel SYear",
                            "Rel SMonth", "Rel SDay", "Statement Date", "Account Rating", "Account Term", "Account Limit",
                            "Account Status", "Debtor Name", "Debtor IC LCNO", "Debtor NIC BRNO", "Address", "Debt Type",
                            "Age 30", "Age 60", "Age 90", "Age 120", "Age 150", "Age 180", "Age 210",
                            "Contact Reference", "Contact Name", "Contact Address", "Contact Tel No", "Contact Fax No",
                            "Contact Email", "Contact Type", "Contact Type Code"
                        ]:
                            self.tree.insert(parent_node, "end", values=[field, record.get(field, "")])
                    continue

            # --- New CTOS XML logic below ---

            # SECTION A (new format)
            if tag == "section_a":
                title = child.getAttribute("title") if child.hasAttribute("title") else "SECTION A"
                self.tree.insert("", "end", values=[title, "-"])
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
                self.tree.insert("", "end", values=[title, "-"])
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
                self.tree.insert("", "end", values=[title, "-"])
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
                self.tree.insert("", "end", values=[title, "-"])
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
            
            # SECTION E (new format)
            if tag == "section_e":
                title = child.getAttribute("title") if child.hasAttribute("title") else "SECTION E"
                self.tree.insert("", "end", values=[title, "-"])
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
        self.is_converting = True
        self.show_progress_popup()
        threading.Thread(target=self.convert_new_ctos_to_excel_thread, daemon=True).start()


    def convert_new_ctos_to_excel_thread(self):

        new_section_columns = {
            "Header&Summary": ["NU", "USR", "CMP", "ACC", "TEL", "FAX", "EDT", "ETM", "EST", "NAME", "IC", "NIC", "IDX", "REF"],
            "Section-A": ["NU", "RID", "NAME", "IC", "NIC", "ADDR", "SRC", "BDT"],
            "Section-B": ["NU", "RID", "TTL", "NAME", "ALS", "IC", "NIC", "REF", "FIRM", "RM1", "RM2", "RM3", "AMT", "ENT"],
            "Section-C": ["NU", "RID", "CO", "ADREG", "LOC", "OBJ", "INC", "LST", "APP", "RSN", "NAME", "NIC", "ADDR", "POS", "CPO", "PD", "SH", "TSH", "RM"],
            "Section-D": [
                "NU", "RID", "RPTTYPE", "STATUS", "TITLE", "SPECIAL_REMARK", "NAME", "NAME_MATCH", "ALIAS", "ADDR",
                "IC_LCNO", "NIC_BRNO", "NIC_BRNO_MATCH", "CASE_NO", "COURT_DETAIL", "FIRM", "PLAINTIFF",
                "ACTION_DATE", "ACTION_SOURCE_DETAIL", "HEAR_DATE", "AMOUNT", "REMARK", "LAWYER", "CEDCON",
                "SETTLEMENT_CODE", "SETTLEMENT_DATE", "SETTLEMENT_SOURCE", "SETTLEMENT_SOURCE_DATE",
                "LATEST_STATUS", "SUBJECT_CMT", "CRA_CMT"
            ],
            "Section-D2": ["NU", "RID"],
            "Section-D4": ["NU", "RID"],
            "Section-ETR_PLUS": ["NU", "RID"],
            "Section-E": [
                "NU_PTL", "ROW_ID",
                "SUBJECT_REF_COM_NAME", "SUBJECT_REF_COM_BUS", "SUBJECT_PARTY_TYPE", "SUBJECT_IC_LCNO", "SUBJECT_NIC_BRNO", "SUBJECT_NAME", "SUBJECT_ADDRESS", "SUBJECT_TREF_DATE",
                "REL_TYPE", "REL_TYPE_CODE", "REL_STATUS", "REL_STATUS_CODE", "REL_ACCOUNT_NO", "REL_SYEAR", "REL_SMONTH", "REL_SDAY", "REL_REMARK",
                "ACC_ACCOUNT_NO", "ACC_STATEMENT_DATE", "ACC_RATING", "ACC_RATING_CODE", "ACC_TERM", "ACC_LIMIT", "ACC_STATUS", "ACC_STATUS_CODE",
                "ACC_DEBTOR_NAME", "ACC_DEBTOR_IC_LCNO", "ACC_DEBTOR_NIC_BRNO", "ACC_ADDRESS", "ACC_DEBT_TYPE", "ACC_DEBT_TYPE_CODE", "ACC_LAST_PAID_AMOUNT",
                "ACC_AGE_30", "ACC_AGE_60", "ACC_AGE_90", "ACC_AGE_120", "ACC_AGE_150", "ACC_AGE_180", "ACC_AGE_OVER_180",
                "LEGAL_ACTION_STATUS",
                "REMINDER_LETTER_TITLE", "REMINDER_LETTER_DATE",
                "DEMAND_LETTER_BY_COMPANY_TITLE", "DEMAND_LETTER_BY_COMPANY_DATE",
                "DEMAND_LETTER_BY_LAWYER_TITLE", "DEMAND_LETTER_BY_LAWYER_DATE", "DEMAND_LETTER_BY_LAWYER_REFERENCE",
                "REF_CONTACT_REF", "REF_CONTACT_NAME", "REF_CONTACT_ADD", "REF_CONTACT_TELNO", "REF_CONTACT_NATURE_OF_BUSINESS",
                "REF_CONTACT_FAXNO", "REF_CONTACT_EMAIL", "REF_CONTACT_TYPE", "REF_CONTACT_TYPE_CODE"
            ]
        }

        new_sheets_data = {k: [] for k in new_section_columns}
        total = len(self.filtered_data)

        for index, (_, row) in enumerate(self.filtered_data.iterrows()):
            nu_ptl = row.get("NU_PTL", f"Row{index}")
            xml_data = clean_malformed_xml(row.get("XML", ""))
            if pd.isna(xml_data) or not str(xml_data).strip():
                continue
            try:
                dom = xml.dom.minidom.parseString(xml_data)
                root = dom.documentElement
                is_new_ctos = any(
                    root.getElementsByTagName(tag)
                    for tag in ["section_d2", "section_d4", "section_etl", "section_etr_plus"]
                )
                if not is_new_ctos:
                    continue

                # --- Header&Summary ---
                header_record = {col: "" for col in new_section_columns["Header&Summary"]}
                header_record["NU"] = nu_ptl
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
                        rec["NU"] = nu_ptl
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

                # --- Section-B ---
                for section_b in root.getElementsByTagName("section_b"):
                    for record in section_b.getElementsByTagName("record"):
                        rec = {col: "" for col in new_section_columns["Section-B"]}
                        rec["NU"] = nu_ptl
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
                        new_sheets_data["Section-B"].append(rec)

                # --- Section-C ---
                for section_c in root.getElementsByTagName("section_c"):
                    for record in section_c.getElementsByTagName("record"):
                        rec = {col: "" for col in new_section_columns["Section-C"]}
                        rec["NU"] = nu_ptl
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
                                elif tag == "LASTDOC":
                                    rec["LST"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "APPOINT":
                                    rec["APP"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "RESIGN":
                                    rec["RSN"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
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
                                elif tag == "PAIDUP":
                                    rec["PD"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "SHARES":
                                    rec["SH"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "TOTAL_SHARES_PERCENTAGE":
                                    rec["TSH"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                elif tag == "REMARK":
                                    rec["RM"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                        new_sheets_data["Section-C"].append(rec)

                # --- Section-D ---
                for section_d in root.getElementsByTagName("section_d"):
                    for record in section_d.getElementsByTagName("record"):
                        rec = {col: "" for col in new_section_columns["Section-D"]}
                        rec["NU"] = nu_ptl
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
                        new_sheets_data["Section-D"].append(rec)

                # --- Section-D2, D4, ETL, ETR_PLUS (minimal: NU, RID) ---
                for sec_tag, sheet in [("section_d2", "Section-D2"), ("section_d4", "Section-D4"), ("section_etr_plus", "Section-ETR_PLUS")]:
                    for section in root.getElementsByTagName(sec_tag):
                        for record in section.getElementsByTagName("record"):
                            rec = {col: "" for col in new_section_columns[sheet]}
                            rec["NU"] = nu_ptl
                            rec["RID"] = record.getAttribute("seq") if record.hasAttribute("seq") else ""
                            new_sheets_data[sheet].append(rec)

                # --- Section-E (Trade Referees) ---
                section_e_columns = new_section_columns["Section-E"]
                row_id_counter = 1
                for section_e in root.getElementsByTagName("section_e"):
                    for enquiry in section_e.getElementsByTagName("enquiry"):
                        rec = {col: "" for col in section_e_columns}
                        rec["NU_PTL"] = nu_ptl
                        rec["ROW_ID"] = str(row_id_counter)
                        row_id_counter += 1

                        # Subject
                        subject = enquiry.getElementsByTagName("subject")
                        if subject:
                            subj = subject[0]
                            for item in subj.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    if tag == "REF_COM_NAME":
                                        rec["SUBJECT_REF_COM_NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "REF_COM_BUS":
                                        rec["SUBJECT_REF_COM_BUS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "PARTY_TYPE":
                                        rec["SUBJECT_PARTY_TYPE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "IC_LCNO":
                                        rec["SUBJECT_IC_LCNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "NIC_BRNO":
                                        rec["SUBJECT_NIC_BRNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "NAME":
                                        rec["SUBJECT_NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ADDRESS":
                                        rec["SUBJECT_ADDRESS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "TREF_DATE":
                                        rec["SUBJECT_TREF_DATE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""

                        # Relationship
                        relationship = enquiry.getElementsByTagName("relationship")
                        if relationship:
                            rel = relationship[0]
                            for item in rel.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    if tag == "REL_TYPE":
                                        rec["REL_TYPE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                        rec["REL_TYPE_CODE"] = item.getAttribute("code") if item.hasAttribute("code") else ""
                                    elif tag == "REL_STATUS":
                                        rec["REL_STATUS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                        rec["REL_STATUS_CODE"] = item.getAttribute("code") if item.hasAttribute("code") else ""
                                    elif tag == "ACCOUNT_NO":
                                        rec["REL_ACCOUNT_NO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "REL_SYEAR":
                                        rec["REL_SYEAR"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "REL_SMONTH":
                                        rec["REL_SMONTH"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "REL_SDAY":
                                        rec["REL_SDAY"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "REMARK":
                                        rec["REL_REMARK"] = item.firstChild.nodeValue.strip() if item.firstChild else ""

                        # Account Status
                        account_status = enquiry.getElementsByTagName("account_status")
                        if account_status:
                            acc = account_status[0]
                            for item in acc.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    if tag == "ACCOUNT_NO":
                                        rec["ACC_ACCOUNT_NO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "STATEMENT_DATE":
                                        rec["ACC_STATEMENT_DATE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ACCOUNT_RATING":
                                        rec["ACC_RATING"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                        rec["ACC_RATING_CODE"] = item.getAttribute("code") if item.hasAttribute("code") else ""
                                    elif tag == "ACCOUNT_TERM":
                                        rec["ACC_TERM"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ACCOUNT_LIMIT":
                                        rec["ACC_LIMIT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ACCOUNT_STATUS":
                                        rec["ACC_STATUS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                        rec["ACC_STATUS_CODE"] = item.getAttribute("code") if item.hasAttribute("code") else ""
                                    elif tag == "DEBTOR_NAME":
                                        rec["ACC_DEBTOR_NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "DEBTOR_IC_LCNO":
                                        rec["ACC_DEBTOR_IC_LCNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "DEBTOR_NIC_BRNO":
                                        rec["ACC_DEBTOR_NIC_BRNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "ADDRESS":
                                        rec["ACC_ADDRESS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "DEBT_TYPE":
                                        rec["ACC_DEBT_TYPE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                        rec["ACC_DEBT_TYPE_CODE"] = item.getAttribute("code") if item.hasAttribute("code") else ""
                                    elif tag == "LAST_PAID_AMOUNT":
                                        rec["ACC_LAST_PAID_AMOUNT"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "AGE":
                                        for age_item in item.childNodes:
                                            if age_item.nodeType == age_item.ELEMENT_NODE:
                                                age_tag = age_item.tagName.strip().upper()
                                                if age_tag == "AGE_30":
                                                    rec["ACC_AGE_30"] = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""
                                                elif age_tag == "AGE_60":
                                                    rec["ACC_AGE_60"] = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""
                                                elif age_tag == "AGE_90":
                                                    rec["ACC_AGE_90"] = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""
                                                elif age_tag == "AGE_120":
                                                    rec["ACC_AGE_120"] = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""
                                                elif age_tag == "AGE_150":
                                                    rec["ACC_AGE_150"] = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""
                                                elif age_tag == "AGE_180":
                                                    rec["ACC_AGE_180"] = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""
                                                elif age_tag == "AGE_OVER_180":
                                                    rec["ACC_AGE_OVER_180"] = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""

                        # Legal Action
                        legal_action = enquiry.getElementsByTagName("legal_action")
                        if legal_action:
                            legal = legal_action[0]
                            rec["LEGAL_ACTION_STATUS"] = legal.getAttribute("status") if legal.hasAttribute("status") else ""
                            for item in legal.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    if tag == "REMINDER_LETTER":
                                        for sub in item.childNodes:
                                            if sub.nodeType == sub.ELEMENT_NODE:
                                                subtag = sub.tagName.strip().upper()
                                                if subtag == "TITLE":
                                                    rec["REMINDER_LETTER_TITLE"] = sub.firstChild.nodeValue.strip() if sub.firstChild else ""
                                                elif subtag == "DATE":
                                                    rec["REMINDER_LETTER_DATE"] = sub.firstChild.nodeValue.strip() if sub.firstChild else ""
                                    elif tag == "DEMAND_LETTER_BY_COMPANY":
                                        for sub in item.childNodes:
                                            if sub.nodeType == sub.ELEMENT_NODE:
                                                subtag = sub.tagName.strip().upper()
                                                if subtag == "TITLE":
                                                    rec["DEMAND_LETTER_BY_COMPANY_TITLE"] = sub.firstChild.nodeValue.strip() if sub.firstChild else ""
                                                elif subtag == "DATE":
                                                    rec["DEMAND_LETTER_BY_COMPANY_DATE"] = sub.firstChild.nodeValue.strip() if sub.firstChild else ""
                                    elif tag == "DEMAND_LETTER_BY_LAWYER":
                                        for sub in item.childNodes:
                                            if sub.nodeType == sub.ELEMENT_NODE:
                                                subtag = sub.tagName.strip().upper()
                                                if subtag == "TITLE":
                                                    rec["DEMAND_LETTER_BY_LAWYER_TITLE"] = sub.firstChild.nodeValue.strip() if sub.firstChild else ""
                                                elif subtag == "DATE":
                                                    rec["DEMAND_LETTER_BY_LAWYER_DATE"] = sub.firstChild.nodeValue.strip() if sub.firstChild else ""
                                                elif subtag == "REFERENCE":
                                                    rec["DEMAND_LETTER_BY_LAWYER_REFERENCE"] = sub.firstChild.nodeValue.strip() if sub.firstChild else ""

                        # Referee Contact
                        referee_contact = enquiry.getElementsByTagName("referee_contact")
                        if referee_contact:
                            refc = referee_contact[0]
                            for item in refc.childNodes:
                                if item.nodeType == item.ELEMENT_NODE:
                                    tag = item.tagName.strip().upper()
                                    if tag == "CONTACT_REF":
                                        rec["REF_CONTACT_REF"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "CONTACT_NAME":
                                        rec["REF_CONTACT_NAME"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "CONTACT_ADD":
                                        rec["REF_CONTACT_ADD"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "CONTACT_TELNO":
                                        rec["REF_CONTACT_TELNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "CONTACT_NATURE_OF_BUSINESS":
                                        rec["REF_CONTACT_NATURE_OF_BUSINESS"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "CONTACT_FAXNO":
                                        rec["REF_CONTACT_FAXNO"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "CONTACT_EMAIL":
                                        rec["REF_CONTACT_EMAIL"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                    elif tag == "CONTACT_TYPE":
                                        rec["REF_CONTACT_TYPE"] = item.firstChild.nodeValue.strip() if item.firstChild else ""
                                        rec["REF_CONTACT_TYPE_CODE"] = item.getAttribute("code") if item.hasAttribute("code") else ""

                        new_sheets_data["Section-E"].append(rec)

            except Exception as e:
                msg = f"Error parsing XML for NU_PTL {nu_ptl}: {str(e)}"
                self.after(0, self.append_error, msg)
                continue

            if index % 10 == 0 or index + 1 == total:
                progress = (index + 1) / total
                self.after(0, self.update_progress, progress, index + 1, total)

        # Ensure all sheets exist and have at least header row
        for sheet in ["Section-D", "Section-D2", "Section-D4", "Section-ETR_PLUS", "Section-E"]:
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
        self.is_converting = True
        self.show_progress_popup()
        threading.Thread(target=self.convert_to_excel_thread, daemon=True).start()

    def update_progress(self, progress, index, total):
        self.progress_bar.set(progress)
        self.status_label.configure(text=f"Processing {index} of {total}")
        # Removed: self.popup.update()  # <-- Avoid explicit update call here
        
    def convert_to_excel_thread(self):
        # Define columns for old CTOS sections
        old_section_columns = {
            "Header&Summary": ["NU_PTL", "user", "company", "account", "tel", "fax", "enq_date", "enq_time", "enq_status", "IC_LCNO", "NIC_BRNO", "NAME", "ALIAS", "STAT", "REF"],
            "Section-A": ["NU_PTL", "Record_ID", "ICNO", "MATCH", "NEWIC", "MATCH1", "NAME", "MATCH2", "ADDR", "ADDR1", "REMARK"],
            "Section-B": ["NU_PTL", "Record_ID", "CODE", "NAME", "MATCH", "ALIAS", "IC_LCNO", "NIC_BRNO", "REF", "CONUM", "CONAME", "REMARK", "REMARK2", "REMARK3", "AMOUNT", "ENTRY"],
            "Section-C": ["NU_PTL", "Record_ID", "EXTRANAME", "EXTRALOCAL", "EXTRALOCALA", "OBJECT", "INCORPRATION", "COMPANY PAIDUP", "SEARCH DATE", "EXTRALASTDOC", "NAME", "IC_LCNO", "NIC_BRNO", "STATUS", "SHARES", "REAMARK", "APPOINTED", "RESIGNED", "ADDRESS"],
            "Section-D": ["NU_PTL", "Record_ID", "ZTITLE", "ZSPECIAL", "NAME", "MATCH", "ALIAS", "I/C NO", "NEW IC", "REMARK", "ADDRESS", "FIRM", "PLAINTIFF", "CASE NO", "ZCOURT", "ACTION DATE", "ZNTPAP", "HEARING DATE", "AMOUNT", "SOLCTR", "LAWADD1", "TEL", "LAWADD2", "REF", "LAWADD3", "PLAINTIFF CONTACT", "CEDCONADD1", "CEDCONADD2", "CEDCONADD3"],
            "Section-E": ["NU_PTL", "Record_ID", "REFEREE", "INCORPORATION DATE", "NATURE OF BUSINESS", "ADDRESS", "TR_URL"]
        }
        # Define header for new Trade Reference sheet (horizontal)
        trade_reference_columns = [
            "NU_PTL", "Row ID", "Date", "Req Name", "Req Com Name", "Req Com Addr", "Ref Com Name", "Ref Com Bus",
            "Report No", "IC LCNO", "NIC BRNO", "Name", "Enquiry Account No", "Rel Type",
            "Rel Status", "Rel SYear", "Rel SMonth", "Rel SDay", "Statement Date", "Account Rating",
            "Account Term", "Account Limit", "Account Status", "Debtor Name", "Debtor IC LCNO",
            "Debtor NIC BRNO", "Address", "Debt Type"
        ]
        # Extend additional fields for ages and contact details
        full_trade_reference_columns = trade_reference_columns + [
            "Age 30", "Age 60", "Age 90", "Age 120", "Age 150", "Age 180", "Age 210",
            "Contact Reference", "Contact Name", "Contact Address", "Contact Tel No", "Contact Fax No", "Contact Email", "Contact Type", "Contact Type Code"
        ]

        # Initialize data storage for old sections and trade reference records
        old_sheets_data = {k: [] for k in old_section_columns}
        trade_reference_data = []  # Will store each trade reference record separately
        total = len(self.filtered_data)

        for index, (_, row) in enumerate(self.filtered_data.iterrows()):
            nu_ptl = row.get("NU_PTL", f"Row{index}")
            xml_data = clean_malformed_xml(row.get("XML", ""))
            if pd.isna(xml_data) or not str(xml_data).strip():
                continue

            try:
                dom = xml.dom.minidom.parseString(xml_data)
                root = dom.documentElement

                # --- Extract Header&Summary ---
                header_record = {col: "" for col in old_section_columns["Header&Summary"]}
                header_record["NU_PTL"] = nu_ptl
                for header in root.getElementsByTagName("header"):
                    for node in header.childNodes:
                        if node.nodeType == node.ELEMENT_NODE:
                            tag = node.tagName.strip()
                            if tag in old_section_columns["Header&Summary"]:
                                header_record[tag] = node.firstChild.nodeValue.strip() if node.firstChild else "-"
                for summary in root.getElementsByTagName("summary"):
                    for enq_sum in summary.getElementsByTagName("enq_sum"):
                        for fs in enq_sum.getElementsByTagName("field_sum"):
                            if fs.hasAttribute("name"):
                                field = fs.getAttribute("name").strip()
                                if field in old_section_columns["Header&Summary"]:
                                    value = fs.firstChild.nodeValue.strip() if fs.firstChild else "-"
                                    header_record[field] = value
                old_sheets_data["Header&Summary"].append(header_record)

                # --- Extract Sections A, B, C, D, E ---
                for section in root.getElementsByTagName("section"):
                    sec_id = section.getAttribute("id").strip().upper()
                    section_key = f"Section-{sec_id}"
                    if section_key not in old_sheets_data:
                        continue
                    for rec in section.getElementsByTagName("record"):
                        record = {col: "" for col in old_section_columns[section_key]}
                        record["NU_PTL"] = nu_ptl
                        record_id = rec.getAttribute("seq").strip() if rec.hasAttribute("seq") else ""
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
                                for expected in old_section_columns.get(section_key, []):
                                    if expected.upper() == key.upper():
                                        matched_field = expected
                                        break
                                if matched_field:
                                    break
                            if matched_field:
                                value = data.firstChild.nodeValue.strip() if data.firstChild else ""
                                record[matched_field] = value
                        old_sheets_data[section_key].append(record)

                # --- Extract Trade Reference Records ---
                # Process each <tr_report> element with type="TR"
                for tr_report in root.getElementsByTagName("tr_report"):
                    if tr_report.hasAttribute("type") and tr_report.getAttribute("type").strip().upper() == "TR":
                        # Gather header info from tr_report
                        header_info = {}
                        for header in tr_report.getElementsByTagName("header"):
                            for node in header.childNodes:
                                if node.nodeType == node.ELEMENT_NODE:
                                    tag = node.tagName.strip().lower()
                                    text_val = node.firstChild.nodeValue.strip() if node.firstChild else ""
                                    header_info[tag] = text_val
                        # Process each enquiry element under this tr_report.
                        # If there are enquiries, output one row per enquiry and add a Row ID.
                        enquiries = tr_report.getElementsByTagName("enquiry")
                        row_id_counter = 1
                        if enquiries:
                            for enq in enquiries:
                                trade_record = {col: "" for col in full_trade_reference_columns}
                                trade_record["NU_PTL"] = nu_ptl
                                trade_record["Row ID"] = str(row_id_counter)
                                row_id_counter += 1
                                # Populate header fields
                                trade_record["Date"] = header_info.get("date", "")
                                trade_record["Req Name"] = header_info.get("req_name", "")
                                trade_record["Req Com Name"] = header_info.get("req_com_name", "")
                                trade_record["Req Com Addr"] = header_info.get("req_com_addr", "")
                                trade_record["Ref Com Name"] = header_info.get("ref_com_name", "")
                                trade_record["Ref Com Bus"] = header_info.get("ref_com_bus", "")
                                trade_record["Report No"] = header_info.get("report_no", "")
                                trade_record["IC LCNO"] = header_info.get("ic_lcno", "")
                                trade_record["NIC BRNO"] = header_info.get("nic_brno", "")
                                trade_record["Name"] = header_info.get("name", "")
                                
                                # Enquiry attribute: Account Number
                                if enq.hasAttribute("account_no"):
                                    trade_record["Enquiry Account No"] = enq.getAttribute("account_no").strip()
                                # Relationship section
                                rels = enq.getElementsByTagName("relationship")
                                if rels:
                                    rel = rels[0]
                                    for data in rel.getElementsByTagName("data"):
                                        dname = data.getAttribute("name").strip().lower()
                                        text_val = data.firstChild.nodeValue.strip() if data.firstChild else ""
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
                                # Account_status section (if present)
                                accs = enq.getElementsByTagName("account_status")
                                if accs:
                                    acc = accs[0]
                                    for data in acc.getElementsByTagName("data"):
                                        dname = data.getAttribute("name").strip().lower()
                                        text_val = data.firstChild.nodeValue.strip() if data.firstChild else ""
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
                                    # Process ages within <age> element if present
                                    age_elements = acc.getElementsByTagName("age")
                                    if age_elements:
                                        age_elem = age_elements[0]
                                        for age_item in age_elem.childNodes:
                                            if age_item.nodeType == age_item.ELEMENT_NODE:
                                                age_tag = age_item.tagName.strip().lower()
                                                value = age_item.firstChild.nodeValue.strip() if age_item.firstChild else ""
                                                if age_tag == "30":
                                                    trade_record["Age 30"] = value
                                                elif age_tag == "60":
                                                    trade_record["Age 60"] = value
                                                elif age_tag == "90":
                                                    trade_record["Age 90"] = value
                                                elif age_tag == "120":
                                                    trade_record["Age 120"] = value
                                                elif age_tag == "150":
                                                    trade_record["Age 150"] = value
                                                elif age_tag == "180":
                                                    trade_record["Age 180"] = value
                                                elif age_tag == "210":
                                                    trade_record["Age 210"] = value
                                # Contact section
                                contacts = enq.getElementsByTagName("contact")
                                if contacts:
                                    contact = contacts[0]
                                    for data in contact.getElementsByTagName("data"):
                                        dname = data.getAttribute("name").strip().lower()
                                        text_val = data.firstChild.nodeValue.strip() if data.firstChild else ""
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
                            # If no enquiry exists, output one record with header info only.
                            trade_record = {col: "" for col in full_trade_reference_columns}
                            trade_record["NU_PTL"] = nu_ptl
                            trade_record["Row ID"] = "1"
                            trade_record["Date"] = header_info.get("date", "")
                            trade_record["Req Name"] = header_info.get("req_name", "")
                            trade_record["Req Com Name"] = header_info.get("req_com_name", "")
                            trade_record["Req Com Addr"] = header_info.get("req_com_addr", "")
                            trade_record["Ref Com Name"] = header_info.get("ref_com_name", "")
                            trade_record["Ref Com Bus"] = header_info.get("ref_com_bus", "")
                            trade_record["Report No"] = header_info.get("report_no", "")
                            trade_record["IC LCNO"] = header_info.get("ic_lcno", "")
                            trade_record["NIC BRNO"] = header_info.get("nic_brno", "")
                            trade_record["Name"] = header_info.get("name", "")
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
            # Write old CTOS sheets
            for sheet_name, records in old_sheets_data.items():
                if records:
                    df = pd.DataFrame(records)
                    df = df.reindex(columns=old_section_columns[sheet_name])
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            # Write Trade Reference sheet if any records were found
            if trade_reference_data:
                df_tr = pd.DataFrame(trade_reference_data)
                df_tr = df_tr.reindex(columns=full_trade_reference_columns)
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
        
        # dd_index badge (square)
        self.dd_index_var = tk.StringVar(value="-")
        self.dd_index_label = ctk.CTkLabel(
            self,
            textvariable=self.dd_index_var,
            width=48,
            height=48,
            fg_color="#1976d2",
            text_color="#fff",
            font=ctk.CTkFont(size=18, weight="bold"),
            corner_radius=12,
            anchor="center"
        )
        self.dd_index_label.place(x=20, y=80)
    
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
                nuptl_to_xmls[base_nuptl].append("".join(xmls))

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

            # Extract dd_index
            dd_index_val = "-"
            try:
                dom = xml.dom.minidom.parseString(clean_malformed_xml(raw_xml))
                dd_index_nodes = dom.getElementsByTagName("dd_index")
                if dd_index_nodes and dd_index_nodes[0].firstChild:
                    dd_index_val = dd_index_nodes[0].firstChild.nodeValue.strip()
            except Exception:
                pass
            self.dd_index_var.set(dd_index_val)

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
            self.dd_index_var.set("-")
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
