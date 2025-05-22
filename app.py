import threading
import time
import math
import xlsxwriter
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

        # Shared data for all views
        self.shared_data = None  # This will hold the imported Excel data

        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")

        # Import Excel Button
        self.import_button = ctk.CTkButton(
            self.sidebar, text="Import Excel File", command=self.import_excel
        )
        self.import_button.pack(pady=20, padx=20)

        # XML Format Button
        self.xml_format_button = ctk.CTkButton(
            self.sidebar, text="XML Format", command=self.show_xml_format
        )
        self.xml_format_button.pack(pady=20, padx=20)

        # CTOS Report Button
        self.ctos_report_button = ctk.CTkButton(
            self.sidebar, text="CTOS Report", command=self.show_ctos_report
        )
        self.ctos_report_button.pack(pady=20, padx=20)
        
        self.ctos_summary_butoon = ctk.CTkButton(
            self.sidebar, text="CTOS Summary", command=self.show_ctos_summary
        )
        self.ctos_summary_butoon.pack(pady=20, padx=20)

        # Add space between sidebar and main content
        self.sidebar_spacer = ctk.CTkFrame(self, width=20, fg_color="transparent")
        self.sidebar_spacer.pack(side="left", fill="y")

        # Main Content Area
        self.main_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_frame.pack(side="right", fill="both", expand=True)

        # Initialize Views
        self.xml_format_view = XMLFormatView(self.main_frame, self)
        self.ctos_report_view = CTOSReportView(self.main_frame, self)
        self.ctos_summary_view = CTOSSummaryView(self.main_frame, self)

        # Show Default View
        self.show_xml_format()
        
    def show_progress_popup(self, title="Processing...", message="Please wait..."):
        self.progress_popup = ctk.CTkToplevel(self)
        self.progress_popup.title(title)
        self.progress_popup.geometry("300x100")
        self.progress_popup.resizable(False, False)
        self.progress_popup.grab_set()  # Block interaction with main window
        self.progress_popup.attributes("-topmost", True)  # ✅ Always on top

        ctk.CTkLabel(self.progress_popup, text=message).pack(pady=10)
        self.progress_bar = ctk.CTkProgressBar(self.progress_popup, mode="indeterminate")
        self.progress_bar.pack(padx=20, pady=10, fill="x")
        self.progress_bar.start()

    def destroy_progress_popup(self):
        if hasattr(self, "progress_popup") and self.progress_popup.winfo_exists():
            self.progress_bar.stop()
            self.progress_popup.destroy()

    def import_excel(self):
        def import_thread():
            try:
                file_path = filedialog.askopenfilename(
                    title="Select Excel File",
                    filetypes=[("Excel Files", "*.xlsx *.xls")]
                )
                if not file_path:
                    self.after(0, self.destroy_progress_popup)
                    return

                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip().str.upper()  # Normalize column names

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
                    messagebox.showinfo("Success", "Excel file imported and XML cleaned successfully!")

                self.after(0, update_data)

            except Exception as e:
                self.after(0, self.destroy_progress_popup)
                self.after(0, lambda: messagebox.showerror("Error", f"Error importing Excel file: {e}"))

        self.show_progress_popup(title="Importing Excel", message="Cleaning XML records...")
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

class CTOSReportView(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app  # Reference to the main app to access shared data
        self.account_var = tk.StringVar()
        self.search_var = tk.StringVar()
        self.all_accounts = []
        self.current_index = 0  # Track the current NU_PTL index
        self.filtered_data = None  # Store filtered data for navigation
        style = ttk.Style()
        self.tabview = CTkTabview(self)
        self.tabview.pack(fill="both", expand=True)
        self.treeviews = {}
        style.configure("Treeview", rowheight=25, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[('selected', '#6fa8dc')])


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

        # Previous Button
        self.prev_button = ctk.CTkButton(self.control_frame, text="Previous", command=self.go_to_previous)
        self.prev_button.grid(row=0, column=0, padx=10, pady=5, sticky="e")

        # ttk Combobox
        self.ttk_style = ttk.Style()
        self.ttk_style.theme_use('clam')
        self.account_combobox = ttk.Combobox(
            self.control_frame, textvariable=self.account_var, values=[], width=25
        )
        self.account_combobox.grid(row=0, column=1, padx=10, pady=5)
        self.account_combobox.bind("<<ComboboxSelected>>", self.display_data)
        self.account_combobox.bind("<KeyRelease>", self.on_account_typing)

        # Next Button
        self.next_button = ctk.CTkButton(self.control_frame, text="Next", command=self.go_to_next)
        self.next_button.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # Convert to Excel Button
        self.convert_button = ctk.CTkButton(self.control_frame, text="Convert to Excel", command=self.convert_to_excel)
        self.convert_button.grid(row=0, column=4, padx=5)
        

        # Treeview for displaying parsed XML data
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(self.tree_frame, show="headings")
        self.tree.pack(fill="both", expand=True, side="left")

        # Add a scrollbar
        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        # Create a refresh button
        refresh_button = ctk.CTkButton(self, text="Refresh", command=self.refresh_data)
        refresh_button.pack(pady=10)
    
    def on_account_typing(self, event):
        typed = self.account_var.get().lower()
        
        # Filter values that contain the typed substring
        filtered = [acct for acct in self.all_accounts if typed in acct.lower()]
        
        # Update combobox values dynamically
        self.account_combobox['values'] = filtered

        # Optionally, show dropdown
        if filtered:
            self.account_combobox.event_generate('<Down>')

        
    def refresh_data(self):
        xml_format_view = self.app.xml_format_view

        # Ensure there is XML data to process
        if not xml_format_view.xml_data:
            return

        # Clean the XML data
        cleaned_data = {
            key: clean_malformed_xml(value)
            for key, value in xml_format_view.xml_data.items()
        }

        # Rebuild filtered_data DataFrame from cleaned XML
        self.filtered_data = pd.DataFrame.from_dict(cleaned_data, orient="index", columns=["XML"])
        self.filtered_data.reset_index(inplace=True)
        self.filtered_data.rename(columns={"index": "NU_PTL"}, inplace=True)

        # Populate account list and update combobox
        self.all_accounts = self.filtered_data["NU_PTL"].tolist()
        self.account_combobox['values'] = self.all_accounts  # <-- Add this line
        if self.all_accounts:
            self.account_combobox.current(self.current_index)   

        self.display_data()

    def display_data(self, event=None):
        if self.filtered_data is not None and not self.filtered_data.empty:
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
                return

            try:
                dom = xml.dom.minidom.parseString(xml_data)
                root = dom.documentElement

                # Clear all tabs
                for tab_name in self.tabview.tabs():
                    self.tabview.delete(tab_name)
                self.treeviews.clear()

                # Add "Header & Summary" tab
                self.tabview.add("Header & Summary")
                header_frame = self.tabview.tab("Header & Summary")
                header_tree = self.create_treeview(header_frame)
                self.treeviews["Header & Summary"] = header_tree

                # Parse <header>
                for header in root.getElementsByTagName("header"):
                    for node in header.childNodes:
                        if node.nodeType == node.ELEMENT_NODE:
                            field = node.tagName
                            value = node.firstChild.nodeValue.strip() if node.firstChild else "-"
                            header_tree.insert("", "end", values=[field, value])

                # Parse <summary>
                for summary in root.getElementsByTagName("summary"):
                    for enq_sum in summary.getElementsByTagName("enq_sum"):
                        for fs in enq_sum.getElementsByTagName("field_sum"):
                            field = fs.getAttribute("name")
                            value = fs.firstChild.nodeValue.strip() if fs.firstChild else "-"
                            header_tree.insert("", "end", values=[field, value])

                # Add section tabs A–E
                for section in root.getElementsByTagName("section"):
                    sec_id = section.getAttribute("id").strip().upper()
                    if sec_id not in ["A", "B", "C", "D", "E"]:
                        continue
                    section_name = f"Section {sec_id}"
                    self.tabview.add(section_name)
                    sec_frame = self.tabview.tab(section_name)
                    sec_tree = self.create_treeview(sec_frame)
                    self.treeviews[section_name] = sec_tree

                    for record in section.getElementsByTagName("record"):
                        seq = record.getAttribute("seq")
                        sec_tree.insert("", "end", values=["Record", seq])
                        for data in record.getElementsByTagName("data"):
                            caption = data.getAttribute("caption").strip()
                            name = data.getAttribute("name").strip()
                            field = caption if caption else name
                            value = data.firstChild.nodeValue.strip() if data.firstChild else "-"
                            sec_tree.insert("", "end", values=[field, value])

            except Exception as e:
                self.tabview.add("Error")
                error_tree = self.create_treeview(self.tabview.tab("Error"))
                error_tree.insert("", "end", values=["Error", str(e)])

    def create_treeview(self, parent):
        tree = ttk.Treeview(parent, columns=("Field", "Value"), show="headings")
        tree.heading("Field", text="Field")
        tree.heading("Value", text="Value")
        tree.column("Field", anchor="w", width=300)
        tree.column("Value", anchor="w", width=600)
        tree.pack(fill="both", expand=True)
        return tree

    def parse_xml_to_treeview(self, node, parent_path=""):
        for child in node.childNodes:
            if child.nodeType != xml.dom.minidom.Node.ELEMENT_NODE:
                continue

            tag = child.tagName

            # 1. Handle <enq_report id="...">
            if tag == "enq_report" and child.hasAttribute("id"):
                field = "Report ID"
                value = child.getAttribute("id")
                self.tree.insert("", "end", values=[field, value])
                self.parse_xml_to_treeview(child, field)
                continue

            if tag == "header":
                for sub in child.childNodes:
                    if sub.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                        sub_tag = sub.tagName
                        value = sub.firstChild.nodeValue.strip() if sub.firstChild and sub.firstChild.nodeValue else "-"
                        self.tree.insert("", "end", values=[sub_tag, value])
                continue

            # 3. Handle <summary> contents
            if tag == "summary":
                has_field_sum = False

                for sub in child.getElementsByTagName("enq_sum"):
                    # Try to find field_sum elements
                    field_sum_nodes = sub.getElementsByTagName("field_sum")
                    if field_sum_nodes:
                        has_field_sum = True
                        for fs in field_sum_nodes:
                            if fs.nodeType == xml.dom.minidom.Node.ELEMENT_NODE and fs.hasAttribute("name"):
                                field = fs.getAttribute("name").strip()
                                value = fs.firstChild.nodeValue.strip() if fs.firstChild and fs.firstChild.nodeValue else "-"
                                self.tree.insert("", "end", values=[field, value])

                # If no field_sum found, fall back to tagName-based entries
                if not has_field_sum:
                    for sub in child.getElementsByTagName("enq_sum"):
                        for item in sub.childNodes:
                            if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                field = item.tagName.strip()
                                value = item.firstChild.nodeValue.strip() if item.firstChild and item.firstChild.nodeValue else "-"
                                self.tree.insert("", "end", values=[field, value])
                continue

            # 4. Handle <section title="...">
            if tag == "section" and child.hasAttribute("title"):
                title = child.getAttribute("title").strip()
                self.tree.insert("", "end", values=[title, "-"])
                self.parse_xml_to_treeview(child, title)
                continue

            # 5. Handle <record seq="...">
            if tag == "record" and child.hasAttribute("seq"):
                seq = child.getAttribute("seq").strip()
                self.tree.insert("", "end", values=["record", seq])
                self.parse_xml_to_treeview(child, f"record_{seq}")
                continue

            # 6. Handle <data caption="..." name="...">...</data>
            if tag == "data":
                caption = child.getAttribute("caption").strip()
                name = child.getAttribute("name").strip()
                field = caption if caption else name
                value = child.firstChild.nodeValue.strip() if child.firstChild and child.firstChild.nodeValue else ""
                self.tree.insert("", "end", values=[field, value])
                continue

            # 🔁 Recurse into children by default
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

        # Previous Button
        self.prev_button = ctk.CTkButton(control_frame, text="Previous", command=self.go_to_previous)
        self.prev_button.grid(row=0, column=0, padx=10, pady=5, sticky="e")

        # ttk Combobox
        self.ttk_style = ttk.Style()
        self.ttk_style.theme_use('clam')
        self.account_combobox = ttk.Combobox(
            control_frame, textvariable=self.account_var, values=[], width=25
        )
        self.account_combobox.grid(row=0, column=1, padx=10, pady=5)
        self.account_combobox.bind("<<ComboboxSelected>>", self.display_xml_data)
        self.account_combobox.bind("<KeyRelease>", self.on_account_typing)

        # Next Button
        self.next_button = ctk.CTkButton(control_frame, text="Next", command=self.go_to_next)
        self.next_button.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # --- XML Display ---
        self.xml_display = ctk.CTkTextbox(self, width=600, height=300)
        self.xml_display.pack(pady=10, fill="both", expand=True)

        # Refresh Button
        self.refresh_button = ctk.CTkButton(self, text="Refresh Data", command=self.refresh_data)
        self.refresh_button.pack(pady=10)
    
    def on_account_typing(self, event):
        typed = self.account_var.get().lower()
        
        # Filter values that contain the typed substring
        filtered = [acct for acct in self.all_accounts if typed in acct.lower()]
        
        # Update combobox values dynamically
        self.account_combobox['values'] = filtered

        # Optionally, show dropdown
        if filtered:
            self.account_combobox.event_generate('<Down>')

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
        self.headers = ["", "Total"]
        self.sections = ["A", "B", "C", "D", "E"]
        self.rows = [""] + self.sections  # Blank (Total), then A–E

        self.data = []  # List of records with 'NU_PTL' and 'Section'
        self.create_main_layout()

    def create_main_layout(self):
        self.header_label = ctk.CTkLabel(self, text="CTOS Summary", font=ctk.CTkFont(size=16, weight="bold"))
        self.header_label.pack(pady=(10, 5))

        self.control_frame = ctk.CTkFrame(self)
        self.control_frame.pack(fill="x", padx=10, pady=5)

        self.table_frame = ctk.CTkFrame(self)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Example button to trigger summary
        self.refresh_button = ctk.CTkButton(self.control_frame, text="Refresh Summary", command=self.refresh_summary)
        self.refresh_button.pack(side="left")

        self.create_summary_table({})  # Initial blank table

    def refresh_summary(self):
        records = self.app.all_parsed_data  # Use your actual parsed data
        summary = self.calculate_summary(records)
        self.create_summary_table(summary)

    def calculate_summary(self, records):
        unique_nu_ptls = set()
        section_record_count = defaultdict(int)  # section => total record count

        for record in records:
            nu_ptl = record.get("NU_PTL")
            sections = record.get("Sections", {})
            unique_nu_ptls.add(nu_ptl)

            for sec_id in self.sections:
                records_in_section = sections.get(sec_id, [])
                section_record_count[sec_id] += len(records_in_section)

        summary = {
            "": {"Total": len(unique_nu_ptls)}
        }

        for sec in self.sections:
            summary[sec] = {"Total": section_record_count[sec]}

        return summary


    def create_summary_table(self, summary):
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        for col, header in enumerate(self.headers):
            label = ctk.CTkLabel(self.table_frame, text=header, font=ctk.CTkFont(weight="bold"))
            label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")

        for row_idx, row_label in enumerate(self.rows):
            for col_idx, header in enumerate(self.headers):
                if col_idx == 0:
                    text = "Total" if row_label == "" else row_label
                else:
                    text = summary.get(row_label, {}).get(header, "")
                label = ctk.CTkLabel(self.table_frame, text=text)
                label.grid(row=row_idx + 1, column=col_idx, padx=5, pady=3, sticky="nsew")

        for i in range(len(self.headers)):
            self.table_frame.grid_columnconfigure(i, weight=1)

        

if __name__ == "__main__":
    app = CTOSReportApp()
    app.mainloop()
