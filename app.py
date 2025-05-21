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
        self.geometry("900x600")

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

        # Add space between sidebar and main content
        self.sidebar_spacer = ctk.CTkFrame(self, width=20, fg_color="transparent")
        self.sidebar_spacer.pack(side="left", fill="y")

        # Main Content Area
        self.main_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_frame.pack(side="right", fill="both", expand=True)

        # Initialize Views
        self.xml_format_view = XMLFormatView(self.main_frame, self)
        self.ctos_report_view = CTOSReportView(self.main_frame, self)

        # Show Default View
        self.show_xml_format()
        
    def show_progress_popup(self, title="Processing...", message="Please wait..."):
        self.progress_popup = ctk.CTkToplevel(self)
        self.progress_popup.title(title)
        self.progress_popup.geometry("300x100")
        self.progress_popup.resizable(False, False)
        self.progress_popup.grab_set()  # Block interaction with main window

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

    def show_xml_format(self):
        self.xml_format_view.pack(fill="both", expand=True)
        self.ctos_report_view.pack_forget()

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

            # Get the NU_PTL value and set it in the search entry
            nu_ptl = current_row.get("NU_PTL", "")
            self.search_var.set(str(nu_ptl))

            # Parse the XML data in the "XML" column
            xml_data = current_row.get("XML", "")
            if pd.isna(xml_data) or not xml_data.strip():
                xml_data = "<No XML Data>"

            # Clear the Treeview
            self.tree.delete(*self.tree.get_children())

            # Parse the XML and display it in the Treeview
            try:
                dom = xml.dom.minidom.parseString(xml_data)
                root = dom.documentElement

                # Set up columns in the Treeview
                self.tree["columns"] = ["Field", "Value"]
                self.tree.heading("#0", text="")  # Hide the default column
                self.tree.column("#0", width=0, stretch=False)
                self.tree.heading("Field", text="Field")
                self.tree.heading("Value", text="Value")
                self.tree.column("Field", anchor="center", width=300)
                self.tree.column("Value", anchor="center", width=400)

                # Insert rows into the Treeview
                self.parse_xml_to_treeview(root, "")
            except Exception as e:
                self.tree.insert("", "end", values=["Error", str(e)])


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

            converted_rows = []
            total = len(self.filtered_data)
            self.after(0, self.update_status, "Converting...")
            self.after(0, lambda: self.error_textbox.delete("1.0", "end"))
            self.after(0, self.progress_bar.set, 0)
            

            for index, (_, row) in enumerate(self.filtered_data.iterrows()):
                if not self.is_converting:
                    break  # Stop if conversion canceled or UI closed

                nu_ptl = row.get("NU_PTL", f"Row{index}")
                xml_data = clean_malformed_xml(row.get("XML", ""))

                if pd.isna(xml_data) or not str(xml_data).strip():
                    continue

                try:
                    dom = xml.dom.minidom.parseString(xml_data)
                    row_data = {"NU_PTL": nu_ptl}

                    def extract_fields(node):
                        for child in node.childNodes:
                            if child.nodeType == xml.dom.minidom.Node.ELEMENT_NODE:
                                tag = child.tagName
                                if tag == "field_sum" and child.hasAttribute("name"):
                                    field = child.getAttribute("name").strip()
                                    value = (child.firstChild.nodeValue.strip()
                                            if child.firstChild and child.firstChild.nodeValue else " - ")
                                    row_data[field] = value
                                elif tag == "data":
                                    caption = child.getAttribute("caption") or child.getAttribute("name")
                                    field = caption.strip() if caption else "Unnamed_Field"
                                    value = (child.firstChild.nodeValue.strip()
                                            if child.firstChild and child.firstChild.nodeValue else " - ")
                                    row_data[field] = value
                                extract_fields(child)

                    extract_fields(dom.documentElement)
                    converted_rows.append(row_data)

                except Exception as e:
                    msg = f"Error parsing XML for NU_PTL {nu_ptl}: {str(e)}"
                    self.after(0, self.append_error, msg)
                    continue

                if index % 10 == 0 or index + 1 == total:
                    progress = (index + 1) / total
                    self.after(0, self.update_progress, progress, index + 1, total)

            if not converted_rows:
                self.after(0, self.update_status, "No valid XML data to export.")
                return

            df = pd.DataFrame(converted_rows)
            df.columns = df.columns.map(lambda x: str(x).strip().replace("\x00", "")[:255])
            df = df.applymap(lambda x: str(x).strip().replace("\x00", "") if pd.notnull(x) else "")
            df.reset_index(drop=True, inplace=True)


            max_rows_per_sheet = 100000
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                total_rows = df.shape[0]
                sheets_needed = (total_rows // max_rows_per_sheet) + 1

                for i in range(sheets_needed):
                    start_row = i * max_rows_per_sheet
                    end_row = start_row + max_rows_per_sheet
                    sheet_df = df.iloc[start_row:end_row]
        
                    if not sheet_df.empty:
                        sheet_name = f"part_{i + 1}"
                        sheet_df.to_excel(writer, sheet_name=sheet_name)


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

        # Next Button
        self.next_button = ctk.CTkButton(control_frame, text="Next", command=self.go_to_next)
        self.next_button.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # --- XML Display ---
        self.xml_display = ctk.CTkTextbox(self, width=600, height=300)
        self.xml_display.pack(pady=10, fill="both", expand=True)

        # Refresh Button
        self.refresh_button = ctk.CTkButton(self, text="Refresh Data", command=self.refresh_data)
        self.refresh_button.pack(pady=10)

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


if __name__ == "__main__":
    app = CTOSReportApp()
    app.mainloop()