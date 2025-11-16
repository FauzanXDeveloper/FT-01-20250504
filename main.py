import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pdfplumber
from pathlib import Path
import re
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import io

class PDFtoExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("PDF to Excel Converter - IRISS Report")
        self.geometry("1400x900")
        
        # Set theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Variables
        self.pdf_path = None
        self.pdf_paths = []  # For batch processing
        self.extracted_sections = {}
        self.file_previews = {}  # Store previews per file: {filename: DataFrame}
        self.pdf_document = None
        self.current_page = 0
        self.pdf_x_offset = 0  # X offset for centered PDF
        self.pdf_y_offset = 0  # Y offset for PDF position
        
        # Selection rectangle variables
        self.selection_start = None
        self.selection_rect = None
        self.selecting = False
        self.saved_selections = {}  # Store selections per page: {page_num: [(x1, y1, x2, y2), ...]}
        self.selection_rectangles = []  # Visual rectangles on canvas
        self.temp_selections = {}  # Temporary selections before saving: {page_num: [(bbox, rect_id), ...]}
        self.selection_history = []  # History for undo: [(page_num, bbox, rect_id), ...]
        
        # Create GUI
        self.create_widgets()
    
    def create_widgets(self):
        # Title Label
        title_label = ctk.CTkLabel(
            self,
            text="PDF to Excel Converter - IRISS Credit Report",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=20)
        
        # Button Frame
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=10, padx=20, fill="x")
        
        # Select PDF Button
        self.select_btn = ctk.CTkButton(
            button_frame,
            text="📁 Select PDF File",
            command=self.select_pdf,
            width=150,
            height=40,
            font=ctk.CTkFont(size=14)
        )
        self.select_btn.pack(side="left", padx=10, pady=10)
        
        # Select Multiple PDFs Button
        self.select_multiple_btn = ctk.CTkButton(
            button_frame,
            text="📁 Select Multiple PDFs",
            command=self.select_multiple_pdfs,
            width=150,
            height=40,
            font=ctk.CTkFont(size=14),
            fg_color="#1f538d",
            hover_color="#14375e"
        )
        self.select_multiple_btn.pack(side="left", padx=10, pady=10)
        
        # Select Folder Button
        self.select_folder_btn = ctk.CTkButton(
            button_frame,
            text="📂 Select Folder",
            command=self.select_folder,
            width=150,
            height=40,
            font=ctk.CTkFont(size=14),
            fg_color="#1f538d",
            hover_color="#14375e"
        )
        self.select_folder_btn.pack(side="left", padx=10, pady=10)
        
        # Second button row
        button_frame2 = ctk.CTkFrame(self)
        button_frame2.pack(pady=10, padx=20, fill="x")
        
        # File path label
        self.file_label = ctk.CTkLabel(
            button_frame2,
            text="No file selected",
            font=ctk.CTkFont(size=12)
        )
        self.file_label.pack(side="left", padx=10, pady=10, fill="x", expand=True)
        
        # Export Button
        self.export_btn = ctk.CTkButton(
            button_frame2,
            text="📊 Export to Excel",
            command=self.export_to_excel,
            width=150,
            height=40,
            font=ctk.CTkFont(size=14),
            state="disabled",
            fg_color="green",
            hover_color="darkgreen"
        )
        self.export_btn.pack(side="left", padx=10, pady=10)
        
        # Batch Process Button
        self.batch_btn = ctk.CTkButton(
            button_frame2,
            text="⚡ Batch Process All",
            command=self.batch_process_pdfs,
            width=150,
            height=40,
            font=ctk.CTkFont(size=14),
            state="disabled",
            fg_color="#9333ea",
            hover_color="#7e22ce"
        )
        self.batch_btn.pack(side="left", padx=10, pady=10)
        
        # Create Tabbed Interface
        self.tab_view = ctk.CTkTabview(self, width=1350, height=650)
        self.tab_view.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Tab 1: PDF Viewer
        self.tab_view.add("📄 PDF Viewer")
        pdf_tab = self.tab_view.tab("📄 PDF Viewer")
        
        # PDF navigation frame
        pdf_nav_frame = ctk.CTkFrame(pdf_tab)
        pdf_nav_frame.pack(pady=5, padx=10, fill="x")
        
        self.prev_btn = ctk.CTkButton(
            pdf_nav_frame,
            text="◀ Previous",
            command=self.previous_page,
            width=100,
            state="disabled"
        )
        self.prev_btn.pack(side="left", padx=5)
        
        self.page_label = ctk.CTkLabel(
            pdf_nav_frame,
            text="No PDF loaded",
            font=ctk.CTkFont(size=12)
        )
        self.page_label.pack(side="left", padx=20, expand=True)
        
        self.next_btn = ctk.CTkButton(
            pdf_nav_frame,
            text="Next ▶",
            command=self.next_page,
            width=100,
            state="disabled"
        )
        self.next_btn.pack(side="left", padx=5)
        
        # Selection mode button
        self.select_table_btn = ctk.CTkButton(
            pdf_nav_frame,
            text="📐 Select Table Area",
            command=self.toggle_selection_mode,
            width=150,
            state="disabled",
            fg_color="purple",
            hover_color="darkviolet"
        )
        self.select_table_btn.pack(side="right", padx=5)
        
        # Save Selection button
        self.save_selection_btn = ctk.CTkButton(
            pdf_nav_frame,
            text="💾 Save Selections",
            command=self.save_current_page_selections,
            width=150,
            state="disabled",
            fg_color="#16a34a",
            hover_color="#15803d"
        )
        self.save_selection_btn.pack(side="right", padx=5)
        
        # Undo Selection button (was Clear Selections)
        self.clear_selections_btn = ctk.CTkButton(
            pdf_nav_frame,
            text="↶ Undo Selection",
            command=self.undo_last_selection,
            width=150,
            state="disabled",
            fg_color="#dc2626",
            hover_color="#991b1b"
        )
        self.clear_selections_btn.pack(side="right", padx=5)
        
        # Auto Extract All Tables button
        self.auto_extract_btn = ctk.CTkButton(
            pdf_nav_frame,
            text="⚡ Auto Extract All Tables",
            command=self.auto_extract_all_tables,
            width=180,
            state="disabled",
            fg_color="#ea580c",
            hover_color="#c2410c"
        )
        self.auto_extract_btn.pack(side="right", padx=5)
        
        # Apply to All Files button
        self.apply_all_btn = ctk.CTkButton(
            pdf_nav_frame,
            text="📋 Apply to All Files",
            command=self.apply_selections_to_all_files,
            width=150,
            state="disabled",
            fg_color="#0891b2",
            hover_color="#0e7490"
        )
        self.apply_all_btn.pack(side="right", padx=5)
        
        # PDF display canvas with scrollbar
        pdf_canvas_frame = ctk.CTkFrame(pdf_tab)
        pdf_canvas_frame.pack(pady=5, padx=10, fill="both", expand=True)
        
        # Create canvas for PDF display
        self.pdf_canvas = ctk.CTkCanvas(
            pdf_canvas_frame,
            bg="#2b2b2b",
            highlightthickness=0
        )
        self.pdf_scrollbar = ctk.CTkScrollbar(
            pdf_canvas_frame,
            command=self.pdf_canvas.yview
        )
        self.pdf_canvas.configure(yscrollcommand=self.pdf_scrollbar.set)
        
        self.pdf_scrollbar.pack(side="right", fill="y")
        self.pdf_canvas.pack(side="left", fill="both", expand=True)
        
        # Bind mouse events for table selection
        self.pdf_canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.pdf_canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.pdf_canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        
        # Bind mouse wheel for scrolling
        self.pdf_canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        
        # Tab 2: Excel Preview
        self.tab_view.add("📊 Excel Preview")
        excel_tab = self.tab_view.tab("📊 Excel Preview")
        
        # Excel sheet selector
        excel_nav_frame = ctk.CTkFrame(excel_tab)
        excel_nav_frame.pack(pady=5, padx=10, fill="x")
        
        ctk.CTkLabel(
            excel_nav_frame,
            text="Select File:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=5)
        
        self.sheet_selector = ctk.CTkComboBox(
            excel_nav_frame,
            values=["No data available"],
            command=self.on_file_selected,
            width=400,
            state="disabled"
        )
        self.sheet_selector.pack(side="left", padx=10)
        
        # Excel data display with scrollbar
        excel_display_frame = ctk.CTkFrame(excel_tab)
        excel_display_frame.pack(pady=5, padx=10, fill="both", expand=True)
        
        self.excel_display = ctk.CTkTextbox(
            excel_display_frame,
            width=1300,
            height=550,
            font=ctk.CTkFont(size=10, family="Consolas")
        )
        self.excel_display.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Status bar
        self.status_label = ctk.CTkLabel(
            self,
            text="Ready - Please select a PDF file to begin",
            font=ctk.CTkFont(size=11),
            anchor="w"
        )
        self.status_label.pack(side="bottom", fill="x", padx=20, pady=10)
    
    def select_pdf(self):
        """Open file dialog to select PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.pdf_path = file_path
            self.pdf_paths = []  # Clear batch list
            self.file_label.configure(text=Path(file_path).name)
            self.batch_btn.configure(state="disabled")
            self.apply_all_btn.configure(state="disabled")
            self.status_label.configure(text=f"Selected: {Path(file_path).name}")
            self.export_btn.configure(state="disabled")
            
            # Clear previous extractions
            self.extracted_sections = {}
            
            # Load PDF for viewing
            self.load_pdf_viewer(file_path)
    
    def select_multiple_pdfs(self):
        """Open file dialog to select multiple PDF files"""
        file_paths = filedialog.askopenfilenames(
            title="Select Multiple PDF Files",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if file_paths:
            self.pdf_paths = list(file_paths)
            self.pdf_path = self.pdf_paths[0]  # Load first PDF for viewing
            
            self.file_label.configure(text=f"{len(self.pdf_paths)} PDFs selected")
            self.batch_btn.configure(state="normal")  # Enable batch process
            self.apply_all_btn.configure(state="normal")  # Enable apply to all
            self.status_label.configure(text=f"Selected {len(self.pdf_paths)} PDFs for batch processing")
            self.export_btn.configure(state="disabled")
            
            # Clear previous extractions
            self.extracted_sections = {}
            self.file_previews = {}  # Clear previous previews
            
            # Initialize preview entries for all files
            for pdf_path in self.pdf_paths:
                file_name = Path(pdf_path).name
                # Create empty placeholder - will be filled when selections are applied
                self.file_previews[file_name] = pd.DataFrame()
            
            # Update preview dropdown to show all files
            self.update_excel_preview()
            
            # Load first PDF for viewing
            self.load_pdf_viewer(self.pdf_paths[0])
    
    def select_folder(self):
        """Open dialog to select folder containing PDF files"""
        folder_path = filedialog.askdirectory(
            title="Select Folder Containing PDF Files"
        )
        
        if folder_path:
            # Find all PDF files in the folder
            folder = Path(folder_path)
            pdf_files = list(folder.glob("*.pdf"))
            
            if pdf_files:
                self.pdf_paths = [str(f) for f in pdf_files]
                self.pdf_path = self.pdf_paths[0]  # Load first PDF for viewing
                
                self.file_label.configure(text=f"{len(self.pdf_paths)} PDFs found in folder")
                self.batch_btn.configure(state="normal")  # Enable batch process
                self.apply_all_btn.configure(state="normal")  # Enable apply to all
                self.status_label.configure(text=f"Found {len(self.pdf_paths)} PDFs in {folder.name}")
                self.export_btn.configure(state="disabled")
                
                # Clear previous extractions
                self.extracted_sections = {}
                self.file_previews = {}  # Clear previous previews
                
                # Initialize preview entries for all files
                for pdf_path in self.pdf_paths:
                    file_name = Path(pdf_path).name
                    # Create empty placeholder - will be filled when selections are applied
                    self.file_previews[file_name] = pd.DataFrame()
                
                # Update preview dropdown to show all files
                self.update_excel_preview()
                
                # Load first PDF for viewing
                self.load_pdf_viewer(self.pdf_paths[0])
            else:
                messagebox.showwarning("No PDFs Found", "No PDF files found in the selected folder.")
                self.status_label.configure(text="No PDFs found in folder")
    
    def is_header_footer(self, text):
        """Check if text is part of header or footer"""
        if not text or pd.isna(text):
            return True
        
        text_str = str(text).strip()
        text_lower = text_str.lower()
        
        # Exact footer patterns
        footer_keywords = [
            'commercial confidential',
            'experian information services (malaysia) sdn. bhd.',
            'is certified to iso/iec',
            'cert. no: ism',
            'notice: the information provided by experian',
            'we do not guarantee the accuracy',
            'while we have used our best endeavours',
            'the information furnished is strictly confidential',
            'experian shall not be liable',
            'customer service division at:',
            'suite 16.02, level 16',
            'centrepoint south mid valley city',
            'lingkaran syed putra',
            'kuala lumpur',
            '+60326151111',
            'page',
            'of 7'
        ]
        
        # Exact header patterns
        header_keywords = [
            'strictly confidential',
            'order id',
            'credittrack by experian',
            'order date',
            'effective date',
            'user name',
            'check/track by experian'
        ]
        
        # Check footer patterns
        for keyword in footer_keywords:
            if keyword in text_lower:
                return True
        
        # Check header patterns
        for keyword in header_keywords:
            if keyword in text_lower:
                return True
        
        # Check if text is mostly the disclaimer
        if len(text_str) > 200 and 'experian' in text_lower:
            return True
        
        # Check if it's just a page number pattern
        if re.match(r'^page\s+\d+\s+of\s+\d+$', text_lower):
            return True
        
        # Check for Order ID pattern
        if re.match(r'.*order\s+id\s*:\s*\d+.*', text_lower):
            return True
            
        return False
    
    def clean_table_data(self, df):
        """Remove header/footer rows from dataframe"""
        if df.empty:
            return df
        
        # Remove rows where ANY column contains header/footer text
        rows_to_keep = []
        for idx, row in df.iterrows():
            # Check if this row contains header/footer content
            is_bad_row = False
            for cell in row:
                if self.is_header_footer(cell):
                    is_bad_row = True
                    break
            
            if not is_bad_row:
                rows_to_keep.append(idx)
        
        df_cleaned = df.loc[rows_to_keep].reset_index(drop=True)
        
        # Remove columns that are all empty or None
        df_cleaned = df_cleaned.dropna(axis=1, how='all')
        
        return df_cleaned
    
    def clean_extracted_data(self):
        """Deep clean extracted data to remove all header/footer contamination"""
        if not self.extracted_sections:
            messagebox.showwarning("Warning", "No data to clean! Please extract data first.")
            return
        
        try:
            self.status_label.configure(text="Cleaning extracted data...")
            self.update()
            
            df = self.extracted_sections["Extracted Data"]
            original_rows = len(df)
            
            # Apply thorough cleaning
            df_cleaned = self.clean_table_data(df)
            
            # Update the section with cleaned data
            self.extracted_sections["Extracted Data"] = df_cleaned
            
            rows_removed = original_rows - len(df_cleaned)
            
            self.status_label.configure(text=f"✓ Cleaned! Removed {rows_removed} rows")
            
            # Update Excel preview
            self.update_excel_preview()
            
            messagebox.showinfo("Success", f"Data cleaned successfully!\n\nRows removed: {rows_removed}\nTotal rows: {len(df_cleaned)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clean data:\n{str(e)}")
            self.status_label.configure(text="✗ Cleaning failed")
            import traceback
            traceback.print_exc()
    
    def load_pdf_viewer(self, pdf_path):
        """Load PDF into the PDF viewer tab"""
        try:
            if self.pdf_document:
                self.pdf_document.close()
            
            self.pdf_document = fitz.open(pdf_path)
            self.current_page = 0
            
            # Enable navigation buttons
            self.prev_btn.configure(state="normal")
            self.next_btn.configure(state="normal")
            self.select_table_btn.configure(state="normal")
            self.save_selection_btn.configure(state="normal")
            self.clear_selections_btn.configure(state="normal")
            self.auto_extract_btn.configure(state="normal")
            
            # Enable apply button if multiple files selected
            if self.pdf_paths and len(self.pdf_paths) > 1:
                self.apply_all_btn.configure(state="normal")
            
            # Display first page
            self.display_pdf_page()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF viewer:\n{str(e)}")
    
    def display_pdf_page(self):
        """Display the current PDF page"""
        if not self.pdf_document:
            return
        
        try:
            page = self.pdf_document[self.current_page]
            
            # Update page label
            self.page_label.configure(
                text=f"Page {self.current_page + 1} of {len(self.pdf_document)}"
            )
            
            # Render page to image
            zoom = 1.5  # Zoom factor for better quality
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            img_data = pix.tobytes("ppm")
            img = Image.open(io.BytesIO(img_data))
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(img)
            
            # Clear canvas
            self.pdf_canvas.delete("all")
            
            # Get canvas size
            canvas_width = self.pdf_canvas.winfo_width()
            canvas_height = self.pdf_canvas.winfo_height()
            
            # Calculate center position
            img_width, img_height = img.size
            x_center = max(0, (canvas_width - img_width) // 2) if canvas_width > img_width else 0
            y_center = 0  # Keep top aligned for scrolling
            
            # Store offsets for coordinate conversion
            self.pdf_x_offset = x_center
            self.pdf_y_offset = y_center
            
            # Display image centered horizontally
            self.pdf_canvas.create_image(x_center, y_center, anchor="nw", image=photo)
            self.pdf_canvas.image = photo  # Keep a reference
            
            # Update scroll region
            self.pdf_canvas.configure(scrollregion=self.pdf_canvas.bbox("all"))
            
            # Redraw saved selections
            self.redraw_selections()
            
        except Exception as e:
            print(f"Error displaying PDF page: {e}")
    
    def previous_page(self):
        """Go to previous PDF page"""
        if self.current_page > 0:
            self.current_page -= 1
            self.display_pdf_page()
    
    def next_page(self):
        """Go to next PDF page"""
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.display_pdf_page()
    
    def on_mouse_wheel(self, event):
        """Handle mouse wheel scrolling"""
        # Scroll the canvas vertically
        self.pdf_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def toggle_selection_mode(self):
        """Toggle table selection mode"""
        self.selecting = not self.selecting
        if self.selecting:
            self.select_table_btn.configure(text="✓ Selection Mode ON", fg_color="green")
            self.status_label.configure(text="📐 Draw a rectangle around the table you want to extract")
        else:
            self.select_table_btn.configure(text="📐 Select Table Area", fg_color="purple")
            self.status_label.configure(text="Selection mode disabled")
    
    def on_mouse_down(self, event):
        """Start drawing selection rectangle"""
        if not self.selecting:
            return
        
        # Convert event coordinates to canvas coordinates (account for scroll)
        canvas_x = self.pdf_canvas.canvasx(event.x)
        canvas_y = self.pdf_canvas.canvasy(event.y)
        self.selection_start = (canvas_x, canvas_y)
        # Delete previous rectangle if exists
        if self.selection_rect:
            self.pdf_canvas.delete(self.selection_rect)
            self.selection_rect = None
    
    def on_mouse_drag(self, event):
        """Draw selection rectangle as user drags"""
        if not self.selecting or not self.selection_start:
            return
        
        # Delete previous rectangle
        if self.selection_rect:
            self.pdf_canvas.delete(self.selection_rect)
        
        # Draw new rectangle
        x1, y1 = self.selection_start
        # Convert event coordinates to canvas coordinates (account for scroll)
        x2 = self.pdf_canvas.canvasx(event.x)
        y2 = self.pdf_canvas.canvasy(event.y)
        
        self.selection_rect = self.pdf_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="red",
            width=3,
            dash=(5, 5)
        )
    
    def on_mouse_up(self, event):
        """Finish selection and store temporarily until saved"""
        if not self.selecting or not self.selection_start:
            return
        
        x1, y1 = self.selection_start
        # Convert event coordinates to canvas coordinates (account for scroll)
        x2 = self.pdf_canvas.canvasx(event.x)
        y2 = self.pdf_canvas.canvasy(event.y)
        
        # Ensure x1 < x2 and y1 < y2
        if x1 > x2:
            x1, x2 = x2, x1
        if y1 > y2:
            y1, y2 = y2, y1
        
        # Convert canvas coordinates to PDF coordinates
        # First subtract the PDF offset (centering)
        pdf_x1 = x1 - self.pdf_x_offset
        pdf_y1 = y1 - self.pdf_y_offset
        pdf_x2 = x2 - self.pdf_x_offset
        pdf_y2 = y2 - self.pdf_y_offset
        
        # Then account for zoom factor
        zoom = 1.5
        
        # Get bounding box in PDF coordinates
        bbox = (pdf_x1/zoom, pdf_y1/zoom, pdf_x2/zoom, pdf_y2/zoom)
        
        # Draw selection rectangle (orange for temporary)
        perm_rect = self.pdf_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="orange",
            width=2
        )
        
        # Store in temporary selections
        if self.current_page not in self.temp_selections:
            self.temp_selections[self.current_page] = []
        self.temp_selections[self.current_page].append((bbox, perm_rect))
        
        # Add to history for undo
        self.selection_history.append((self.current_page, bbox, perm_rect))
        
        # Clear temporary selection
        if self.selection_rect:
            self.pdf_canvas.delete(self.selection_rect)
            self.selection_rect = None
        
        self.selection_start = None
        
        # Count temp selections on current page
        temp_count = len(self.temp_selections.get(self.current_page, []))
        self.status_label.configure(text=f"✓ Selection marked! Page {self.current_page + 1}: {temp_count} selection(s) (Click 💾 Save to finalize)")
    
    def save_current_page_selections(self):
        """Save all temporary selections for current page and extract data"""
        if self.current_page not in self.temp_selections or not self.temp_selections[self.current_page]:
            messagebox.showinfo("Info", "No selections to save on this page!")
            return
        
        try:
            # Get current viewing file (could be different from batch files)
            current_file = self.pdf_path
            
            # Extract data from all temp selections on this page
            with pdfplumber.open(current_file) as pdf:
                page = pdf.pages[self.current_page]
                file_name = Path(current_file).name
                
                for bbox, rect_id in self.temp_selections[self.current_page]:
                    # Extract table
                    table = page.crop(bbox).extract_table()
                    
                    if table:
                        new_df = pd.DataFrame(table)
                        
                        # Append to file preview
                        if file_name in self.file_previews:
                            existing_df = self.file_previews[file_name]
                            # Add separator
                            num_cols = max(len(existing_df.columns), len(new_df.columns))
                            separator = pd.DataFrame([[None] * num_cols] * 2)
                            self.file_previews[file_name] = pd.concat([existing_df, separator, new_df], ignore_index=True)
                        else:
                            self.file_previews[file_name] = new_df
                    
                    # Change rectangle color to blue (saved)
                    self.pdf_canvas.itemconfig(rect_id, outline="blue")
                    
                    # Move to saved selections
                    if self.current_page not in self.saved_selections:
                        self.saved_selections[self.current_page] = []
                    self.saved_selections[self.current_page].append(bbox)
            
            # Clear temp selections for this page
            saved_count = len(self.temp_selections[self.current_page])
            self.temp_selections[self.current_page] = []
            
            # Update preview
            self.update_excel_preview()
            self.export_btn.configure(state="normal")
            
            total_saved = sum(len(sels) for sels in self.saved_selections.values())
            self.status_label.configure(text=f"✓ Saved {saved_count} selection(s) from page {self.current_page + 1}! Total saved: {total_saved}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save selections:\n{str(e)}")
            print(f"Save error: {e}")
    
    def undo_last_selection(self):
        """Undo the most recent selection (remove from temp selections)"""
        if not self.selection_history:
            messagebox.showinfo("Info", "No selections to undo!")
            return
        
        # Get last selection from history
        page_num, bbox, rect_id = self.selection_history.pop()
        
        # Remove from temp selections
        if page_num in self.temp_selections:
            self.temp_selections[page_num] = [(b, r) for b, r in self.temp_selections[page_num] if r != rect_id]
            if not self.temp_selections[page_num]:
                del self.temp_selections[page_num]
        
        # Delete rectangle from canvas
        self.pdf_canvas.delete(rect_id)
        
        # Update status
        remaining = sum(len(sels) for sels in self.temp_selections.values())
        self.status_label.configure(text=f"↶ Undone! {remaining} unsaved selection(s) remaining")
    
    def auto_extract_all_tables(self):
        """Automatically extract all tables from all pages with header/footer removal"""
        if not self.pdf_document:
            messagebox.showinfo("Info", "No PDF loaded!")
            return
        
        try:
            current_file = self.pdf_path
            file_name = Path(current_file).name
            
            self.status_label.configure(text=f"Auto-extracting tables from {file_name}...")
            self.update()
            
            with pdfplumber.open(current_file) as pdf:
                all_tables = []
                
                for page_num, page in enumerate(pdf.pages):
                    # Get page dimensions
                    page_height = page.height
                    
                    # Define crop area (remove top 80px and bottom 80px)
                    # pdfplumber uses points (72 points = 1 inch)
                    header_crop = 80  # Top 80px to remove
                    footer_crop = 80  # Bottom 80px to remove
                    
                    # Crop the page to remove header and footer
                    cropped_page = page.crop((0, header_crop, page.width, page_height - footer_crop))
                    
                    # Extract all tables from cropped page
                    tables = cropped_page.extract_tables()
                    
                    if tables:
                        for table in tables:
                            if table and len(table) > 0:
                                all_tables.append(table)
                
                if all_tables:
                    # Combine all tables with separators
                    combined_rows = []
                    for table_idx, table in enumerate(all_tables):
                        combined_rows.extend(table)
                        # Add separator between tables
                        if table_idx < len(all_tables) - 1:
                            num_cols = len(table[0]) if table else 1
                            combined_rows.extend([[None] * num_cols] * 2)
                    
                    # Create DataFrame
                    df = pd.DataFrame(combined_rows)
                    
                    # Store in file previews
                    self.file_previews[file_name] = df
                    
                    # Update preview
                    self.update_excel_preview()
                    self.export_btn.configure(state="normal")
                    
                    self.status_label.configure(text=f"✓ Auto-extracted {len(all_tables)} tables from {len(pdf.pages)} pages!")
                    messagebox.showinfo("Success", f"Auto-extracted {len(all_tables)} tables!\n\nHeader/footer areas (80px) removed automatically.\n\nCheck Excel Preview tab.")
                else:
                    messagebox.showwarning("No Tables", "No tables found in the PDF.")
                    self.status_label.configure(text="⚠ No tables found")
        
        except Exception as e:
            messagebox.showerror("Error", f"Auto-extraction failed:\n{str(e)}")
            self.status_label.configure(text="✗ Auto-extraction failed")
            import traceback
            traceback.print_exc()
    
    def apply_selections_to_all_files(self):
        """Apply saved selections to all imported PDF files with header/footer removal"""
        if not self.pdf_paths:
            messagebox.showinfo("Info", "No multiple files selected!")
            return
        
        if not self.saved_selections:
            messagebox.showerror("Error", "No selections saved! Please mark and save selections first.")
            return
        
        try:
            self.status_label.configure(text=f"Applying selections to {len(self.pdf_paths)} files...")
            self.update()
            
            processed_count = 0
            header_crop = 80  # Top 80px to remove
            footer_crop = 80  # Bottom 80px to remove
            
            for pdf_path in self.pdf_paths:
                file_name = Path(pdf_path).name
                
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        file_tables = []
                        
                        # Apply selections from each saved page
                        for page_num in sorted(self.saved_selections.keys()):
                            if page_num < len(pdf.pages):
                                page = pdf.pages[page_num]
                                page_height = page.height
                                
                                # Crop page to remove header/footer
                                cropped_page = page.crop((0, header_crop, page.width, page_height - footer_crop))
                                
                                for bbox in self.saved_selections[page_num]:
                                    try:
                                        # Adjust bbox coordinates for the cropped page
                                        adjusted_bbox = (
                                            bbox[0],
                                            max(0, bbox[1] - header_crop),  # Adjust Y coordinate
                                            bbox[2],
                                            bbox[3] - header_crop
                                        )
                                        
                                        # Only extract if bbox is within cropped area
                                        if adjusted_bbox[1] >= 0 and adjusted_bbox[3] <= (page_height - header_crop - footer_crop):
                                            table = cropped_page.crop(adjusted_bbox).extract_table()
                                            if table:
                                                file_tables.append(pd.DataFrame(table))
                                    except:
                                        continue
                        
                        # Combine all tables for this file
                        if file_tables:
                            combined_df = file_tables[0]
                            for df in file_tables[1:]:
                                # Add separator
                                num_cols = max(len(combined_df.columns), len(df.columns))
                                separator = pd.DataFrame([[None] * num_cols] * 2)
                                combined_df = pd.concat([combined_df, separator, df], ignore_index=True)
                            
                            self.file_previews[file_name] = combined_df
                            processed_count += 1
                
                except Exception as e:
                    print(f"Error processing {file_name}: {e}")
                    continue
            
            # Update preview
            self.update_excel_preview()
            self.export_btn.configure(state="normal")
            
            self.status_label.configure(text=f"✓ Applied to {processed_count} files with header/footer removed!")
            messagebox.showinfo("Success", f"Successfully applied selections to {processed_count} files!\n\nHeader/footer (80px) removed automatically.\n\nCheck the Excel Preview tab.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to apply selections:\n{str(e)}")
            print(f"Apply error: {e}")
    
    def redraw_selections(self):
        """Redraw saved and temporary selections for the current page only"""
        # Note: We don't clear rectangles here as they persist across page changes
        # Only redraw if switching pages
        
        zoom = 1.5
        
        # Redraw saved selections (blue)
        if self.current_page in self.saved_selections:
            for bbox in self.saved_selections[self.current_page]:
                x1, y1, x2, y2 = bbox
                # Convert PDF coords to canvas coords and add offset
                canvas_x1 = x1 * zoom + self.pdf_x_offset
                canvas_y1 = y1 * zoom + self.pdf_y_offset
                canvas_x2 = x2 * zoom + self.pdf_x_offset
                canvas_y2 = y2 * zoom + self.pdf_y_offset
                
                self.pdf_canvas.create_rectangle(
                    canvas_x1, canvas_y1, canvas_x2, canvas_y2,
                    outline="blue",
                    width=2
                )
        
        # Redraw temp selections (orange)
        if self.current_page in self.temp_selections:
            for bbox, old_rect_id in self.temp_selections[self.current_page]:
                x1, y1, x2, y2 = bbox
                # Convert PDF coords to canvas coords and add offset
                canvas_x1 = x1 * zoom + self.pdf_x_offset
                canvas_y1 = y1 * zoom + self.pdf_y_offset
                canvas_x2 = x2 * zoom + self.pdf_x_offset
                canvas_y2 = y2 * zoom + self.pdf_y_offset
                
                self.pdf_canvas.create_rectangle(
                    canvas_x1, canvas_y1, canvas_x2, canvas_y2,
                    outline="orange",
                    width=2
                )
    
    def extract_selected_table(self, bbox):
        """Extract table from the selected bounding box"""
        if not self.pdf_document:
            return
        
        try:
            self.status_label.configure(text="Extracting table from selected area...")
            self.update()
            
            with pdfplumber.open(self.pdf_path) as pdf:
                page = pdf.pages[self.current_page]
                
                # Extract table from the selected area
                table = page.crop(bbox).extract_table()
                
                if table:
                    # Convert to DataFrame
                    new_df = pd.DataFrame(table)
                    
                    # Append to single sheet named "Extracted Data"
                    if "Extracted Data" in self.extracted_sections:
                        # Append to existing data with blank rows as separator
                        existing_df = self.extracted_sections["Extracted Data"]
                        
                        # Create separator (2 blank rows)
                        num_cols = max(len(existing_df.columns), len(new_df.columns))
                        separator = pd.DataFrame([[None] * num_cols] * 2)
                        
                        # Concatenate: existing data + separator + new data
                        self.extracted_sections["Extracted Data"] = pd.concat([existing_df, separator, new_df], ignore_index=True)
                    else:
                        # Create new sheet
                        self.extracted_sections["Extracted Data"] = new_df
                    
                    # Enable export buttons
                    self.export_btn.configure(state="normal")
                    
                    # Update displays
                    self.update_excel_preview()
                    
                    # Show success message
                    total_rows = len(self.extracted_sections["Extracted Data"])
                    self.status_label.configure(text=f"✓ Table extracted! Total: {total_rows} rows")
                    
                    # Turn off selection mode
                    self.selecting = False
                    self.select_table_btn.configure(text="📐 Select Table Area", fg_color="purple")
                    
                    # Clear selection rectangle
                    if self.selection_rect:
                        self.pdf_canvas.delete(self.selection_rect)
                        self.selection_rect = None
                else:
                    messagebox.showwarning("No Table", "No table found in selected area. Try selecting a different area.")
                    self.status_label.configure(text="⚠ No table found in selected area")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract table:\n{str(e)}")
            self.status_label.configure(text="✗ Extraction failed")
            import traceback
            traceback.print_exc()
    
    def update_excel_preview(self):
        """Update Excel preview tab with extracted data"""
        if not self.file_previews:
            self.sheet_selector.configure(values=["No data available"], state="disabled")
            self.excel_display.delete("1.0", "end")
            self.excel_display.insert("1.0", "No data extracted yet. Mark areas on PDF to preview.")
            return
        
        # Show list of files
        file_names = list(self.file_previews.keys())
        self.sheet_selector.configure(values=file_names, state="readonly")
        self.sheet_selector.set(file_names[0])
        
        # Display first file
        self.display_file_preview(file_names[0])
    
    def on_file_selected(self, file_name):
        """Handle file selection change"""
        self.display_file_preview(file_name)
    
    def display_file_preview(self, file_name):
        """Display selected file preview data"""
        if file_name not in self.file_previews:
            return
        
        df = self.file_previews[file_name]
        
        self.excel_display.delete("1.0", "end")
        self.excel_display.insert("1.0", f"File: {file_name}\n")
        self.excel_display.insert("end", f"{'='*100}\n")
        
        # Check if dataframe is empty
        if df.empty:
            self.excel_display.insert("end", "No data extracted yet.\n\n")
            self.excel_display.insert("end", "To extract data:\n")
            self.excel_display.insert("end", "1. Mark selection areas on PDF\n")
            self.excel_display.insert("end", "2. Click 💾 Save Selections\n")
            self.excel_display.insert("end", "3. Click 📋 Apply to All Files to process all PDFs\n")
        else:
            self.excel_display.insert("end", f"Rows: {len(df)} | Columns: {len(df.columns)}\n")
            self.excel_display.insert("end", f"{'='*100}\n\n")
            # Display full dataframe
            self.excel_display.insert("end", df.to_string())
    
    def extract_data(self):
        """Extract ALL tables from PDF directly - simple and accurate"""
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first!")
            return
        
        try:
            self.status_label.configure(text="Extracting data from PDF...")
            self.update()
            
            self.extracted_sections = {}
            table_counter = 1
            
            with pdfplumber.open(self.pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    # Extract ALL tables from this page
                    tables = page.extract_tables()
                    
                    if tables:
                        for table in tables:
                            if not table or len(table) < 1:
                                continue
                            
                            # Convert to DataFrame - keep everything as-is
                            df = pd.DataFrame(table)
                            
                            if not df.empty:
                                # Simple naming: Table_1, Table_2, etc.
                                section_name = f"Table_{table_counter}"
                                
                                # Try to identify section from first row
                                try:
                                    first_row = ' '.join([str(cell) for cell in df.iloc[0] if pd.notna(cell)])
                                    
                                    if "SECTION 1" in first_row.upper():
                                        section_name = "Section 1"
                                    elif "SECTION 2" in first_row.upper():
                                        section_name = "Section 2"
                                    elif "SECTION 3" in first_row.upper():
                                        section_name = "Section 3"
                                    elif "SECTION 4" in first_row.upper():
                                        section_name = "Section 4"
                                    elif "SECTION 5" in first_row.upper():
                                        section_name = "Section 5"
                                except:
                                    pass
                                
                                self.extracted_sections[section_name] = df
                                table_counter += 1
            
            # Display extracted sections
            self.text_display.delete("1.0", "end")
            self.text_display.insert("1.0", f"✓ Extraction Complete!\n")
            self.text_display.insert("end", f"{'='*100}\n\n")
            self.text_display.insert("end", f"Total Sections Found: {len(self.extracted_sections)}\n\n")
            
            for idx, (section_name, df) in enumerate(self.extracted_sections.items(), 1):
                self.text_display.insert("end", f"\n{'='*100}\n")
                self.text_display.insert("end", f"SECTION {idx}: {section_name}\n")
                self.text_display.insert("end", f"{'='*100}\n")
                self.text_display.insert("end", f"Rows: {len(df)} | Columns: {len(df.columns)}\n")
                self.text_display.insert("end", f"Columns: {list(df.columns)}\n\n")
                
                # Show first 10 rows
                preview = df.head(10).to_string()
                self.text_display.insert("end", preview + "\n")
                
                if len(df) > 10:
                    self.text_display.insert("end", f"\n... and {len(df) - 10} more rows\n")
            
            # Enable export button
            if self.extracted_sections:
                self.export_btn.configure(state="normal")
                self.clean_btn.configure(state="normal")
                self.status_label.configure(text=f"✓ Extracted {len(self.extracted_sections)} sections successfully!")
                
                # Update Excel preview
                self.update_excel_preview()
                
                messagebox.showinfo("Success", f"Successfully extracted {len(self.extracted_sections)} sections from the PDF!")
            else:
                self.status_label.configure(text="⚠ No sections found in PDF")
                messagebox.showwarning("Warning", "No recognizable sections found in the PDF!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract data:\n{str(e)}")
            self.status_label.configure(text="✗ Extraction failed")
            self.text_display.delete("1.0", "end")
            self.text_display.insert("1.0", f"Error: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def batch_process_pdfs(self):
        """Process multiple PDFs using saved selection areas with header/footer removal"""
        if not self.pdf_paths:
            messagebox.showerror("Error", "No PDF files selected for batch processing!")
            return
        
        if not self.saved_selections:
            messagebox.showerror("Error", "No selection areas defined! Please mark selection areas first.")
            return
        
        # Ask user to select output folder
        output_folder = filedialog.askdirectory(
            title="Select Output Folder for Excel Files"
        )
        
        if not output_folder:
            return
        
        output_path = Path(output_folder)
        success_count = 0
        failed_files = []
        header_crop = 80  # Top 80px to remove
        footer_crop = 80  # Bottom 80px to remove
        
        try:
            total_files = len(self.pdf_paths)
            
            for idx, pdf_path in enumerate(self.pdf_paths, 1):
                try:
                    pdf_name = Path(pdf_path).stem
                    self.status_label.configure(text=f"Processing {idx}/{total_files}: {pdf_name}...")
                    self.update()
                    
                    # Extract data from this PDF using saved selections
                    with pdfplumber.open(pdf_path) as pdf:
                        all_extracted_tables = []
                        
                        # Sort pages to maintain order
                        sorted_pages = sorted(self.saved_selections.keys())
                        
                        # Apply selections from each page in order
                        for page_num in sorted_pages:
                            if page_num < len(pdf.pages):
                                page = pdf.pages[page_num]
                                page_height = page.height
                                
                                # Crop page to remove header/footer
                                cropped_page = page.crop((0, header_crop, page.width, page_height - footer_crop))
                                
                                for bbox in self.saved_selections[page_num]:
                                    try:
                                        # Adjust bbox coordinates for the cropped page
                                        adjusted_bbox = (
                                            bbox[0],
                                            max(0, bbox[1] - header_crop),
                                            bbox[2],
                                            bbox[3] - header_crop
                                        )
                                        
                                        # Only extract if bbox is within cropped area
                                        if adjusted_bbox[1] >= 0 and adjusted_bbox[3] <= (page_height - header_crop - footer_crop):
                                            table = cropped_page.crop(adjusted_bbox).extract_table()
                                            if table:
                                                all_extracted_tables.append(table)
                                    except:
                                        continue
                        
                        if all_extracted_tables:
                            # Combine all tables with separators
                            combined_rows = []
                            for table_idx, table in enumerate(all_extracted_tables):
                                if table:
                                    combined_rows.extend(table)
                                    # Add separator between tables
                                    if table_idx < len(all_extracted_tables) - 1:
                                        num_cols = len(table[0]) if table else 1
                                        combined_rows.extend([[None] * num_cols] * 2)
                            
                            df = pd.DataFrame(combined_rows)
                            
                            # Export to Excel
                            excel_path = output_path / f"{pdf_name}.xlsx"
                            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                                df.to_excel(writer, sheet_name="Extracted Data", index=False, header=False)
                            
                            success_count += 1
                        else:
                            failed_files.append(f"{pdf_name} (no tables found)")
                
                except Exception as e:
                    failed_files.append(f"{Path(pdf_path).name} (error: {str(e)})")
                    continue
            
            # Show summary
            total_selections = sum(len(sels) for sels in self.saved_selections.values())
            summary_msg = f"Batch Processing Complete!\n\n"
            summary_msg += f"Successfully processed: {success_count}/{total_files}\n"
            summary_msg += f"Used {total_selections} selection areas across {len(self.saved_selections)} pages\n"
            summary_msg += f"Header/footer (80px) removed automatically\n"
            summary_msg += f"Output folder: {output_folder}\n"
            
            if failed_files:
                summary_msg += f"\nFailed files ({len(failed_files)}): \n"
                summary_msg += "\n".join(failed_files[:5])
                if len(failed_files) > 5:
                    summary_msg += f"\n... and {len(failed_files) - 5} more"
            
            self.status_label.configure(text=f"✓ Batch complete: {success_count}/{total_files} successful")
            messagebox.showinfo("Batch Processing Complete", summary_msg)
            
        except Exception as e:
            messagebox.showerror("Batch Processing Error", f"An error occurred:\n{str(e)}")
            self.status_label.configure(text="✗ Batch processing failed")
            import traceback
            traceback.print_exc()
    
    def export_to_excel(self):
        """Export all extracted tables to Excel with borders and merged cells"""
        if not self.file_previews:
            messagebox.showerror("Error", "No data to export! Please mark areas first.")
            return
        
        try:
            # Open save dialog
            save_folder = filedialog.askdirectory(
                title="Select Output Folder for Excel Files"
            )
            
            if save_folder:
                self.status_label.configure(text="Exporting to Excel...")
                self.update()
                
                from openpyxl.styles import Border, Side, Alignment
                from openpyxl.utils import get_column_letter
                
                output_path = Path(save_folder)
                
                # Export each file preview
                for file_name, df in self.file_previews.items():
                    excel_name = Path(file_name).stem + ".xlsx"
                    excel_path = output_path / excel_name
                    
                    # Write to Excel
                    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name="Extracted Data", index=False, header=False)
                        
                        # Get worksheet
                        worksheet = writer.sheets["Extracted Data"]
                        
                        # Define border style
                        thin_border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
                        
                        # Apply borders and detect merged cells
                        for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=len(df)), 1):
                            for col_idx, cell in enumerate(row, 1):
                                # Apply border
                                cell.border = thin_border
                                
                                # Center align
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                
                                # Detect if cell should be merged (same value as previous cell in row)
                                if col_idx > 1:
                                    prev_cell = worksheet.cell(row_idx, col_idx - 1)
                                    if cell.value and prev_cell.value and str(cell.value).strip() == str(prev_cell.value).strip():
                                        # Try to merge with previous cell
                                        try:
                                            # Find the start of merge range
                                            merge_start = col_idx - 1
                                            while merge_start > 1:
                                                check_cell = worksheet.cell(row_idx, merge_start - 1)
                                                if check_cell.value and str(check_cell.value).strip() == str(cell.value).strip():
                                                    merge_start -= 1
                                                else:
                                                    break
                                            
                                            # Check if not already merged
                                            range_str = f"{get_column_letter(merge_start)}{row_idx}:{get_column_letter(col_idx)}{row_idx}"
                                            if range_str not in [str(m) for m in worksheet.merged_cells]:
                                                worksheet.merge_cells(range_str)
                                        except:
                                            pass
                        
                        # Auto-adjust column widths
                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = get_column_letter(column[0].column)
                            for cell in column:
                                try:
                                    if cell.value:
                                        max_length = max(max_length, len(str(cell.value)))
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 50)
                            worksheet.column_dimensions[column_letter].width = adjusted_width
                
                self.status_label.configure(text=f"✓ Exported {len(self.file_previews)} files")
                messagebox.showinfo("Success", f"Exported {len(self.file_previews)} files to:\n{save_folder}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data:\n{str(e)}")
            self.status_label.configure(text="✗ Export failed")
            import traceback
            traceback.print_exc()

def main():
    app = PDFtoExcelApp()
    app.mainloop()

if __name__ == "__main__":
    main()
