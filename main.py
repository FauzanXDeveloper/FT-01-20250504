import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pdfplumber
from pathlib import Path
import re
import fitz  # PyMuPDF
from PIL import Image, ImageTk, ImageFilter
import io
import threading
import time

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
        
        # Loading screen variables
        self.loading_frame = None
        self.loading_gif_frames = []
        self.loading_gif_label = None
        self.loading_progress_label = None
        self.loading_current_frame = 0
        self.loading_animation_job = None
        self.blur_overlay = None
        
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
        
        # File count label
        self.file_count_label = ctk.CTkLabel(
            button_frame,
            text="0 files",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.file_count_label.pack(side="left", padx=20, pady=10)
        
        # Export Button
        self.export_btn = ctk.CTkButton(
            button_frame,
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
            button_frame,
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
        
        # Create Tabbed Interfaceabel to prevent errors (but don't pack it)
        self.file_label = ctk.CTkLabel(self, text="")
        
        # Create Tabbed Interface
        self.tab_view = ctk.CTkTabview(self, width=1350, height=650)
        self.tab_view.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Tab 1: PDF Viewer
        self.tab_view.add("📄 PDF Viewer")
        pdf_tab = self.tab_view.tab("📄 PDF Viewer")
        
        # PDF navigation frame
        pdf_nav_frame = ctk.CTkFrame(pdf_tab)
        pdf_nav_frame.pack(pady=5, padx=10, fill="x")
        
        # Create a frame for centered buttonsent errors (but don't pack it)
        self.page_label = ctk.CTkLabel(pdf_nav_frame, text="")
        
        # Create a frame for centered buttons
        button_container = ctk.CTkFrame(pdf_nav_frame)
        button_container.pack(pady=5)
        
        # Selection mode button
        self.select_table_btn = ctk.CTkButton(
            button_container,
            text="📐 Select Table Area",
            command=self.toggle_selection_mode,
            width=150,
            state="disabled",
            fg_color="purple",
            hover_color="darkviolet"
        )
        self.select_table_btn.pack(side="left", padx=5)
        
        # Save Selection button
        self.save_selection_btn = ctk.CTkButton(
            button_container,
            text="💾 Save Selections",
            command=self.save_current_page_selections,
            width=150,
            state="disabled",
            fg_color="#16a34a",
            hover_color="#15803d"
        )
        self.save_selection_btn.pack(side="left", padx=5)
        
        # Undo Selection button (was Clear Selections)
        self.clear_selections_btn = ctk.CTkButton(
            button_container,
            text="↶ Undo Selection",
            command=self.undo_last_selection,
            width=150,
            state="disabled",
            fg_color="#dc2626",
            hover_color="#991b1b"
        )
        self.clear_selections_btn.pack(side="left", padx=5)
        
        # Auto Extract All Tables button
        self.auto_extract_btn = ctk.CTkButton(
            button_container,
            text="⚡ Auto Extract All Tables",
            command=self.auto_extract_all_tables,
            width=180,
            state="disabled",
            fg_color="#ea580c",
            hover_color="#c2410c"
        )
        self.auto_extract_btn.pack(side="left", padx=5)
        
        # Clear Preview button
        self.clear_preview_btn = ctk.CTkButton(
            button_container,
            text="🧹 Clear Preview",
            command=self.clear_all_previews,
            width=150,
            state="disabled",
            fg_color="#64748b",
            hover_color="#475569"
        )
        self.clear_preview_btn.pack(side="left", padx=5)
        
        # Apply to All Files button
        self.apply_all_btn = ctk.CTkButton(
            button_container,
            text="📋 Apply to All Files",
            command=self.apply_selections_to_all_files,
            width=150,
            state="disabled",
            fg_color="#0891b2",
            hover_color="#0e7490"
        )
        self.apply_all_btn.pack(side="left", padx=5)
        
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

        # Sticky nav buttons overlay (appear over canvas)
        # Placed in the same frame so they float above the canvas content
        self.sticky_prev = ctk.CTkButton(pdf_canvas_frame, text="◀ Previous", width=80, command=self.previous_page, state="disabled")
        self.sticky_prev.place(relx=0.02, rely=0.98, anchor="sw")
        
        # Center sticky button for page info
        self.sticky_center = ctk.CTkLabel(pdf_canvas_frame, text="No PDF loaded", width=120, 
                                         font=ctk.CTkFont(size=11), 
                                         fg_color="#2b2b2b", corner_radius=6)
        self.sticky_center.place(relx=0.5, rely=0.98, anchor="s")
        
        self.sticky_next = ctk.CTkButton(pdf_canvas_frame, text="Next ▶", width=80, command=self.next_page, state="disabled")
        self.sticky_next.place(relx=0.98, rely=0.98, anchor="se")
        
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
        
        # Create treeview with scrollbars for Excel-like display
        tree_scroll_y = ctk.CTkScrollbar(excel_display_frame, orientation="vertical")
        tree_scroll_y.pack(side="right", fill="y")
        
        tree_scroll_x = ctk.CTkScrollbar(excel_display_frame, orientation="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")
        
        # Use tkinter Treeview for table display
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview",
                        background="#2b2b2b",
                        foreground="white",
                        fieldbackground="#2b2b2b",
                        borderwidth=1,
                        relief="solid")
        style.configure("Treeview.Heading",
                        background="#1f538d",
                        foreground="white",
                        borderwidth=1)
        style.map("Treeview",
                  background=[("selected", "#144870")])
        
        self.excel_display = ttk.Treeview(
            excel_display_frame,
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
            show="tree headings",
            selectmode="extended"
        )
        self.excel_display.pack(pady=10, padx=10, fill="both", expand=True)
        
        tree_scroll_y.configure(command=self.excel_display.yview)
        tree_scroll_x.configure(command=self.excel_display.xview)
        
        # Status bar
        self.status_label = ctk.CTkLabel(
            self,
            text="Ready - Please select a PDF file to begin",
            font=ctk.CTkFont(size=11),
            anchor="w"
        )
        self.status_label.pack(side="bottom", fill="x", padx=20, pady=10)
        
        # Setup loading screen
        self.setup_loading_screen()
    
    def setup_loading_screen(self):
        """Setup the loading screen with GIF animation and blur effect"""
        try:
            # Load GIF frames
            gif_path = r"C:\Users\aaafauzan\Desktop\VS Code\Pdf to Excel\Picture\Loading.gif"
            gif_image = Image.open(gif_path)
            
            # Extract all frames from GIF
            self.loading_gif_frames = []
            try:
                while True:
                    # Resize frame to reasonable size
                    frame = gif_image.copy().resize((100, 100), Image.Resampling.LANCZOS)
                    # Convert to CTkImage to avoid warnings
                    ctk_frame = ctk.CTkImage(frame, size=(100, 100))
                    self.loading_gif_frames.append(ctk_frame)
                    gif_image.seek(len(self.loading_gif_frames))
            except EOFError:
                pass  # End of GIF frames
                
        except Exception as e:
            print(f"Warning: Could not load loading GIF: {e}")
            # Create a simple default loading image
            default_img = Image.new('RGBA', (100, 100), (100, 100, 100, 255))
            self.loading_gif_frames = [ctk.CTkImage(default_img, size=(100, 100))]
    
    def show_loading_screen(self, message="Loading..."):
        """Show loading screen with blur effect"""
        if self.loading_frame:
            return  # Already showing
            
        # Create blur overlay
        self.blur_overlay = ctk.CTkFrame(
            self,
            fg_color=("gray90", "gray10"),
            bg_color="transparent"
        )
        self.blur_overlay.place(x=0, y=0, relwidth=1, relheight=1)
        
        # Create loading frame (centered)
        self.loading_frame = ctk.CTkFrame(
            self.blur_overlay,
            width=300,
            height=200,
            corner_radius=20,
            fg_color=("white", "#2b2b2b"),
            border_width=2,
            border_color=("gray70", "gray30")
        )
        self.loading_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Loading message
        loading_message = ctk.CTkLabel(
            self.loading_frame,
            text=message,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=("black", "white")
        )
        loading_message.pack(pady=20)
        
        # GIF animation label
        self.loading_gif_label = ctk.CTkLabel(
            self.loading_frame,
            text="",
            width=100,
            height=100
        )
        self.loading_gif_label.pack(pady=10)
        
        # Progress label
        self.loading_progress_label = ctk.CTkLabel(
            self.loading_frame,
            text="0%",
            font=ctk.CTkFont(size=14),
            text_color=("black", "white")
        )
        self.loading_progress_label.pack(pady=10)
        
        # Start GIF animation
        self.loading_current_frame = 0
        self.animate_loading_gif()
        
        # Update display
        self.update()
    
    def animate_loading_gif(self):
        """Animate the loading GIF"""
        if not self.loading_gif_label or not self.loading_gif_frames:
            return
            
        # Update frame
        if self.loading_current_frame >= len(self.loading_gif_frames):
            self.loading_current_frame = 0
            
        self.loading_gif_label.configure(image=self.loading_gif_frames[self.loading_current_frame])
        self.loading_current_frame += 1
        
        # Schedule next frame (adjust timing as needed)
        self.loading_animation_job = self.after(100, self.animate_loading_gif)
    
    def update_loading_progress(self, progress, message=None):
        """Update loading progress (0-100)"""
        if self.loading_progress_label:
            self.loading_progress_label.configure(text=f"{progress}%")
            
        if message and hasattr(self, 'loading_message_label'):
            self.loading_message_label.configure(text=message)
            
        self.update()
    
    def hide_loading_screen(self):
        """Hide loading screen"""
        # Stop animation
        if self.loading_animation_job:
            self.after_cancel(self.loading_animation_job)
            self.loading_animation_job = None
            
        # Destroy loading elements
        if self.loading_frame:
            self.loading_frame.destroy()
            self.loading_frame = None
            
        if self.blur_overlay:
            self.blur_overlay.destroy()
            self.blur_overlay = None
            
        self.loading_gif_label = None
        self.loading_progress_label = None
    
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
            self.file_count_label.configure(text="1 file")
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
            self.show_loading_screen("Loading PDFs...")
            
            def load_pdfs_thread():
                try:
                    self.pdf_paths = list(file_paths)
                    self.pdf_path = self.pdf_paths[0]  # Load first PDF for viewing
                    
                    # Update progress
                    total_files = len(self.pdf_paths)
                    for i, pdf_path in enumerate(self.pdf_paths):
                        progress = int((i / total_files) * 50)  # First 50% for file processing
                        self.after(0, lambda p=progress: self.update_loading_progress(p))
                        time.sleep(0.01)  # Small delay to show progress
                    
                    # Update UI elements on main thread
                    def update_ui():
                        self.file_label.configure(text=f"{len(self.pdf_paths)} PDFs selected")
                        self.file_count_label.configure(text=f"{len(self.pdf_paths)} files")
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
                        
                        # Update progress to 100%
                        self.update_loading_progress(100)
                        time.sleep(0.2)
                        
                        # Update preview dropdown to show all files
                        self.update_excel_preview()
                        
                        # Load first PDF for viewing
                        self.load_pdf_viewer(self.pdf_paths[0])
                        
                        # Hide loading screen
                        self.hide_loading_screen()
                    
                    self.after(0, update_ui)
                    
                except Exception as e:
                    def show_error():
                        self.hide_loading_screen()
                        messagebox.showerror("Error", f"Failed to load PDFs:\n{str(e)}")
                    self.after(0, show_error)
            
            # Start thread
            thread = threading.Thread(target=load_pdfs_thread, daemon=True)
            thread.start()
    
    def select_folder(self):
        """Open dialog to select folder containing PDF files"""
        folder_path = filedialog.askdirectory(
            title="Select Folder Containing PDF Files"
        )
        
        if folder_path:
            self.show_loading_screen("Scanning folder...")
            
            def scan_folder_thread():
                try:
                    # Find all PDF files in the folder
                    self.after(0, lambda: self.update_loading_progress(25))
                    folder = Path(folder_path)
                    pdf_files = list(folder.glob("*.pdf"))
                    
                    self.after(0, lambda: self.update_loading_progress(75))
                    
                    if pdf_files:
                        self.pdf_paths = [str(f) for f in pdf_files]
                        self.pdf_path = self.pdf_paths[0]  # Load first PDF for viewing
                        
                        def update_ui():
                            self.file_label.configure(text=f"{len(self.pdf_paths)} PDFs found in folder")
                            self.file_count_label.configure(text=f"{len(self.pdf_paths)} files")
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
                            
                            self.update_loading_progress(100)
                            time.sleep(0.2)
                            self.hide_loading_screen()
                        
                        self.after(0, update_ui)
                    else:
                        def show_no_files():
                            self.hide_loading_screen()
                            messagebox.showwarning("No PDFs Found", "No PDF files found in the selected folder.")
                            self.status_label.configure(text="No PDFs found in folder")
                        self.after(0, show_no_files)
                        
                except Exception as e:
                    def show_error():
                        self.hide_loading_screen()
                        messagebox.showerror("Error", f"Failed to scan folder:\n{str(e)}")
                    self.after(0, show_error)
            
            # Start thread
            thread = threading.Thread(target=scan_folder_thread, daemon=True)
            thread.start()
    
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
            self.select_table_btn.configure(state="normal")
            self.save_selection_btn.configure(state="normal")
            self.clear_selections_btn.configure(state="normal")
            self.auto_extract_btn.configure(state="normal")
            
            # Enable apply button if multiple files selected
            if self.pdf_paths and len(self.pdf_paths) > 1:
                self.apply_all_btn.configure(state="normal")
            # Enable sticky overlay nav buttons
            try:
                self.sticky_prev.configure(state="normal")
                self.sticky_next.configure(state="normal")
            except Exception:
                pass
            # Enable clear preview if previews exist with data
            try:
                has_data = any((not df.empty) for df in self.file_previews.values()) if self.file_previews else False
                if has_data:
                    self.clear_preview_btn.configure(state="normal")
                else:
                    self.clear_preview_btn.configure(state="disabled")
            except Exception:
                pass
            
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
            
            # Update page label and sticky center
            page_text = f"Page {self.current_page + 1} of {len(self.pdf_document)}"
            self.page_label.configure(text=page_text)
            self.sticky_center.configure(text=page_text)
            
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
        """Save all temporary selections for current page and extract data with formatting"""
        if self.current_page not in self.temp_selections or not self.temp_selections[self.current_page]:
            messagebox.showinfo("Info", "No selections to save on this page!")
            return
        
        self.show_loading_screen("Saving selections...")
        
        def save_selections_thread():
            try:
                self.after(0, lambda: self.update_loading_progress(25))
                
                # Get current viewing file (could be different from batch files)
                current_file = self.pdf_path
                
                # Extract data from all temp selections on this page
                with pdfplumber.open(current_file) as pdf:
                    page = pdf.pages[self.current_page]
                    file_name = Path(current_file).name
                    
                    selection_count = len(self.temp_selections[self.current_page])
                    
                    for idx, (bbox, rect_id) in enumerate(self.temp_selections[self.current_page]):
                        # Update progress for each selection
                        progress = 25 + int((idx / selection_count) * 50)
                        self.after(0, lambda p=progress: self.update_loading_progress(p))
                        
                        # Extract table data
                        cropped = page.crop(bbox)
                        table = cropped.extract_table()
                        
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
                        
                        # Change rectangle color to blue (saved) on main thread
                        self.after(0, lambda rid=rect_id: self.pdf_canvas.itemconfig(rid, outline="blue"))
                        
                        # Move to saved selections
                        if self.current_page not in self.saved_selections:
                            self.saved_selections[self.current_page] = []
                        self.saved_selections[self.current_page].append(bbox)
                
                # Clear temp selections for this page
                saved_count = len(self.temp_selections[self.current_page])
                self.temp_selections[self.current_page] = []
                
                def finish_save():
                    self.update_loading_progress(100)
                    time.sleep(0.2)
                    
                    # Update preview
                    self.update_excel_preview()
                    self.export_btn.configure(state="normal")
                    
                    total_saved = sum(len(sels) for sels in self.saved_selections.values())
                    self.status_label.configure(text=f"✓ Saved {saved_count} selection(s) from page {self.current_page + 1}! Total saved: {total_saved}")
                    
                    self.hide_loading_screen()
                
                self.after(0, finish_save)
                
            except Exception as e:
                def show_error():
                    self.hide_loading_screen()
                    messagebox.showerror("Error", f"Failed to save selections:\n{str(e)}")
                    print(f"Save error: {e}")
                self.after(0, show_error)
        
        # Start thread
        thread = threading.Thread(target=save_selections_thread, daemon=True)
        thread.start()
    
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
        # Auto-extract should run across all selected PDFs (batch) or the single loaded PDF
        files = self.pdf_paths if self.pdf_paths else ([self.pdf_path] if self.pdf_path else [])
        if not files:
            messagebox.showinfo("Info", "No PDF(s) loaded!")
            return

        self.show_loading_screen("Auto-extracting tables...")
        
        def auto_extract_thread():
            try:
                header_crop = 90  # Top 80px to remove
                footer_crop = 90  # Bottom 80px to remove
                processed = 0
                total_files = len(files)

                for file_idx, current_file in enumerate(files):
                    try:
                        file_name = Path(current_file).name
                        
                        # Update progress
                        progress = int((file_idx / total_files) * 90)
                        self.after(0, lambda p=progress, fn=file_name: (
                            self.update_loading_progress(p),
                            self.status_label.configure(text=f"Auto-extracting tables from {fn}...")
                        ))

                        with pdfplumber.open(current_file) as pdf:
                            all_tables = []
                            
                            for page in pdf.pages:
                                page_height = page.height
                                cropped_page = page.crop((0, header_crop, page.width, page_height - footer_crop))
                                tables = cropped_page.extract_tables()
                                
                                if tables:
                                    # Get table objects with cell bounding boxes for color detection
                                    table_settings = {
                                        "vertical_strategy": "lines",
                                        "horizontal_strategy": "lines",
                                        "snap_tolerance": 3,
                                        "join_tolerance": 3,
                                    }
                                    tables_with_cells = cropped_page.find_tables(table_settings)
                                    
                                    for table_idx, table in enumerate(tables):
                                        if table and len(table) > 0:
                                            all_tables.append(table)

                            if all_tables:
                                combined_rows = []
                                for table_idx, table in enumerate(all_tables):
                                    combined_rows.extend(table)
                                    if table_idx < len(all_tables) - 1:
                                        num_cols = len(table[0]) if table else 1
                                        combined_rows.extend([[None] * num_cols] * 2)

                                df = pd.DataFrame(combined_rows)
                                self.file_previews[file_name] = df
                                processed += 1
                    except Exception as e:
                        # Continue to next file if one fails
                        print(f"Auto-extract failed for {current_file}: {e}")
                        continue

                # Final UI updates
                def finish_extraction():
                    self.update_loading_progress(100)
                    time.sleep(0.3)
                    
                    if processed > 0:
                        self.update_excel_preview()
                        self.export_btn.configure(state="normal")
                        self.clear_preview_btn.configure(state="normal")
                        self.status_label.configure(text=f"✓ Auto-extracted tables from {processed}/{len(files)} files")
                        self.hide_loading_screen()
                        messagebox.showinfo("Success", f"Auto-extracted tables from {processed}/{len(files)} files.\n\nCheck Excel Preview tab.")
                    else:
                        self.hide_loading_screen()
                        messagebox.showwarning("No Tables", "No tables found in the selected PDF(s).")
                        self.status_label.configure(text="⚠ No tables found")

                self.after(0, finish_extraction)

            except Exception as e:
                def show_error():
                    self.hide_loading_screen()
                    messagebox.showerror("Error", f"Auto-extraction failed:\n{str(e)}")
                    self.status_label.configure(text="✗ Auto-extraction failed")
                    import traceback
                    traceback.print_exc()
                self.after(0, show_error)

        # Start thread
        thread = threading.Thread(target=auto_extract_thread, daemon=True)
        thread.start()
    
    def apply_selections_to_all_files(self):
        """Apply saved selections to all imported PDF files with header/footer removal"""
        if not self.pdf_paths:
            messagebox.showinfo("Info", "No multiple files selected!")
            return
        
        if not self.saved_selections:
            messagebox.showerror("Error", "No selections saved! Please mark and save selections first.")
            return
        
        self.show_loading_screen("Applying selections to all files...")
        
        def apply_selections_thread():
            try:
                processed_count = 0
                header_crop = 80  # Top 80px to remove
                footer_crop = 80  # Bottom 80px to remove
                total_files = len(self.pdf_paths)
                
                for file_idx, pdf_path in enumerate(self.pdf_paths):
                    file_name = Path(pdf_path).name
                    
                    # Update progress
                    progress = int((file_idx / total_files) * 90)
                    self.after(0, lambda p=progress, fn=file_name: (
                        self.update_loading_progress(p),
                        self.status_label.configure(text=f"Applying selections to {fn}...")
                    ))
                    
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
                
                def finish_apply():
                    self.update_loading_progress(100)
                    time.sleep(0.3)
                    
                    # Update preview
                    self.update_excel_preview()
                    self.export_btn.configure(state="normal")
                    
                    self.status_label.configure(text=f"✓ Applied to {processed_count} files with header/footer removed!")
                    self.hide_loading_screen()
                    messagebox.showinfo("Success", f"Successfully applied selections to {processed_count} files!\n\nHeader/footer (80px) removed automatically.\n\nCheck the Excel Preview tab.")
                
                self.after(0, finish_apply)
                
            except Exception as e:
                def show_error():
                    self.hide_loading_screen()
                    messagebox.showerror("Error", f"Failed to apply selections:\n{str(e)}")
                    print(f"Apply error: {e}")
                self.after(0, show_error)
        
        # Start thread
        thread = threading.Thread(target=apply_selections_thread, daemon=True)
        thread.start()
    
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
            # Clear treeview
            for item in self.excel_display.get_children():
                self.excel_display.delete(item)
            try:
                self.clear_preview_btn.configure(state="disabled")
            except Exception:
                pass
            return
        
        # Show list of files
        file_names = list(self.file_previews.keys())
        self.sheet_selector.configure(values=file_names, state="readonly")
        self.sheet_selector.set(file_names[0])
        
        # Display first file
        self.display_file_preview(file_names[0])
        # Enable Clear Preview if any file has data
        try:
            has_data = any((not df.empty) for df in self.file_previews.values())
            if has_data:
                self.clear_preview_btn.configure(state="normal")
            else:
                self.clear_preview_btn.configure(state="disabled")
        except Exception:
            pass
    
    def on_file_selected(self, file_name):
        """Handle file selection change"""
        self.display_file_preview(file_name)

    def clear_all_previews(self):
        """Clear all Excel preview data (reset preview dropdown and display)"""
        self.show_loading_screen("Clearing previews...")
        
        def clear_previews_thread():
            try:
                self.after(0, lambda: self.update_loading_progress(25))
                
                # Clear file previews
                self.file_previews = {}
                
                self.after(0, lambda: self.update_loading_progress(50))
                
                def update_ui():
                    # Reset selector and display
                    try:
                        self.sheet_selector.configure(values=["No data available"], state="disabled")
                    except Exception:
                        pass
                    
                    self.update_loading_progress(75)
                    
                    # Clear treeview
                    for item in self.excel_display.get_children():
                        self.excel_display.delete(item)
                    
                    # Disable export and clear-preview button
                    try:
                        self.export_btn.configure(state="disabled")
                        self.clear_preview_btn.configure(state="disabled")
                    except Exception:
                        pass
                    
                    self.update_loading_progress(100)
                    time.sleep(0.2)
                    
                    self.status_label.configure(text="Excel preview cleared")
                    self.hide_loading_screen()
                
                self.after(0, update_ui)
                
            except Exception as e:
                def show_error():
                    self.hide_loading_screen()
                    messagebox.showerror("Error", f"Failed to clear previews:\n{str(e)}")
                self.after(0, show_error)
        
        # Start thread
        thread = threading.Thread(target=clear_previews_thread, daemon=True)
        thread.start()
    
    def display_file_preview(self, file_name):
        """Display selected file preview data in Excel-like grid"""
        if file_name not in self.file_previews:
            return
        
        df = self.file_previews[file_name]
        
        # Clear existing treeview
        for item in self.excel_display.get_children():
            self.excel_display.delete(item)
        
        # Clear existing columns
        self.excel_display["columns"] = []
        
        # Check if dataframe is empty
        if df.empty:
            # Show message in treeview
            self.excel_display["columns"] = ["Message"]
            self.excel_display.heading("#0", text="")
            self.excel_display.column("#0", width=0, stretch=False)
            self.excel_display.heading("Message", text=f"File: {file_name}")
            self.excel_display.column("Message", width=800, anchor="w")
            
            self.excel_display.insert("", "end", values=("No data extracted yet.",))
            self.excel_display.insert("", "end", values=("",))
            self.excel_display.insert("", "end", values=("To extract data:",))
            self.excel_display.insert("", "end", values=("1. Mark selection areas on PDF",))
            self.excel_display.insert("", "end", values=("2. Click 💾 Save Selections",))
            self.excel_display.insert("", "end", values=("3. Click 📋 Apply to All Files to process all PDFs",))
        else:
            # Set up columns - use Excel-style letters (A, B, C, etc.)
            num_cols = len(df.columns)
            col_letters = [self.excel_column_letter(i) for i in range(num_cols)]
            
            self.excel_display["columns"] = col_letters
            
            # Hide the first column (tree column)
            self.excel_display.heading("#0", text="Row")
            self.excel_display.column("#0", width=50, anchor="center")
            
            # Configure column headers
            for i, col_letter in enumerate(col_letters):
                self.excel_display.heading(col_letter, text=col_letter)
                self.excel_display.column(col_letter, width=100, anchor="center")
            
            # Add data rows
            for idx, row in df.iterrows():
                # Convert row to list and replace None/NaN with empty string
                row_values = [str(val) if pd.notna(val) and val is not None else "" for val in row]
                self.excel_display.insert("", "end", text=str(idx + 1), values=row_values)
    
    def excel_column_letter(self, n):
        """Convert column number to Excel-style letter (0->A, 1->B, 25->Z, 26->AA, etc.)"""
        result = ""
        while n >= 0:
            result = chr(n % 26 + 65) + result
            n = n // 26 - 1
            if n < 0:
                break
        return result
    
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
        """Export all extracted tables to Excel with thick borders, bold detection, and merged headers"""
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
                
                from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
                from openpyxl.utils import get_column_letter
                
                output_path = Path(save_folder)
                
                # Export each file preview
                for file_name, df in self.file_previews.items():
                    excel_name = Path(file_name).stem + ".xlsx"
                    excel_path = output_path / excel_name
                    
                    # Clean dataframe before export - PRESERVE TABLE STRUCTURE
                    df_cleaned = df.copy()
                    
                    # Step 1: Reset column names to numeric indices first
                    df_cleaned.columns = range(len(df_cleaned.columns))
                    
                    # Step 2: DETECT TABLE BOUNDARIES with improved logic for better detection
                    table_boundaries = []
                    current_table = None
                    empty_row_count = 0
                    max_empty_gap = 3  # Allow up to 3 empty rows within a table
                    
                    for row_idx in range(len(df_cleaned)):
                        row = df_cleaned.iloc[row_idx]
                        row_has_data = False
                        first_col = None
                        last_col = None
                        
                        for col_idx in range(len(row)):
                            cell = row.iloc[col_idx]
                            if pd.notna(cell) and str(cell).strip():
                                row_has_data = True
                                if first_col is None:
                                    first_col = col_idx
                                last_col = col_idx
                        
                        if row_has_data:
                            empty_row_count = 0  # Reset empty row counter
                            
                            if current_table is None:
                                current_table = {
                                    'start_row': row_idx,
                                    'end_row': row_idx,
                                    'start_col': first_col,
                                    'end_col': last_col
                                }
                            else:
                                current_table['end_row'] = row_idx
                                current_table['start_col'] = min(current_table['start_col'], first_col)
                                current_table['end_col'] = max(current_table['end_col'], last_col)
                        else:
                            empty_row_count += 1
                            
                            # Only end the table if we have too many consecutive empty rows
                            if current_table is not None and empty_row_count > max_empty_gap:
                                table_boundaries.append(current_table)
                                current_table = None
                                empty_row_count = 0
                    
                    if current_table is not None:
                        table_boundaries.append(current_table)
                    
                    # Step 3: Smart cleaning - preserve table structure, clean outside tables
                    def is_inside_table(row_idx, col_idx):
                        """Check if cell is inside any detected table boundary"""
                        for table in table_boundaries:
                            if (table['start_row'] <= row_idx <= table['end_row'] and 
                                table['start_col'] <= col_idx <= table['end_col']):
                                return True
                        return False
                    
                    # Step 4: First remove completely empty columns outside tables
                    cols_to_keep = []
                    for col_idx in range(len(df_cleaned.columns)):
                        col_data = df_cleaned.iloc[:, col_idx]
                        
                        # Check if column intersects with any table
                        col_in_table = any(table['start_col'] <= col_idx <= table['end_col'] for table in table_boundaries)
                        
                        if col_in_table:
                            # Always keep columns that are part of tables
                            cols_to_keep.append(col_idx)
                        else:
                            # For columns outside tables, check if they have any data
                            has_data = col_data.notna().any() and any(str(cell).strip() != '' for cell in col_data if pd.notna(cell))
                            if has_data:
                                cols_to_keep.append(col_idx)
                    
                    # Create column mapping for updating table boundaries
                    old_to_new_col = {}
                    for new_idx, old_idx in enumerate(cols_to_keep):
                        old_to_new_col[old_idx] = new_idx
                    
                    if cols_to_keep:
                        df_cleaned = df_cleaned.iloc[:, cols_to_keep]
                        df_cleaned.columns = range(len(df_cleaned.columns))  # Reset to 0, 1, 2, ...
                        
                        # Update table boundaries after column removal
                        updated_table_boundaries = []
                        for table in table_boundaries:
                            # Map old column indices to new ones
                            new_start_col = None
                            new_end_col = None
                            
                            for old_col in range(table['start_col'], table['end_col'] + 1):
                                if old_col in old_to_new_col:
                                    new_col = old_to_new_col[old_col]
                                    if new_start_col is None:
                                        new_start_col = new_col
                                    new_end_col = new_col
                            
                            if new_start_col is not None and new_end_col is not None:
                                updated_table_boundaries.append({
                                    'start_row': table['start_row'],
                                    'end_row': table['end_row'],
                                    'start_col': new_start_col,
                                    'end_col': new_end_col
                                })
                        
                        table_boundaries = updated_table_boundaries
                    
                    # Step 5: Improved table-unit shifting with fallback for missed tables
                    # Process each table as a unit to maintain internal spacing and alignment
                    processed_tables = set()
                    
                    for idx in range(len(df_cleaned)):
                        row = df_cleaned.iloc[idx]
                        
                        # Skip completely empty rows
                        if row.isna().all() or all(str(cell).strip() == '' for cell in row if pd.notna(cell)):
                            continue
                        
                        # Check if this row is part of any table
                        current_table = None
                        for table in table_boundaries:
                            if table['start_row'] <= idx <= table['end_row']:
                                current_table = table
                                break
                        
                        if current_table and id(current_table) not in processed_tables:
                            # Process entire table as a unit
                            processed_tables.add(id(current_table))
                            
                            # Find the minimum leading empty columns across ALL data rows in this table
                            min_leading_empty = float('inf')
                            table_has_data = False
                            
                            for table_row_idx in range(current_table['start_row'], current_table['end_row'] + 1):
                                table_row = df_cleaned.iloc[table_row_idx]
                                
                                # Skip empty rows in the table
                                if table_row.isna().all() or all(str(cell).strip() == '' for cell in table_row if pd.notna(cell)):
                                    continue
                                
                                # Count leading empty cells in this row
                                leading_empty = 0
                                for cell in table_row:
                                    if pd.isna(cell) or str(cell).strip() == '':
                                        leading_empty += 1
                                    else:
                                        break
                                
                                # Only consider rows that have actual data
                                if leading_empty < len(table_row):
                                    min_leading_empty = min(min_leading_empty, leading_empty)
                                    table_has_data = True
                            
                            # If we found consistent leading empty space, shift the entire table
                            if table_has_data and min_leading_empty > 0 and min_leading_empty != float('inf'):
                                for table_row_idx in range(current_table['start_row'], current_table['end_row'] + 1):
                                    table_row = df_cleaned.iloc[table_row_idx]
                                    
                                    # Shift ALL rows in the table by the same amount (preserves internal structure)
                                    row_values = table_row.tolist()
                                    shifted = row_values[min_leading_empty:] + [None] * min_leading_empty
                                    df_cleaned.iloc[table_row_idx] = shifted
                        
                        elif not current_table:
                            # Rows OUTSIDE tables - shift individually
                            leading_empty = 0
                            for cell in row:
                                if pd.isna(cell) or str(cell).strip() == '':
                                    leading_empty += 1
                                else:
                                    break
                            
                            if leading_empty > 0 and leading_empty < len(row):
                                row_values = row.tolist()
                                shifted = row_values[leading_empty:] + [None] * leading_empty
                                df_cleaned.iloc[idx] = shifted
                    
                    # Step 5.5: Additional pass for any remaining rows with leading empty space
                    # This catches any rows that might have been missed by table detection
                    # BUT only shift rows that are OUTSIDE detected tables
                    for idx in range(len(df_cleaned)):
                        row = df_cleaned.iloc[idx]
                        
                        # Skip completely empty rows
                        if row.isna().all() or all(str(cell).strip() == '' for cell in row if pd.notna(cell)):
                            continue
                        
                        # Check if this row is inside any detected table
                        row_in_table = False
                        for table in table_boundaries:
                            if table['start_row'] <= idx <= table['end_row']:
                                row_in_table = True
                                break
                        
                        # Only shift if row is OUTSIDE tables
                        if not row_in_table:
                            # Count leading empty cells
                            leading_empty = 0
                            for cell in row:
                                if pd.isna(cell) or str(cell).strip() == '':
                                    leading_empty += 1
                                else:
                                    break
                            
                            # If row still has leading empty space, shift it
                            if leading_empty > 0 and leading_empty < len(row):
                                row_values = row.tolist()
                                shifted = row_values[leading_empty:] + [None] * leading_empty
                                df_cleaned.iloc[idx] = shifted
                    
                    # Step 6: Remove excessive empty rows (keep max 2 consecutive empty rows as separators)
                    rows_to_keep = []
                    consecutive_empty = 0
                    
                    for idx in range(len(df_cleaned)):
                        row = df_cleaned.iloc[idx]
                        is_empty = row.isna().all() or all(str(cell).strip() == '' for cell in row if pd.notna(cell))
                        
                        # Check if row contains header/footer text like "COMMERCIAL CONFIDENTIAL"
                        is_header_footer_row = False
                        for cell in row:
                            if pd.notna(cell):
                                cell_text = str(cell).strip().lower()
                                if 'commercial confidential' in cell_text:
                                    is_header_footer_row = True
                                    break
                        
                        if is_header_footer_row:
                            # Skip this row entirely
                            continue
                        elif is_empty:
                            consecutive_empty += 1
                            # Keep only first 2 consecutive empty rows
                            if consecutive_empty <= 2:
                                rows_to_keep.append(idx)
                        else:
                            consecutive_empty = 0
                            rows_to_keep.append(idx)
                    
                    # Apply the filter
                    if rows_to_keep:
                        df_cleaned = df_cleaned.iloc[rows_to_keep].reset_index(drop=True)
                    
                    # Write to Excel
                    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                        df_cleaned.to_excel(writer, sheet_name="Extracted Data", index=False, header=False)
                        
                        # Get worksheet
                        worksheet = writer.sheets["Extracted Data"]
                        
                        # Define thick border style for tables
                        thick_border = Border(
                            left=Side(style='medium'),
                            right=Side(style='medium'),
                            top=Side(style='medium'),
                            bottom=Side(style='medium')
                        )
                        
                        # Detect table boundaries (non-empty cell regions)
                        table_ranges = []
                        current_table = None
                        
                        for row_idx in range(1, len(df_cleaned) + 1):
                            row_has_data = False
                            first_col = None
                            last_col = None
                            
                            for col_idx in range(1, len(df_cleaned.columns) + 1):
                                cell = worksheet.cell(row_idx, col_idx)
                                if cell.value is not None and str(cell.value).strip():
                                    row_has_data = True
                                    if first_col is None:
                                        first_col = col_idx
                                    last_col = col_idx
                            
                            if row_has_data:
                                if current_table is None:
                                    current_table = {
                                        'start_row': row_idx,
                                        'end_row': row_idx,
                                        'start_col': first_col,
                                        'end_col': last_col
                                    }
                                else:
                                    current_table['end_row'] = row_idx
                                    current_table['start_col'] = min(current_table['start_col'], first_col)
                                    current_table['end_col'] = max(current_table['end_col'], last_col)
                            else:
                                if current_table is not None:
                                    table_ranges.append(current_table)
                                    current_table = None
                        
                        if current_table is not None:
                            table_ranges.append(current_table)
                        
                        # Apply formatting to each table
                        for table_idx, table_range in enumerate(table_ranges):
                            start_row = table_range['start_row']
                            end_row = table_range['end_row']
                            start_col = table_range['start_col']
                            end_col = table_range['end_col']
                            
                            # First pass: Detect Risk Grade row and mark the score cell
                            risk_grade_row = None
                            risk_grade_score_col = None
                            
                            for row_idx in range(start_row, end_row + 1):
                                for col_idx in range(start_col, end_col + 1):
                                    cell = worksheet.cell(row_idx, col_idx)
                                    cell_value = str(cell.value).strip() if cell.value else ""
                                    
                                    # Check if this row contains "Risk Grade"
                                    if "Risk Grade" in cell_value:
                                        risk_grade_row = row_idx
                                        
                                        # Find the cell with just a number (1-10) in this row
                                        for score_col in range(start_col, end_col + 1):
                                            score_cell = worksheet.cell(row_idx, score_col)
                                            score_value = str(score_cell.value).strip() if score_cell.value else ""
                                            
                                            # Check if this is a single digit number (the actual score)
                                            if score_value.isdigit() and 1 <= int(score_value) <= 10:
                                                risk_grade_score_col = score_col
                                                break
                                        break
                                if risk_grade_row:
                                    break
                            
                            # Second pass: Apply thick borders and formatting to table cells
                            for row_idx in range(start_row, end_row + 1):
                                for col_idx in range(start_col, end_col + 1):
                                    cell = worksheet.cell(row_idx, col_idx)
                                    
                                    # Apply thick border only to table cells
                                    cell.border = thick_border
                                    
                                    # Center align
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                                    
                                    # Check if this is the Risk Grade score cell - make it bold with yellow background
                                    is_risk_score_cell = (row_idx == risk_grade_row and col_idx == risk_grade_score_col)
                                    
                                    if is_risk_score_cell:
                                        # Make the risk grade number stand out
                                        cell.font = Font(bold=True, size=14, color="000000")  # Bold, larger, black
                                        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Yellow background
                            
                            # Smart merging: Merge horizontally (same row) when cells have identical values
                            for row_idx in range(start_row, end_row + 1):
                                col = start_col
                                while col <= end_col:
                                    cell = worksheet.cell(row_idx, col)
                                    cell_value = cell.value
                                    
                                    # Skip empty cells
                                    if cell_value is None or str(cell_value).strip() == "":
                                        col += 1
                                        continue
                                    
                                    # Find consecutive cells with same value in this row
                                    merge_end = col
                                    while merge_end < end_col:
                                        next_cell = worksheet.cell(row_idx, merge_end + 1)
                                        next_value = next_cell.value
                                        
                                        # Check if values match (handling None and whitespace)
                                        if next_value is not None and str(cell_value).strip() == str(next_value).strip():
                                            merge_end += 1
                                        else:
                                            break
                                    
                                    # Merge if more than one cell has the same value
                                    if merge_end > col:
                                        try:
                                            range_str = f"{get_column_letter(col)}{row_idx}:{get_column_letter(merge_end)}{row_idx}"
                                            if range_str not in [str(m) for m in worksheet.merged_cells]:
                                                worksheet.merge_cells(range_str)
                                                # Re-apply formatting to merged cell
                                                merged_cell = worksheet.cell(row_idx, col)
                                                merged_cell.border = thick_border
                                                merged_cell.alignment = Alignment(horizontal='center', vertical='center')
                                        except Exception as e:
                                            print(f"Merge error: {e}")
                                    
                                    col = merge_end + 1
                            
                            # Vertical merging: Merge down (same column) when cells have identical values
                            for col_idx in range(start_col, end_col + 1):
                                row = start_row
                                while row <= end_row:
                                    cell = worksheet.cell(row, col_idx)
                                    cell_value = cell.value
                                    
                                    # Skip empty cells
                                    if cell_value is None or str(cell_value).strip() == "":
                                        row += 1
                                        continue
                                    
                                    # Find consecutive cells with same value in this column
                                    merge_end_row = row
                                    while merge_end_row < end_row:
                                        next_cell = worksheet.cell(merge_end_row + 1, col_idx)
                                        next_value = next_cell.value
                                        
                                        # Check if values match
                                        if next_value is not None and str(cell_value).strip() == str(next_value).strip():
                                            merge_end_row += 1
                                        else:
                                            break
                                    
                                    # Merge if more than one cell has the same value (at least 2 rows)
                                    if merge_end_row > row:
                                        try:
                                            range_str = f"{get_column_letter(col_idx)}{row}:{get_column_letter(col_idx)}{merge_end_row}"
                                            if range_str not in [str(m) for m in worksheet.merged_cells]:
                                                worksheet.merge_cells(range_str)
                                                # Re-apply formatting to merged cell
                                                merged_cell = worksheet.cell(row, col_idx)
                                                merged_cell.border = thick_border
                                                merged_cell.alignment = Alignment(horizontal='center', vertical='center')
                                        except Exception as e:
                                            print(f"Vertical merge error: {e}")
                                    
                                    row = merge_end_row + 1
                        
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

