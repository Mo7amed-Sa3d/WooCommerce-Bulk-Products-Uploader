import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import queue
from datetime import datetime
from dotenv import load_dotenv
import pandas as pd
from excel_processor import ExcelProductProcessor
# Import our modules
from api_client import WooCommerceAPI, WordPressMediaAPI
from upload_queue import UploadQueueManager
from ai_helper import AIHelper
from bulk_processor import BulkProductProcessor
from utils import setup_logging, validate_image_paths, format_price, build_category_tree

load_dotenv()
setup_logging()

class ProductUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WooCommerce Product Uploader v3.0 - Bulk + Single")
        self.root.geometry("1200x900")
        
        # APIs
        self._wc_api = WooCommerceAPI()
        self._wp_api = WordPressMediaAPI()
        self.ai_helper = AIHelper()
        self.bulk_processor = BulkProductProcessor()
        
            # Add Excel processor
        self.excel_processor = ExcelProductProcessor()

        # Queue Manager
        self.queue_manager = UploadQueueManager(
            wc_api=self._wc_api,
            wp_api=self._wp_api,
            max_workers=3
        )
        self.queue_manager.on_upload_complete = self._on_upload_complete
        
        # Variables
        self.categories = []
        self.category_tree = []
        self.images = []
        self.category_dict = {}
        self.upload_history = []
        
        # Bulk upload variables
        self.bulk_products = []
        self.bulk_category_id = None
        self.bulk_directory = ""
        self.current_batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # UI Setup
        self.setup_ui()
        
        # Load categories
        self.load_categories()
        
        # Setup periodic queue update
        self.update_queue_status()
    
    @property
    def wc_api(self):
        return self._wc_api
    
    @property
    def wp_api(self):
        return self._wp_api
    
    def setup_ui(self):
        # Configure grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # Single Product Tab
        self.product_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.product_frame, text="📤 Single Product")
        
        # Bulk Upload Tab
        self.bulk_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.bulk_frame, text="📦 Bulk Upload")
        
        # Queue Tab
        self.queue_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.queue_frame, text="📊 Upload Queue")
        
        # Setup frames
        self.setup_product_frame()
        self.setup_bulk_frame()
        self.setup_queue_frame()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))
    
    def setup_product_frame(self):
        # Configure product frame grid
        for i in range(4):
            self.product_frame.columnconfigure(i, weight=1)
        
        # Product Title
        row = 0
        ttk.Label(self.product_frame, text="Product Title:*").grid(
            row=row, column=0, sticky="w", padx=10, pady=5
        )
        self.title_var = tk.StringVar()
        self.title_entry = ttk.Entry(self.product_frame, textvariable=self.title_var, width=50)
        self.title_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=5, columnspan=2)
        
        if self.ai_helper.available:
            ttk.Button(self.product_frame, text="AI Generate Title", 
                      command=self.generate_ai_title).grid(
                row=row, column=3, padx=5, pady=5
            )
        
        # Price
        row += 1
        ttk.Label(self.product_frame, text="Price ($):*").grid(
            row=row, column=0, sticky="w", padx=10, pady=5
        )
        self.price_var = tk.StringVar()
        self.price_entry = ttk.Entry(self.product_frame, textvariable=self.price_var, width=20)
        self.price_entry.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        
        # Category
        ttk.Label(self.product_frame, text="Category:*").grid(
            row=row, column=2, sticky="w", padx=10, pady=5
        )
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(
            self.product_frame, 
            textvariable=self.category_var,
            width=30,
            state="readonly"
        )
        self.category_combo.grid(row=row, column=3, sticky="ew", padx=5, pady=5)
        
        # Description
        row += 1
        ttk.Label(self.product_frame, text="Description:").grid(
            row=row, column=0, sticky="nw", padx=10, pady=5
        )
        self.desc_text = tk.Text(self.product_frame, height=12, width=50)
        self.desc_text.grid(row=row, column=1, sticky="nsew", padx=5, pady=5, columnspan=2, rowspan=2)
        
        desc_scroll = ttk.Scrollbar(self.product_frame, orient=tk.VERTICAL, command=self.desc_text.yview)
        desc_scroll.grid(row=row, column=3, sticky="ns", pady=5, rowspan=2)
        self.desc_text['yscrollcommand'] = desc_scroll.set
        
        if self.ai_helper.available:
            ttk.Button(self.product_frame, text="AI Generate Description", 
                      command=self.generate_ai_description).grid(
                row=row, column=3, padx=5, pady=5, sticky="n"
            )
        
        # Images Section
        row += 2
        ttk.Label(self.product_frame, text="Product Images:*").grid(
            row=row, column=0, sticky="w", padx=10, pady=5
        )
        
        # Image list frame
        image_frame = ttk.LabelFrame(self.product_frame, text="Selected Images", padding=10)
        image_frame.grid(row=row, column=1, columnspan=3, sticky="nsew", padx=5, pady=5, rowspan=3)
        image_frame.columnconfigure(0, weight=1)
        image_frame.rowconfigure(0, weight=1)
        
        # Image listbox with scrollbar
        self.image_listbox = tk.Listbox(image_frame, height=8)
        self.image_listbox.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        list_scroll = ttk.Scrollbar(image_frame, orient=tk.VERTICAL, command=self.image_listbox.yview)
        list_scroll.grid(row=0, column=1, sticky="ns")
        self.image_listbox['yscrollcommand'] = list_scroll.set
        
        # Image control buttons
        btn_frame = ttk.Frame(image_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(btn_frame, text="📁 Add Images", 
                  command=self.add_images).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🗑️ Remove Selected", 
                  command=self.remove_image).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="⬆️ Move Up", 
                  command=lambda: self.move_image(-1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="⬇️ Move Down", 
                  command=lambda: self.move_image(1)).pack(side=tk.LEFT, padx=2)
        
        # Note about images
        ttk.Label(self.product_frame, 
                 text="Note: First image is featured image. Drag or use buttons to reorder.",
                 font=('TkDefaultFont', 9, 'italic')).grid(
            row=row+3, column=1, columnspan=3, sticky="w", padx=5, pady=(0, 10)
        )
        
        # Control buttons
        row += 4
        button_frame = ttk.Frame(self.product_frame)
        button_frame.grid(row=row, column=0, columnspan=4, pady=20)
        
        ttk.Button(button_frame, text="➕ Add to Upload Queue", 
                  command=self.queue_product_upload, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🗑️ Clear Form", 
                  command=self.clear_form).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🔄 Refresh Categories", 
                  command=self.load_categories).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🔗 Test Connection", 
                  command=self.test_connection).pack(side=tk.LEFT, padx=5)
        
        # Queue status in product tab
        row += 1
        self.queue_status_var = tk.StringVar()
        self.queue_status_var.set("Queue: 0 items waiting | 0 active uploads")
        queue_status = ttk.Label(self.product_frame, 
                                textvariable=self.queue_status_var,
                                font=('TkDefaultFont', 10, 'bold'))
        queue_status.grid(row=row, column=0, columnspan=4, pady=(10, 5))
        
        # Style configuration
        style = ttk.Style()
        style.configure("Accent.TButton", font=('TkDefaultFont', 10, 'bold'))
    
    def setup_bulk_frame(self):
        """Setup the bulk upload tab for Excel files"""
        # Configure bulk frame grid
        for i in range(3):
            self.bulk_frame.columnconfigure(i, weight=1)
        
        # Header
        ttk.Label(self.bulk_frame, text="📊 Excel Bulk Product Upload", 
                font=('TkDefaultFont', 14, 'bold')).grid(
            row=0, column=0, columnspan=3, pady=(10, 20)
        )
        
        # Excel File Selection
        row = 1
        ttk.Label(self.bulk_frame, text="Excel File:").grid(
            row=row, column=0, sticky="w", padx=10, pady=5
        )
        
        self.excel_file_var = tk.StringVar()
        excel_entry = ttk.Entry(self.bulk_frame, textvariable=self.excel_file_var, width=50)
        excel_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Button(self.bulk_frame, text="📁 Browse Excel", 
                command=self.select_excel_file).grid(
            row=row, column=2, padx=5, pady=5
        )
        
        # Category Selection for Bulk
        row += 1
        ttk.Label(self.bulk_frame, text="Category for All Products:*").grid(
            row=row, column=0, sticky="w", padx=10, pady=5
        )
        
        self.bulk_category_var = tk.StringVar()
        self.bulk_category_combo = ttk.Combobox(
            self.bulk_frame,
            textvariable=self.bulk_category_var,
            width=40,
            state="readonly"
        )
        self.bulk_category_combo.grid(row=row, column=1, sticky="ew", padx=5, pady=5, columnspan=2)
        
        # Control Buttons Frame
        row += 1
        control_frame = ttk.Frame(self.bulk_frame)
        control_frame.grid(row=row, column=0, columnspan=3, pady=10)
        
        ttk.Button(control_frame, text="🔍 Load & Validate Excel", 
                command=self.load_excel_file,
                style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="📝 Create Template", 
                command=self.create_excel_template).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="📄 Export to CSV", 
                command=self.export_bulk_products).pack(side=tk.LEFT, padx=5)
        
        # File Info Frame
        row += 1
        info_frame = ttk.LabelFrame(self.bulk_frame, text="File Information", padding=10)
        info_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=10, pady=10)
        
        self.file_info_var = tk.StringVar()
        self.file_info_var.set("No Excel file loaded")
        file_info_label = ttk.Label(info_frame, textvariable=self.file_info_var, justify=tk.LEFT)
        file_info_label.grid(row=0, column=0, sticky="w")
        
        # Products Preview Frame
        row += 1
        preview_frame = ttk.LabelFrame(self.bulk_frame, text="Products Preview", padding=10)
        preview_frame.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=10, pady=10)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Create Treeview for products
        columns = ('row', 'title', 'sku', 'price', 'images', 'images_path', 'status')
        self.bulk_tree = ttk.Treeview(
            preview_frame,
            columns=columns,
            show='headings',
            height=10
        )
        
        # Define headings
        self.bulk_tree.heading('row', text='Row')
        self.bulk_tree.heading('title', text='Title')
        self.bulk_tree.heading('sku', text='SKU')
        self.bulk_tree.heading('price', text='Price')
        self.bulk_tree.heading('images', text='Images')
        self.bulk_tree.heading('images_path', text='Images Path')
        self.bulk_tree.heading('status', text='Status')
        
        # Define columns
        self.bulk_tree.column('row', width=50)
        self.bulk_tree.column('title', width=200)
        self.bulk_tree.column('sku', width=100)
        self.bulk_tree.column('price', width=80)
        self.bulk_tree.column('images', width=70)
        self.bulk_tree.column('images_path', width=200)
        self.bulk_tree.column('status', width=100)
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.bulk_tree.yview)
        self.bulk_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.bulk_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.grid(row=0, column=1, sticky="ns")
        
        # Statistics Frame
        row += 1
        stats_frame = ttk.LabelFrame(self.bulk_frame, text="Batch Statistics", padding=10)
        stats_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=10, pady=10)
        
        self.bulk_stats_var = tk.StringVar()
        self.bulk_stats_var.set("No products loaded")
        stats_label = ttk.Label(stats_frame, textvariable=self.bulk_stats_var, justify=tk.LEFT)
        stats_label.grid(row=0, column=0, sticky="w")
        
        # Bulk Upload Controls
        row += 1
        upload_frame = ttk.Frame(self.bulk_frame)
        upload_frame.grid(row=row, column=0, columnspan=3, pady=20)
        
        self.upload_bulk_btn = ttk.Button(
            upload_frame,
            text="🚀 Upload All to Queue",
            command=self.queue_bulk_products,
            style="Accent.TButton"
        )
        self.upload_bulk_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(upload_frame, text="🗑️ Clear List",
                command=self.clear_bulk_list).pack(side=tk.LEFT, padx=5)
        
        # Configure row/column weights for expansion
        self.bulk_frame.rowconfigure(row-2, weight=1)  # Make preview frame expandable
        self.bulk_frame.columnconfigure(1, weight=1)

    def select_excel_file(self):
        """Select Excel file for bulk upload"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsm"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.excel_file_var.set(file_path)
            self.log_message(f"Selected Excel file: {file_path}")

    def load_excel_file(self):
        """Load and validate Excel file"""
        excel_path = self.excel_file_var.get()
        
        if not excel_path:
            messagebox.showwarning("Warning", "Please select an Excel file first")
            return
        
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", "Selected Excel file does not exist")
            return
        
        # Clear existing products
        self.bulk_products = []
        for item in self.bulk_tree.get_children():
            self.bulk_tree.delete(item)
        
        # Show progress
        self.log_message(f"Loading Excel file: {excel_path}")
        self.file_info_var.set(f"Loading {os.path.basename(excel_path)}...")
        
        def load_thread():
            try:
                # Read products from Excel
                products, stats = self.excel_processor.read_excel_file(excel_path)
                
                if not products:
                    error_msg = "No valid products found in Excel file."
                    if stats.get('errors'):
                        error_msg += f"\n\nErrors:\n" + "\n".join(stats['errors'][:5])
                    
                    self.root.after(0, lambda: messagebox.showerror(
                        "No Products", 
                        error_msg
                    ))
                    self.root.after(0, lambda: self.file_info_var.set(
                        f"Failed to load: {os.path.basename(excel_path)}"
                    ))
                    return
                
                # Update UI with products
                self.root.after(0, lambda: self._update_bulk_tree_excel(products))
                
                # Update statistics
                total_images = sum(p.get('image_count', 0) for p in products)
                stats_text = f"""
    📊 EXCEL LOAD COMPLETE
    ──────────────────────────────────
    File: {os.path.basename(excel_path)}
    Total Products: {len(products)}
    Total Images: {total_images}
    Products with Images: {sum(1 for p in products if p.get('has_images'))}
    Products without Images: {sum(1 for p in products if not p.get('has_images'))}
    ──────────────────────────────────
    """
                self.root.after(0, lambda: self.bulk_stats_var.set(stats_text))
                self.root.after(0, lambda: self.file_info_var.set(
                    f"Loaded: {os.path.basename(excel_path)} ({len(products)} products)"
                ))
                
                # Store products
                self.bulk_products = products
                
                # Log results
                self.root.after(0, lambda: self.log_message(
                    f"Loaded {len(products)} products from Excel"
                ))
                
                # Show success message
                self.root.after(0, lambda: messagebox.showinfo(
                    "Excel Load Complete",
                    f"Successfully loaded {len(products)} products from Excel.\n\n"
                    f"Total images found: {total_images}\n"
                    f"Review the products and click 'Upload All to Queue' when ready."
                ))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Load Error",
                    f"Error loading Excel file: {str(e)}"
                ))
                self.root.after(0, lambda: self.log_message(f"Excel load error: {e}", "error"))
                self.root.after(0, lambda: self.file_info_var.set(
                    f"Load failed: {str(e)[:50]}..."
                ))
        
        # Run load in separate thread
        threading.Thread(target=load_thread, daemon=True).start()

    def _update_bulk_tree_excel(self, products):
        """Update the bulk treeview with products from Excel"""
        # Clear existing items
        for item in self.bulk_tree.get_children():
            self.bulk_tree.delete(item)
        
        # Add products to treeview
        for product in products:
            status = "✅ Ready" if product.get('has_images') else "⚠️ No Images"
            images_path = product.get('images_path', '')
            # Truncate long paths for display
            if len(images_path) > 40:
                images_path = images_path[:20] + "..." + images_path[-20:]
            
            self.bulk_tree.insert(
                '', 'end',
                values=(
                    product.get('excel_row', ''),
                    product.get('title', '')[:40],
                    product.get('sku', ''),
                    f"${product.get('price', '0')}",
                    product.get('image_count', 0),
                    images_path,
                    status
                ),
                tags=('has_images' if product.get('has_images') else 'no_images',)
            )
        
        # Configure tags for coloring
        self.bulk_tree.tag_configure('has_images', foreground='green')
        self.bulk_tree.tag_configure('no_images', foreground='orange')

    def create_excel_template(self):
        """Create an Excel template file"""
        file_path = filedialog.asksaveasfilename(
            title="Save Excel Template",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="product_template.xlsx"
        )
        
        if file_path:
            try:
                success = self.excel_processor.export_to_excel_template(file_path)
                if success:
                    messagebox.showinfo(
                        "Template Created",
                        f"Excel template created successfully!\n\n"
                        f"Location: {file_path}\n\n"
                        f"The template includes:\n"
                        f"• Sample data\n"
                        f"• Instructions sheet\n"
                        f"• Required columns: title, description, price, sku, images_path"
                    )
                    self.log_message(f"Created Excel template: {file_path}")
                else:
                    messagebox.showerror("Error", "Failed to create template")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create template: {str(e)}")

    def setup_queue_frame(self):
        """Setup the queue management tab"""
        # Configure queue frame
        for i in range(3):
            self.queue_frame.columnconfigure(i, weight=1)
        
        # Queue status
        ttk.Label(self.queue_frame, text="Upload Queue Status", 
                 font=('TkDefaultFont', 12, 'bold')).grid(
            row=0, column=0, columnspan=3, pady=(10, 20)
        )
        
        # Stats frame
        stats_frame = ttk.LabelFrame(self.queue_frame, text="Statistics", padding=10)
        stats_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=5)
        
        self.stats_vars = {
            'queue_size': tk.StringVar(value="0"),
            'active_workers': tk.StringVar(value="0"),
            'completed': tk.StringVar(value="0"),
            'failed': tk.StringVar(value="0"),
            'total_bulk': tk.StringVar(value="0")
        }
        
        ttk.Label(stats_frame, text="Items in queue:").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Label(stats_frame, textvariable=self.stats_vars['queue_size']).grid(row=0, column=1, sticky="w", padx=5)
        
        ttk.Label(stats_frame, text="Active uploads:").grid(row=0, column=2, sticky="w", padx=5)
        ttk.Label(stats_frame, textvariable=self.stats_vars['active_workers']).grid(row=0, column=3, sticky="w", padx=5)
        
        ttk.Label(stats_frame, text="Completed:").grid(row=1, column=0, sticky="w", padx=5)
        ttk.Label(stats_frame, textvariable=self.stats_vars['completed']).grid(row=1, column=1, sticky="w", padx=5)
        
        ttk.Label(stats_frame, text="Failed:").grid(row=1, column=2, sticky="w", padx=5)
        ttk.Label(stats_frame, textvariable=self.stats_vars['failed']).grid(row=1, column=3, sticky="w", padx=5)
        
        ttk.Label(stats_frame, text="Bulk in queue:").grid(row=2, column=0, sticky="w", padx=5)
        ttk.Label(stats_frame, textvariable=self.stats_vars['total_bulk']).grid(row=2, column=1, sticky="w", padx=5)
        
        # Queue control buttons
        control_frame = ttk.Frame(self.queue_frame)
        control_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        ttk.Button(control_frame, text="⏸️ Pause Queue", 
                  command=self.pause_queue).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="▶️ Resume Queue", 
                  command=self.resume_queue).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="🗑️ Clear Completed", 
                  command=self.clear_completed).pack(side=tk.LEFT, padx=5)
        
        # Upload history
        history_frame = ttk.LabelFrame(self.queue_frame, text="Upload History", padding=10)
        history_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=10, pady=10)
        history_frame.columnconfigure(0, weight=1)
        history_frame.rowconfigure(0, weight=1)
        
        # Treeview for history
        columns = ('time', 'title', 'status', 'id', 'details', 'type')
        self.history_tree = ttk.Treeview(
            history_frame, 
            columns=columns,
            show='headings',
            height=15
        )
        
        # Define headings
        self.history_tree.heading('time', text='Time')
        self.history_tree.heading('title', text='Product Title')
        self.history_tree.heading('status', text='Status')
        self.history_tree.heading('id', text='Product ID')
        self.history_tree.heading('details', text='Details')
        self.history_tree.heading('type', text='Type')
        
        # Define columns
        self.history_tree.column('time', width=100)
        self.history_tree.column('title', width=200)
        self.history_tree.column('status', width=80)
        self.history_tree.column('id', width=70)
        self.history_tree.column('details', width=250)
        self.history_tree.column('type', width=80)
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.history_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.grid(row=0, column=1, sticky="ns")
        
        # Configure weights for expansion
        self.queue_frame.rowconfigure(3, weight=1)
    
    # ==========================
    # BULK UPLOAD METHODS
    # ==========================
    
    def select_bulk_directory(self):
        """Select directory containing product folders"""
        directory = filedialog.askdirectory(
            title="Select Directory with Product Folders"
        )
        
        if directory:
            self.bulk_dir_var.set(directory)
            self.bulk_directory = directory
            self.log_message(f"Selected bulk directory: {directory}")
    
    def scan_bulk_directory(self):
        """Scan the selected directory for product folders"""
        directory = self.bulk_dir_var.get()
        
        if not directory:
            messagebox.showwarning("Warning", "Please select a directory first")
            return
        
        if not os.path.exists(directory):
            messagebox.showerror("Error", "Selected directory does not exist")
            return
        
        # Clear existing products
        self.bulk_products = []
        for item in self.bulk_tree.get_children():
            self.bulk_tree.delete(item)
        
        # Show progress
        self.log_message(f"Scanning directory: {directory}")
        self.bulk_stats_var.set("Scanning directory... Please wait.")
        
        def scan_thread():
            try:
                # Scan directory using bulk processor
                products = self.bulk_processor.scan_directory(directory)
                
                if not products:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "No Products", 
                        "No valid product folders found in the selected directory."
                    ))
                    self.root.after(0, lambda: self.bulk_stats_var.set(
                        "No valid product folders found."
                    ))
                    return
                
                # Validate products
                validation = self.bulk_processor.validate_products(products)
                valid_products = validation['valid_products']
                stats = validation['stats']
                
                # Update UI with products
                self.root.after(0, lambda: self._update_bulk_tree(valid_products))
                
                # Update statistics
                total_images = sum(p.get('image_count', 0) for p in valid_products)
                stats_text = f"""
📊 SCAN COMPLETE
────────────────────────
Total Products: {len(valid_products)}
Total Images: {total_images}
Products with Images: {stats['with_images']}
Products without Images: {stats['without_images']}
────────────────────────
"""
                self.root.after(0, lambda: self.bulk_stats_var.set(stats_text))
                
                # Store products
                self.bulk_products = valid_products
                
                # Log results
                self.root.after(0, lambda: self.log_message(
                    f"Scanned {len(valid_products)} products from directory"
                ))
                
                # Show summary
                if stats['errors']:
                    error_count = len(stats['errors'])
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Scan Complete with Warnings",
                        f"Found {len(valid_products)} valid products.\n"
                        f"{error_count} warnings (see log for details)."
                    ))
                else:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Scan Complete",
                        f"Successfully scanned {len(valid_products)} products."
                    ))
                    
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Scan Error",
                    f"Error scanning directory: {str(e)}"
                ))
                self.root.after(0, lambda: self.log_message(f"Scan error: {e}", "error"))
                self.root.after(0, lambda: self.bulk_stats_var.set(
                    f"Scan failed: {str(e)}"
                ))
        
        # Run scan in separate thread
        threading.Thread(target=scan_thread, daemon=True).start()
    
    def _update_bulk_tree(self, products):
        """Update the bulk treeview with scanned products"""
        # Clear existing items
        for item in self.bulk_tree.get_children():
            self.bulk_tree.delete(item)
        
        # Add products to treeview
        for product in products:
            status = "✅ Ready" if product.get('has_images') else "⚠️ No Images"
            
            self.bulk_tree.insert(
                '', 'end',
                values=(
                    product.get('folder_name', ''),
                    product.get('title', '')[:50],
                    product.get('sku', ''),
                    f"${product.get('price', '0')}",
                    product.get('image_count', 0),
                    status
                ),
                tags=('has_images' if product.get('has_images') else 'no_images',)
            )
        
        # Configure tags for coloring
        self.bulk_tree.tag_configure('has_images', foreground='green')
        self.bulk_tree.tag_configure('no_images', foreground='orange')
        
    def queue_bulk_products(self):
        """Add all loaded products to the upload queue"""
        if not self.bulk_products:
            messagebox.showwarning("Warning", "No products to upload. Please load an Excel file first.")
            return
        
        # Get selected category
        category_display = self.bulk_category_var.get()
        if not category_display:
            messagebox.showwarning("Warning", "Please select a category for all products")
            return
        
        category_id = self.category_dict.get(category_display)
        if not category_id:
            messagebox.showerror("Error", "Invalid category selected")
            return
        
        # Confirm bulk upload
        product_count = len(self.bulk_products)
        total_images = sum(p.get('image_count', 0) for p in self.bulk_products)
        
        confirm = messagebox.askyesno(
            "Confirm Bulk Upload",
            f"Add {product_count} products to upload queue?\n\n"
            f"• Total images: {total_images}\n"
            f"• Category: {category_display.split('(')[0].strip()}\n"
            f"• Products without images: {sum(1 for p in self.bulk_products if not p.get('has_images'))}\n\n"
            f"Upload will run in background. Continue?"
        )
        
        if not confirm:
            return
        
        # Create batch ID
        batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Add each product to queue
        added_count = 0
        skipped_count = 0
        
        for product in self.bulk_products:
            # Check if product has images
            if not product.get('has_images'):
                # Ask user preference
                upload_without_images = messagebox.askyesno(
                    "Missing Images",
                    f"Product '{product.get('title')}' has no images.\n\n"
                    f"Upload without images?"
                )
                
                if not upload_without_images:
                    skipped_count += 1
                    continue
            
            # Prepare product data for queue
            queue_data = {
                'title': product.get('title', ''),
                'description': product.get('description', ''),
                'price': product.get('price', '0'),
                'category_id': category_id,
                'images': product.get('images', []),
                'sku': product.get('sku', ''),
                'batch_id': batch_id,
                'excel_row': product.get('excel_row', ''),
                'timestamp': datetime.now().isoformat()
            }
            
            # Add to queue
            self.queue_manager.add_to_queue(queue_data)
            added_count += 1
            
            # Update product status in treeview
            self._update_product_status_excel(product.get('excel_row'), "⏳ Queued")
        
        # Update statistics
        bulk_in_queue = int(self.stats_vars['total_bulk'].get()) + added_count
        self.stats_vars['total_bulk'].set(str(bulk_in_queue))
        
        # Log results
        self.log_message(
            f"Excel bulk upload: Added {added_count} products to queue. "
            f"Skipped {skipped_count} products."
        )
        
        messagebox.showinfo(
            "Bulk Upload Started",
            f"✅ Added {added_count} products to upload queue.\n\n"
            f"• Uploading to category: {category_display.split('(')[0].strip()}\n"
            f"• Batch ID: {batch_id}\n"
            f"• Skipped products: {skipped_count}\n\n"
            f"Upload will run in background. Check the Queue tab for progress."
        )
        
        # Clear bulk list after adding to queue
        self.clear_bulk_list()

    def _update_product_status_excel(self, row_number, status):
        """Update status of a product in the bulk treeview (Excel version)"""
        for item in self.bulk_tree.get_children():
            values = self.bulk_tree.item(item, 'values')
            if values and values[0] == str(row_number):
                # Update status column (index 6)
                new_values = list(values)
                new_values[6] = status
                self.bulk_tree.item(item, values=new_values)
                break    
    def _update_product_status(self, folder_name, status):
        """Update status of a product in the bulk treeview"""
        for item in self.bulk_tree.get_children():
            values = self.bulk_tree.item(item, 'values')
            if values and values[0] == folder_name:
                # Update status column (index 5)
                new_values = list(values)
                new_values[5] = status
                self.bulk_tree.item(item, values=new_values)
                break
    
    def export_bulk_products(self):
        """Export scanned products to CSV"""
        if not self.bulk_products:
            messagebox.showwarning("Warning", "No products to export")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Products to CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.bulk_processor.export_products_to_csv(self.bulk_products, file_path)
                messagebox.showinfo("Export Successful", f"Products exported to:\n{file_path}")
                self.log_message(f"Exported {len(self.bulk_products)} products to CSV")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export: {str(e)}")
    
    def clear_bulk_list(self):
        """Clear the bulk products list"""
        self.bulk_products = []
        for item in self.bulk_tree.get_children():
            self.bulk_tree.delete(item)
        
        self.bulk_stats_var.set("No products scanned")
        self.log_message("Bulk list cleared")
    
    def copy_bulk_summary(self):
        """Copy bulk summary to clipboard"""
        if not self.bulk_products:
            messagebox.showwarning("Warning", "No products to summarize")
            return
        
        summary = self.bulk_processor.create_batch_summary(self.bulk_products)
        self.root.clipboard_clear()
        self.root.clipboard_append(summary)
        self.root.update()  # Keep the clipboard content after the window is closed
        
        messagebox.showinfo("Summary Copied", "Batch summary copied to clipboard!")
    
    def _on_upload_complete(self, result):
        """Callback for when an upload completes"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if result.get('success'):
            product_id = result.get('data', {}).get('id', 'N/A')
            title = result.get('task', {}).get('title', 'Unknown')
            batch_id = result.get('task', {}).get('batch_id', '')
            
            # Determine upload type
            upload_type = "Bulk" if batch_id else "Single"
            
            # Add to history tree
            self.root.after(0, lambda: self.history_tree.insert(
                '', 'end',
                values=(
                    timestamp,
                    title[:40] + "..." if len(title) > 40 else title,
                    "✅ Success",
                    product_id,
                    f"Batch: {batch_id}" if batch_id else "Single upload",
                    upload_type
                )
            ))
            
            # Update stats
            completed = int(self.stats_vars['completed'].get()) + 1
            self.stats_vars['completed'].set(str(completed))
            
            # Update bulk count if it's a bulk upload
            if batch_id:
                current_bulk = int(self.stats_vars['total_bulk'].get())
                if current_bulk > 0:
                    self.stats_vars['total_bulk'].set(str(current_bulk - 1))
            
            self.root.after(0, lambda: self.log_message(
                f"Upload successful: {title} (ID: {product_id})"
            ))
        else:
            title = result.get('task', {}).get('title', 'Unknown')
            error = result.get('error', 'Unknown error')
            batch_id = result.get('task', {}).get('batch_id', '')
            upload_type = "Bulk" if batch_id else "Single"
            
            # Add to history tree
            self.root.after(0, lambda: self.history_tree.insert(
                '', 'end',
                values=(
                    timestamp,
                    title[:40] + "..." if len(title) > 40 else title,
                    "❌ Failed",
                    "N/A",
                    error[:80],
                    upload_type
                )
            ))
            
            # Update stats
            failed = int(self.stats_vars['failed'].get()) + 1
            self.stats_vars['failed'].set(str(failed))
            
            self.root.after(0, lambda: self.log_message(
                f"Upload failed: {title} - {error}", "error"
            ))
        
        # Update queue stats
        self.update_stats()
    
    def update_queue_status(self):
        """Periodically update queue status"""
        queue_size = self.queue_manager.get_queue_size()
        active_workers = self.queue_manager.get_active_workers()
        
        self.queue_status_var.set(
            f"Queue: {queue_size} items waiting | {active_workers} active uploads"
        )
        
        # Update stats in queue tab
        self.stats_vars['queue_size'].set(str(queue_size))
        self.stats_vars['active_workers'].set(str(active_workers))
        
        # Schedule next update
        self.root.after(1000, self.update_queue_status)
    
    # ==========================
    # EXISTING SINGLE PRODUCT METHODS
    # ==========================
    
    def log_message(self, message, level="info"):
        """Add message to status bar and log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_var.set(f"{timestamp}: {message}")
        
        if level == "error":
            print(f"ERROR: {message}")
        else:
            print(f"INFO: {message}")
    
    def load_categories(self):
        """Load categories from WooCommerce"""
        def worker():
            self.log_message("Loading categories...")
            try:
                categories = self.wc_api.get_categories()
                if categories:
                    self.categories = categories
                    category_list, category_tree = build_category_tree(categories)
                    
                    # Update both comboboxes in main thread
                    self.root.after(0, lambda: self._update_category_combos(category_list))
                    self.log_message(f"Loaded {len(category_list)} categories")
                else:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Warning", "No categories found or failed to load"
                    ))
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Error loading categories: {e}", "error"))
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _update_category_combos(self, category_list):
        """Update both category comboboxes"""
        self.category_combo['values'] = category_list
        self.bulk_category_combo['values'] = category_list
        
        # Build mapping for display names to IDs
        self.category_dict = {}
        for cat in self.categories:
            for display_name in category_list:
                if f"(ID: {cat['id']})" in display_name:
                    self.category_dict[display_name] = cat['id']
                    break
    
    def add_images(self):
        """Add images to the list"""
        files = filedialog.askopenfilenames(
            title="Select Product Images",
            filetypes=[
                ("Image files", "*.jpg *.jpeg *.png *.gif *.webp *.bmp *.tiff"),
                ("All files", "*.*")
            ]
        )
        
        for file in files:
            if file not in self.images:  # Avoid duplicates
                self.images.append(file)
                self.image_listbox.insert(tk.END, os.path.basename(file))
        
        if files:
            self.log_message(f"Added {len(files)} image(s)")
    
    def remove_image(self):
        """Remove selected image from list"""
        selection = self.image_listbox.curselection()
        if selection:
            index = selection[0]
            removed = self.images.pop(index)
            self.image_listbox.delete(index)
            self.log_message(f"Removed: {os.path.basename(removed)}")
    
    def move_image(self, direction):
        """Move image up or down in the list"""
        selection = self.image_listbox.curselection()
        if selection:
            index = selection[0]
            new_index = index + direction
            
            if 0 <= new_index < len(self.images):
                # Swap in images list
                self.images[index], self.images[new_index] = self.images[new_index], self.images[index]
                
                # Update listbox
                items = list(self.image_listbox.get(0, tk.END))
                items[index], items[new_index] = items[new_index], items[index]
                
                self.image_listbox.delete(0, tk.END)
                for item in items:
                    self.image_listbox.insert(tk.END, item)
                
                self.image_listbox.selection_set(new_index)
                self.log_message("Image order updated")
    
    def queue_product_upload(self):
        """Add single product to the upload queue"""
        # Validate inputs
        if not self.title_var.get().strip():
            messagebox.showwarning("Warning", "Please enter a product title")
            return
        
        try:
            price = float(self.price_var.get())
            if price < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a valid price (positive number)")
            return
        
        if not self.images:
            upload_without_images = messagebox.askyesno(
                "No Images",
                "No images selected. Upload product without images?"
            )
            if not upload_without_images:
                return
        
        category_display = self.category_var.get()
        if not category_display:
            messagebox.showwarning("Warning", "Please select a category")
            return
        
        category_id = self.category_dict.get(category_display)
        if not category_id:
            messagebox.showerror("Error", "Invalid category selected")
            return
        
        # Prepare product data for the queue
        product_data = {
            'title': self.title_var.get().strip(),
            'description': self.desc_text.get("1.0", tk.END).strip(),
            'price': format_price(self.price_var.get()),
            'category_id': category_id,
            'images': self.images.copy(),
            'timestamp': datetime.now().isoformat()
        }
        
        # Add to queue
        queue_size = self.queue_manager.add_to_queue(product_data)
        
        # Update status
        self.queue_status_var.set(
            f"Product added to queue. {queue_size} item(s) waiting | {self.queue_manager.get_active_workers()} active"
        )
        self.log_message(f"Product '{product_data['title']}' added to upload queue")
        
        # Clear form for next entry
        self.clear_form()
        
        # Update stats
        self.update_stats()
    
    def clear_form(self):
        """Clear all form fields"""
        self.title_var.set("")
        self.desc_text.delete("1.0", tk.END)
        self.price_var.set("")
        self.category_var.set("")
        self.images.clear()
        self.image_listbox.delete(0, tk.END)
        self.log_message("Form cleared")
    
    def test_connection(self):
        """Test connection to WooCommerce API"""
        def worker():
            self.log_message("Testing connection...")
            try:
                if self.wc_api.test_connection():
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Connection Test", 
                        "✅ Successfully connected to WooCommerce API!"
                    ))
                    self.root.after(0, lambda: self.log_message("Connection test successful"))
                else:
                    self.root.after(0, lambda: messagebox.showerror(
                        "Connection Test", 
                        "❌ Failed to connect to WooCommerce API. Check your credentials."
                    ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Connection Test", 
                    f"Error: {str(e)}"
                ))
        
        threading.Thread(target=worker, daemon=True).start()
    
    def generate_ai_title(self):
        """Generate product title using AI"""
        if not self.ai_helper.available:
            messagebox.showinfo("AI Not Available", 
                              "OpenAI library not installed or API key not set.\n"
                              "Install with: pip install openai\n"
                              "And set OPENAI_API_KEY in .env file")
            return
        
        # Get some context
        prompt = self.title_var.get().strip()
        if not prompt:
            # Use tkinter.simpledialog
            try:
                import tkinter.simpledialog as simpledialog
                prompt = simpledialog.askstring(
                    "AI Title Generation", 
                    "Enter a brief description for title generation:"
                )
                if not prompt:
                    return
            except ImportError:
                # Fallback to custom dialog
                dialog = tk.Toplevel(self.root)
                dialog.title("AI Title Generation")
                dialog.geometry("300x100")
                
                tk.Label(dialog, text="Enter a brief description:").pack(pady=10)
                prompt_entry = ttk.Entry(dialog, width=40)
                prompt_entry.pack(pady=5)
                
                def get_prompt():
                    nonlocal prompt
                    prompt = prompt_entry.get()
                    dialog.destroy()
                
                ttk.Button(dialog, text="Generate", command=get_prompt).pack(pady=5)
                dialog.wait_window()
                if not prompt:
                    return
        
        def worker():
            try:
                titles = self.ai_helper.generate_title(prompt, num_titles=3)
                if titles:
                    self.root.after(0, lambda: self._show_ai_titles(titles))
                else:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "No Titles", 
                        "Could not generate titles. Try again."
                    ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "AI Error", 
                    f"Failed to generate title: {str(e)}"
                ))
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _show_ai_titles(self, titles):
        """Show AI-generated titles in a dialog"""
        title_dialog = tk.Toplevel(self.root)
        title_dialog.title("Select AI-Generated Title")
        title_dialog.geometry("400x300")
        
        tk.Label(title_dialog, text="Choose a title:", 
                font=('TkDefaultFont', 11, 'bold')).pack(pady=10)
        
        for i, title in enumerate(titles, 1):
            frame = ttk.Frame(title_dialog)
            frame.pack(fill="x", padx=20, pady=2)
            
            ttk.Label(frame, text=f"{i}.", width=3).pack(side=tk.LEFT)
            title_text = tk.Text(frame, height=2, width=40)
            title_text.pack(side=tk.LEFT, padx=5)
            title_text.insert("1.0", title)
            title_text.config(state=tk.DISABLED)
            
            ttk.Button(frame, text="Use", 
                      command=lambda t=title: self._select_ai_title(t, title_dialog)).pack(side=tk.LEFT)
    
    def _select_ai_title(self, title, dialog):
        """Select an AI-generated title"""
        self.title_var.set(title)
        dialog.destroy()
        self.log_message("AI title applied")
    
    def generate_ai_description(self):
        """Generate product description using AI"""
        if not self.ai_helper.available:
            messagebox.showinfo("AI Not Available", 
                              "OpenAI library not installed or API key not set.\n"
                              "Install with: pip install openai\n"
                              "And set OPENAI_API_KEY in .env file")
            return
        
        product_title = self.title_var.get().strip()
        if not product_title:
            messagebox.showwarning("Input Required", "Please enter a product title first")
            return
        
        def worker():
            try:
                description = self.ai_helper.generate_description(product_title)
                if description:
                    self.root.after(0, lambda: self._apply_ai_description(description))
                else:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "No Description", 
                        "Could not generate description. Try again."
                    ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "AI Error", 
                    f"Failed to generate description: {str(e)}"
                ))
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _apply_ai_description(self, description):
        """Apply AI-generated description"""
        self.desc_text.delete("1.0", tk.END)
        self.desc_text.insert("1.0", description)
        self.log_message("AI description generated")
    
    def pause_queue(self):
        """Pause the upload queue"""
        # In a real implementation, you would add pause/resume functionality
        # to the UploadQueueManager class
        self.log_message("Queue pause requested (feature to be implemented)")
    
    def resume_queue(self):
        """Resume the upload queue"""
        self.log_message("Queue resume requested (feature to be implemented)")
    
    def clear_completed(self):
        """Clear completed items from history"""
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        self.log_message("Upload history cleared")
    
    def on_closing(self):
        """Clean up when closing the application"""
        self.queue_manager.stop()
        self.root.destroy()
    
    def update_stats(self):
        """Update statistics display"""
        # This is called from various places, metrics are updated elsewhere
        pass


def main():
    """Main function"""
    # Create main window
    root = tk.Tk()
    
    # Set style
    style = ttk.Style()
    style.theme_use('clam')
    
    # Create app
    app = ProductUploaderApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Start main loop
    root.mainloop()


if __name__ == "__main__":
    main()