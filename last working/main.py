import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import json
from datetime import datetime
import mimetypes
import base64
from PIL import Image
import io

# Optional AI imports
try:
    import openai
    AI_AVAILABLE = True
except ImportError:
    AI_AVAILABLE = False

# Load environment variables
load_dotenv()

class WooCommerceProductUploader:
    def __init__(self, root):
        self.root = root
        self.root.title("WooCommerce Product Uploader")
        self.root.geometry("900x800")
        
        # Store configuration
        self.store_url = os.getenv('STORE_URL')
        self.wp_username = os.getenv('WP_USERNAME')
        self.wp_password = os.getenv('WP_APP_PASSWORD')
        self.wc_consumer_key = os.getenv('WC_CONSUMER_KEY')
        self.wc_consumer_secret = os.getenv('WC_CONSUMER_SECRET')
        
        # WooCommerce API endpoints
        self.wc_api_base = f"{self.store_url}/wp-json/wc/v3"
        
        # Variables
        self.categories = []
        self.images = []
        
        # Setup UI
        self.setup_ui()
        
        # Load categories
        self.load_categories()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Product Title
        ttk.Label(main_frame, text="Product Title:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.title_var = tk.StringVar()
        self.title_entry = ttk.Entry(main_frame, textvariable=self.title_var, width=50)
        self.title_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # AI Title Generation Button (if AI available)
        if AI_AVAILABLE:
            self.ai_title_btn = ttk.Button(main_frame, text="AI Generate", 
                                          command=self.generate_ai_title)
            self.ai_title_btn.grid(row=0, column=2, padx=(5, 0), pady=5)
        
        # Description
        ttk.Label(main_frame, text="Description:").grid(row=1, column=0, sticky=tk.NW, pady=5)
        self.desc_text = tk.Text(main_frame, height=10, width=50)
        self.desc_text.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5, padx=(5, 0))
        
        # AI Description Generation Button
        if AI_AVAILABLE:
            self.ai_desc_btn = ttk.Button(main_frame, text="AI Generate", 
                                         command=self.generate_ai_description)
            self.ai_desc_btn.grid(row=1, column=2, padx=(5, 0), pady=5, sticky=tk.N)
        
        # Description scrollbar
        desc_scroll = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.desc_text.yview)
        desc_scroll.grid(row=1, column=3, sticky=(tk.N, tk.S), pady=5)
        self.desc_text['yscrollcommand'] = desc_scroll.set
        
        # Category
        ttk.Label(main_frame, text="Category:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(main_frame, textvariable=self.category_var, 
                                          width=47, state="readonly")
        self.category_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Images Section
        ttk.Label(main_frame, text="Product Images:").grid(row=3, column=0, sticky=tk.W, pady=5)
        
        # Image list frame
        image_frame = ttk.LabelFrame(main_frame, text="Selected Images", padding="10")
        image_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Image listbox with scrollbar
        self.image_listbox = tk.Listbox(image_frame, height=6)
        self.image_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        list_scroll = ttk.Scrollbar(image_frame, orient=tk.VERTICAL, command=self.image_listbox.yview)
        list_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.image_listbox['yscrollcommand'] = list_scroll.set
        
        # Image control buttons
        btn_frame = ttk.Frame(image_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(5, 0))
        
        ttk.Button(btn_frame, text="Add Images", command=self.add_images).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Remove Selected", command=self.remove_image).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Move Up", command=lambda: self.move_image(-1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Move Down", command=lambda: self.move_image(1)).pack(side=tk.LEFT, padx=2)
        
        # Note about first image
        ttk.Label(main_frame, text="Note: First image will be set as featured image", 
                 font=('TkDefaultFont', 9, 'italic')).grid(row=4, column=1, sticky=tk.W, pady=(0, 10))
        
        # Status/Log
        ttk.Label(main_frame, text="Log:").grid(row=5, column=0, sticky=tk.NW, pady=5)
        self.log_text = tk.Text(main_frame, height=10, width=50, state=tk.DISABLED)
        self.log_text.grid(row=5, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5, padx=(5, 0))
        
        log_scroll = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.grid(row=5, column=3, sticky=(tk.N, tk.S), pady=5)
        self.log_text['yscrollcommand'] = log_scroll.set
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Upload Product", command=self.upload_product).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Form", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Test Connection", command=self.test_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Refresh Categories", command=self.load_categories).pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=7, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def log_message(self, message):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.status_var.set(message)
    
    def load_categories(self):
        """Load categories from WooCommerce"""
        try:
            self.log_message("Loading categories...")
            
            # WooCommerce API v3 uses different authentication
            endpoint = f"{self.wc_api_base}/products/categories"
            params = {
                'per_page': 100,
                'hide_empty': False
            }
            
            response = requests.get(
                endpoint,
                auth=HTTPBasicAuth(self.wc_consumer_key, self.wc_consumer_secret),
                params=params
            )
            
            if response.status_code == 200:
                categories = response.json()
                self.categories = categories
                
                # Build category list with hierarchy
                category_list = []
                self.category_dict = {}
                
                def build_category_tree(cats, parent_id=0, level=0):
                    for cat in cats:
                        if cat['parent'] == parent_id:
                            indent = "  " * level
                            display_name = f"{indent}{cat['name']} (ID: {cat['id']})"
                            category_list.append(display_name)
                            self.category_dict[display_name] = cat['id']
                            build_category_tree(cats, cat['id'], level + 1)
                
                build_category_tree(categories)
                
                self.category_combo['values'] = category_list
                self.log_message(f"Loaded {len(category_list)} categories")
            else:
                self.log_message(f"Error loading categories: {response.status_code}")
                messagebox.showerror("Error", f"Failed to load categories: {response.status_code}")
                
        except Exception as e:
            self.log_message(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to load categories: {str(e)}")
    
    def add_images(self):
        """Add images to the list"""
        files = filedialog.askopenfilenames(
            title="Select Product Images",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.webp")]
        )
        
        for file in files:
            self.images.append(file)
            self.image_listbox.insert(tk.END, os.path.basename(file))
        
        if files:
            self.log_message(f"Added {len(files)} image(s)")
    
    def remove_image(self):
        """Remove selected image from list"""
        selection = self.image_listbox.curselection()
        if selection:
            index = selection[0]
            self.images.pop(index)
            self.image_listbox.delete(index)
            self.log_message("Image removed")
    
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
    
    def upload_media_to_wordpress(self, image_path):
        """Upload image to WordPress media library"""
        try:
            # Open and prepare image
            with open(image_path, 'rb') as img_file:
                img_data = img_file.read()
            
            # Get filename and mime type
            filename = os.path.basename(image_path)
            mime_type, _ = mimetypes.guess_type(image_path)
            
            # WordPress media endpoint
            media_url = f"{self.store_url}/wp-json/wp/v2/media"
            
            # Prepare headers
            headers = {
                'Content-Disposition': f'attachment; filename="{filename}"',
                'Content-Type': mime_type
            }
            
            # Upload using WordPress REST API with application password
            response = requests.post(
                media_url,
                headers=headers,
                data=img_data,
                auth=HTTPBasicAuth(self.wp_username, self.wp_password)
            )
            
            if response.status_code in [200, 201]:
                media_data = response.json()
                self.log_message(f"Uploaded: {filename}")
                return {
                    'id': media_data['id'],
                    'src': media_data['source_url']
                }
            else:
                self.log_message(f"Failed to upload {filename}: {response.status_code}")
                return None
                
        except Exception as e:
            self.log_message(f"Error uploading {image_path}: {str(e)}")
            return None
    
    def upload_product(self):
        """Upload product to WooCommerce"""
        # Validate inputs
        if not self.title_var.get().strip():
            messagebox.showwarning("Warning", "Please enter a product title")
            return
        
        if not self.images:
            messagebox.showwarning("Warning", "Please add at least one image")
            return
        
        category_display = self.category_var.get()
        if not category_display:
            messagebox.showwarning("Warning", "Please select a category")
            return
        
        # Get category ID
        category_id = self.category_dict.get(category_display)
        if not category_id:
            messagebox.showerror("Error", "Invalid category selected")
            return
        
        try:
            self.log_message("Starting product upload...")
            
            # Upload images first
            uploaded_images = []
            for i, image_path in enumerate(self.images):
                self.log_message(f"Uploading image {i+1}/{len(self.images)}...")
                media_info = self.upload_media_to_wordpress(image_path)
                if media_info:
                    uploaded_images.append(media_info)
                else:
                    self.log_message(f"Failed to upload image: {image_path}")
            
            if not uploaded_images:
                messagebox.showerror("Error", "Failed to upload any images")
                return
            
            # Prepare product data
            product_data = {
                'name': self.title_var.get().strip(),
                'description': self.desc_text.get("1.0", tk.END).strip(),
                'type': 'simple',
                'regular_price': '0',  # You might want to add price field
                'categories': [{'id': category_id}],
                'images': []
            }
            
            # Add images to product
            for i, img_info in enumerate(uploaded_images):
                image_data = {
                    'id': img_info['id']
                }
                if i == 0:  # First image is featured
                    product_data['images'].insert(0, image_data)
                else:
                    product_data['images'].append(image_data)
            
            # Create product using WooCommerce REST API
            endpoint = f"{self.wc_api_base}/products"
            
            response = requests.post(
                endpoint,
                auth=HTTPBasicAuth(self.wc_consumer_key, self.wc_consumer_secret),
                json=product_data,
                headers={'Content-Type': 'application/json'}
            )
            
            if response.status_code == 201:
                product = response.json()
                self.log_message(f"Product created successfully! ID: {product['id']}")
                messagebox.showinfo("Success", f"Product '{product['name']}' created successfully!\nID: {product['id']}")
                
                # Clear form after successful upload
                self.clear_form()
            else:
                self.log_message(f"Failed to create product: {response.status_code}")
                self.log_message(f"Response: {response.text}")
                messagebox.showerror("Error", f"Failed to create product: {response.status_code}\n{response.text}")
                
        except Exception as e:
            self.log_message(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to upload product: {str(e)}")
    
    def clear_form(self):
        """Clear all form fields"""
        self.title_var.set("")
        self.desc_text.delete("1.0", tk.END)
        self.category_var.set("")
        self.images.clear()
        self.image_listbox.delete(0, tk.END)
        self.log_message("Form cleared")
    
    def test_connection(self):
        """Test connection to WooCommerce API"""
        try:
            self.log_message("Testing connection...")
            
            # Test WordPress connection
            wp_test = requests.get(
                f"{self.store_url}/wp-json/wp/v2",
                auth=HTTPBasicAuth(self.wp_username, self.wp_password)
            )
            
            # Test WooCommerce connection
            wc_test = requests.get(
                f"{self.wc_api_base}/products",
                auth=HTTPBasicAuth(self.wc_consumer_key, self.wc_consumer_secret),
                params={'per_page': 1}
            )
            
            if wp_test.status_code == 200 and wc_test.status_code == 200:
                self.log_message("✓ Connection successful!")
                messagebox.showinfo("Connection Test", "Successfully connected to both WordPress and WooCommerce APIs!")
            else:
                self.log_message(f"✗ Connection failed: WP={wp_test.status_code}, WC={wc_test.status_code}")
                messagebox.showerror("Connection Test", 
                                   f"Connection failed:\nWordPress: {wp_test.status_code}\nWooCommerce: {wc_test.status_code}")
                
        except Exception as e:
            self.log_message(f"✗ Connection error: {str(e)}")
            messagebox.showerror("Connection Test", f"Connection error: {str(e)}")
    
    def generate_ai_title(self):
        """Generate product title using AI"""
        if not AI_AVAILABLE:
            messagebox.showinfo("AI Not Available", 
                              "OpenAI library not installed.\nInstall with: pip install openai")
            return
        
        # Get some context or let user enter a prompt
        prompt = self.title_var.get().strip()
        if not prompt:
            prompt = simpledialog.askstring("AI Title Generation", 
                                          "Enter a brief description for title generation:")
            if not prompt:
                return
        
        try:
            # You'll need to set your OpenAI API key
            api_key = os.getenv('OPENAI_API_KEY')
            if not api_key:
                api_key = simpledialog.askstring("OpenAI API Key", 
                                               "Enter your OpenAI API key:", show='*')
                if not api_key:
                    return
                os.environ['OPENAI_API_KEY'] = api_key
            
            openai.api_key = api_key
            
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a product title generator for e-commerce. Generate compelling product titles."},
                    {"role": "user", "content": f"Generate 3 product titles for: {prompt}"}
                ],
                max_tokens=100
            )
            
            titles = response.choices[0].message.content.strip().split('\n')
            titles = [t.replace('1.', '').replace('2.', '').replace('3.', '').strip() for t in titles if t.strip()]
            
            # Show titles in a dialog
            title_dialog = tk.Toplevel(self.root)
            title_dialog.title("Select AI-Generated Title")
            
            tk.Label(title_dialog, text="Choose a title:").pack(pady=10)
            
            for i, title in enumerate(titles[:3]):  # Show first 3
                btn = ttk.Button(title_dialog, text=title, 
                               command=lambda t=title: self.select_ai_title(t, title_dialog))
                btn.pack(pady=2, padx=20)
            
        except Exception as e:
            messagebox.showerror("AI Error", f"Failed to generate title: {str(e)}")
    
    def select_ai_title(self, title, dialog):
        """Select an AI-generated title"""
        self.title_var.set(title)
        dialog.destroy()
        self.log_message("AI title applied")
    
    def generate_ai_description(self):
        """Generate product description using AI"""
        if not AI_AVAILABLE:
            messagebox.showinfo("AI Not Available", 
                              "OpenAI library not installed.\nInstall with: pip install openai")
            return
        
        product_title = self.title_var.get().strip()
        if not product_title:
            messagebox.showwarning("Input Required", "Please enter a product title first")
            return
        
        try:
            api_key = os.getenv('OPENAI_API_KEY')
            if not api_key:
                api_key = simpledialog.askstring("OpenAI API Key", 
                                               "Enter your OpenAI API key:", show='*')
                if not api_key:
                    return
                os.environ['OPENAI_API_KEY'] = api_key
            
            openai.api_key = api_key
            
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a product description writer for e-commerce. Write SEO-friendly product descriptions."},
                    {"role": "user", "content": f"Write a detailed product description for: {product_title}\nInclude features, benefits, and specifications."}
                ],
                max_tokens=300
            )
            
            description = response.choices[0].message.content.strip()
            
            # Clear and insert new description
            self.desc_text.delete("1.0", tk.END)
            self.desc_text.insert("1.0", description)
            self.log_message("AI description generated")
            
        except Exception as e:
            messagebox.showerror("AI Error", f"Failed to generate description: {str(e)}")

# Install required packages function
def check_dependencies():
    """Check and install required packages"""
    import subprocess
    import sys
    
    required = [
        'requests',
        'python-dotenv',
        'Pillow'  # For image handling if needed
    ]
    
    missing = []
    for package in required:
        try:
            __import__(package.replace('-', '_'))
        except ImportError:
            missing.append(package)
    
    if missing:
        print(f"Missing packages: {', '.join(missing)}")
        response = input("Install missing packages? (y/n): ")
        if response.lower() == 'y':
            for package in missing:
                print(f"Installing {package}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print("Installation complete!")
    
    # Check for optional AI package
    try:
        import openai
        print("OpenAI package found - AI features enabled")
    except ImportError:
        print("OpenAI not installed - AI features disabled")
        print("To enable AI: pip install openai")

def main():
    """Main function"""
    # Check dependencies
    check_dependencies()
    
    # Create main window
    root = tk.Tk()
    
    # Set style
    style = ttk.Style()
    style.theme_use('clam')
    
    # Create app
    app = WooCommerceProductUploader(root)
    
    # Start main loop
    root.mainloop()

if __name__ == "__main__":
    main()