import pandas as pd
import os
import glob
from pathlib import Path
from typing import List, Dict, Any, Tuple
import logging

logger = logging.getLogger(__name__)

class ExcelProductProcessor:
    def __init__(self):
        self.supported_image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.tiff']
    
    def read_excel_file(self, excel_path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
        """
        Read products from Excel file with specified columns
        
        Expected columns:
        - title
        - description
        - price
        - sku
        - images_path
        
        Returns: (list of products, stats dictionary)
        """
        products = []
        stats = {
            'total_rows': 0,
            'valid_products': 0,
            'invalid_products': 0,
            'products_with_images': 0,
            'products_without_images': 0,
            'errors': [],
            'columns_found': [],
            'columns_missing': []
        }
        
        try:
            # Read the Excel file
            logger.info(f"Reading Excel file: {excel_path}")
            
            # Try different engines for different file formats
            try:
                df = pd.read_excel(excel_path, engine='openpyxl')
            except:
                try:
                    df = pd.read_excel(excel_path, engine='xlrd')
                except Exception as e:
                    logger.error(f"Failed to read Excel file: {e}")
                    stats['errors'].append(f"Failed to read Excel file: {str(e)}")
                    return [], stats
            
            stats['total_rows'] = len(df)
            logger.info(f"Found {stats['total_rows']} rows in Excel file")
            
            # Normalize column names (strip whitespace, lowercase)
            df.columns = df.columns.str.strip().str.lower()
            
            # Check for required columns
            required_columns = ['title', 'description', 'price', 'images_path']
            optional_columns = ['sku']
            
            stats['columns_found'] = list(df.columns)
            stats['columns_missing'] = [col for col in required_columns if col not in df.columns]
            
            if stats['columns_missing']:
                error_msg = f"Missing required columns: {', '.join(stats['columns_missing'])}"
                logger.error(error_msg)
                stats['errors'].append(error_msg)
                return [], stats
            
            # Process each row
            for index, row in df.iterrows():
                try:
                    product_data = self._process_excel_row(row, index + 2)  # +2 for header and 1-based index
                    if product_data:
                        products.append(product_data)
                        stats['valid_products'] += 1
                        if product_data.get('has_images'):
                            stats['products_with_images'] += 1
                        else:
                            stats['products_without_images'] += 1
                    else:
                        stats['invalid_products'] += 1
                        stats['errors'].append(f"Row {index + 2}: Failed to process")
                        
                except Exception as e:
                    stats['invalid_products'] += 1
                    stats['errors'].append(f"Row {index + 2}: Error - {str(e)}")
                    logger.error(f"Error processing row {index + 2}: {e}")
            
            logger.info(f"Successfully processed {stats['valid_products']} products from Excel")
            return products, stats
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {e}")
            stats['errors'].append(f"Error reading Excel file: {str(e)}")
            return [], stats
    
    def _process_excel_row(self, row: pd.Series, row_num: int) -> Dict[str, Any]:
        """Process a single Excel row"""
        try:
            # Extract and validate required fields
            title = str(row['title']).strip()
            description = str(row['description']).strip()
            price_str = str(row['price']).strip()
            
            # Validate required fields
            if not title:
                logger.warning(f"Row {row_num}: Missing title")
                return None
            
            if not price_str:
                logger.warning(f"Row {row_num}: Missing price")
                return None
            
            # Process price
            price = self._validate_price(price_str)
            if price == "0":
                logger.warning(f"Row {row_num}: Invalid price format: {price_str}")
            
            # Process SKU (optional)
            sku = ""
            if 'sku' in row and pd.notna(row['sku']):
                sku = str(row['sku']).strip()
            
            # Process images path
            images_path = str(row['images_path']).strip()
            images = self._get_images_from_path(images_path)
            
            # Create product data dictionary
            product_data = {
                'row_number': row_num,
                'title': title,
                'description': description,
                'price': price,
                'sku': sku,
                'images': images,
                'has_images': len(images) > 0,
                'image_count': len(images),
                'images_path': images_path,
                'status': 'pending',
                'excel_row': row_num
            }
            
            return product_data
            
        except Exception as e:
            logger.error(f"Error processing row {row_num}: {e}")
            return None
    
    def _get_images_from_path(self, images_path: str) -> List[str]:
        """Get images from a path (could be file, directory, or list of files)"""
        images = []
        
        if not images_path or pd.isna(images_path):
            return images
        
        images_path = str(images_path).strip()
        
        try:
            # Case 1: Multiple paths separated by semicolon or comma
            if ';' in images_path or ',' in images_path:
                separators = [';', ',']
                for sep in separators:
                    if sep in images_path:
                        paths = [p.strip() for p in images_path.split(sep) if p.strip()]
                        for path in paths:
                            if os.path.exists(path):
                                if os.path.isfile(path) and self._is_image_file(path):
                                    images.append(path)
                                elif os.path.isdir(path):
                                    images.extend(self._get_images_from_directory(path))
                        break
            
            # Case 2: Single file path
            elif os.path.exists(images_path):
                if os.path.isfile(images_path) and self._is_image_file(images_path):
                    images.append(images_path)
                
                # Case 3: Directory path
                elif os.path.isdir(images_path):
                    images.extend(self._get_images_from_directory(images_path))
            
            # Case 4: Wildcard pattern (e.g., C:/images/product*.jpg)
            else:
                # Try as a wildcard pattern
                matched_files = glob.glob(images_path, recursive=True)
                for file_path in matched_files:
                    if os.path.isfile(file_path) and self._is_image_file(file_path):
                        images.append(file_path)
            
            # Sort images for consistency
            images.sort()
            return images
            
        except Exception as e:
            logger.error(f"Error getting images from path '{images_path}': {e}")
            return []
    
    def _get_images_from_directory(self, directory_path: str) -> List[str]:
        """Get all images from a directory"""
        images = []
        try:
            directory = Path(directory_path)
            if directory.exists() and directory.is_dir():
                for ext in self.supported_image_extensions:
                    pattern = f"*{ext}"
                    for image_path in directory.rglob(pattern):
                        images.append(str(image_path))
            
            # Sort by filename
            images.sort()
            return images
            
        except Exception as e:
            logger.error(f"Error getting images from directory {directory_path}: {e}")
            return []
    
    def _is_image_file(self, file_path: str) -> bool:
        """Check if file is an image based on extension"""
        ext = os.path.splitext(file_path)[1].lower()
        return ext in self.supported_image_extensions
    
    def _validate_price(self, price_str: str) -> str:
        """Validate and format price"""
        try:
            # Remove any currency symbols, commas, and whitespace
            import re
            cleaned = re.sub(r'[^\d.]', '', price_str)
            if not cleaned:
                return "0"
            
            # Convert to float and format
            price_float = float(cleaned)
            if price_float < 0:
                return "0"
            
            return f"{price_float:.2f}"
            
        except ValueError:
            return "0"
    
    def export_to_excel_template(self, output_path: str = "product_template.xlsx"):
        """Create an Excel template with the required columns"""
        try:
            # Create a sample DataFrame
            sample_data = {
                'title': ['Product 1', 'Product 2'],
                'description': ['Description of product 1', 'Description of product 2'],
                'price': [29.99, 49.99],
                'sku': ['SKU001', 'SKU002'],
                'images_path': [
                    r'C:\Images\Product1\image1.jpg',  # Single image
                    r'C:\Images\Product2'              # Directory with images
                ]
            }
            
            df = pd.DataFrame(sample_data)
            
            # Save to Excel
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Products', index=False)
                
                # Add instructions sheet
                instructions = pd.DataFrame({
                    'Column': ['title', 'description', 'price', 'sku', 'images_path'],
                    'Required': ['Yes', 'Yes', 'Yes', 'No', 'Yes'],
                    'Description': [
                        'Product title',
                        'Product description (HTML supported)',
                        'Product price (numbers only, no currency symbols)',
                        'Stock Keeping Unit (optional)',
                        'Path to image file or directory containing images. Can be:\n- Single file: C:/images/product.jpg\n- Directory: C:/images/product/\n- Multiple files: C:/images/img1.jpg;C:/images/img2.jpg'
                    ],
                    'Example': [
                        'Wireless Headphones',
                        'Premium wireless headphones with noise cancellation',
                        '99.99',
                        'WH-2024-BLK',
                        'C:\\Products\\Headphones\\images\\'
                    ]
                })
                
                instructions.to_excel(writer, sheet_name='Instructions', index=False)
            
            logger.info(f"Template created: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error creating template: {e}")
            return False
    
    def validate_excel_file(self, excel_path: str) -> Dict[str, Any]:
        """Validate Excel file structure before processing"""
        validation_result = {
            'valid': False,
            'errors': [],
            'warnings': [],
            'file_exists': False,
            'has_required_columns': False,
            'sample_data': []
        }
        
        try:
            # Check if file exists
            if not os.path.exists(excel_path):
                validation_result['errors'].append(f"File does not exist: {excel_path}")
                return validation_result
            
            validation_result['file_exists'] = True
            
            # Try to read the file
            try:
                df = pd.read_excel(excel_path, nrows=5)  # Read first 5 rows for validation
            except Exception as e:
                validation_result['errors'].append(f"Cannot read Excel file: {str(e)}")
                return validation_result
            
            # Check columns
            df.columns = df.columns.str.strip().str.lower()
            required_columns = ['title', 'description', 'price', 'images_path']
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                validation_result['errors'].append(f"Missing columns: {', '.join(missing_columns)}")
                validation_result['has_required_columns'] = False
            else:
                validation_result['has_required_columns'] = True
            
            # Check for empty rows in required columns
            for index, row in df.iterrows():
                if pd.isna(row.get('title')) or pd.isna(row.get('price')):
                    validation_result['warnings'].append(f"Row {index + 2}: Missing title or price")
            
            # Get sample data
            if len(df) > 0:
                sample = df.head(3).to_dict('records')
                validation_result['sample_data'] = sample
            
            validation_result['valid'] = len(validation_result['errors']) == 0
            
            return validation_result
            
        except Exception as e:
            validation_result['errors'].append(f"Validation error: {str(e)}")
            return validation_result