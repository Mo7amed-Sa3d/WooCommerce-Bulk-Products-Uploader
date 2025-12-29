import os
import glob
import json
from pathlib import Path
from typing import List, Dict, Any
import logging

logger = logging.getLogger(__name__)

class BulkProductProcessor:
    def __init__(self):
        self.supported_image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.tiff']
    
    def scan_directory(self, directory_path: str) -> List[Dict[str, Any]]:
        """
        Scan a directory for product folders and extract product data
        
        Expected structure:
        ParentDirectory/
        ├── ProductFolder1/
        │   ├── title.txt
        │   ├── description.txt
        │   ├── price.txt
        │   ├── sku.txt
        │   └── images/
        │       ├── image1.jpg
        │       └── image2.jpg
        ├── ProductFolder2/
        └── ...
        """
        products = []
        scanned_count = 0
        error_count = 0
        
        try:
            directory = Path(directory_path)
            
            # Check if directory exists
            if not directory.exists():
                raise FileNotFoundError(f"Directory not found: {directory_path}")
            
            # Iterate through all subdirectories (product folders)
            for item in directory.iterdir():
                if item.is_dir():
                    product_data = self._process_product_folder(item)
                    if product_data:
                        products.append(product_data)
                        scanned_count += 1
                    else:
                        error_count += 1
                        logger.warning(f"Failed to process product folder: {item.name}")
            
            logger.info(f"Scanned {scanned_count} products, {error_count} errors")
            return products
            
        except Exception as e:
            logger.error(f"Error scanning directory: {e}")
            raise
    
    def _process_product_folder(self, folder_path: Path) -> Dict[str, Any]:
        """Process a single product folder"""
        try:
            folder_name = folder_path.name
            
            # Check required files
            required_files = ['title.txt', 'description.txt', 'price.txt']
            for file_name in required_files:
                file_path = folder_path / file_name
                if not file_path.exists():
                    logger.warning(f"Missing required file {file_name} in {folder_name}")
                    return None
            
            # Read text files
            title = self._read_text_file(folder_path / 'title.txt')
            description = self._read_text_file(folder_path / 'description.txt')
            price = self._read_text_file(folder_path / 'price.txt')
            sku = self._read_text_file(folder_path / 'sku.txt') if (folder_path / 'sku.txt').exists() else ""
            
            # Validate required fields
            if not title or not price:
                logger.warning(f"Missing title or price in {folder_name}")
                return None
            
            # Process images
            images_folder = folder_path / 'images'
            images = []
            if images_folder.exists() and images_folder.is_dir():
                images = self._get_images_from_folder(images_folder)
            else:
                # Try to find images directly in the product folder
                images = self._find_images_in_folder(folder_path)
            
            # Create product data dictionary
            product_data = {
                'folder_name': folder_name,
                'title': title.strip(),
                'description': description.strip(),
                'price': self._validate_price(price.strip()),
                'sku': sku.strip() if sku else "",
                'images': images,
                'has_images': len(images) > 0,
                'image_count': len(images),
                'status': 'pending',
                'folder_path': str(folder_path)
            }
            
            return product_data
            
        except Exception as e:
            logger.error(f"Error processing folder {folder_path}: {e}")
            return None
    
    def _read_text_file(self, file_path: Path) -> str:
        """Read content from a text file"""
        try:
            if file_path.exists():
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read()
            return ""
        except Exception as e:
            logger.error(f"Error reading file {file_path}: {e}")
            return ""
    
    def _get_images_from_folder(self, images_folder: Path) -> List[str]:
        """Get all images from an images folder"""
        images = []
        try:
            # Search for images recursively
            for ext in self.supported_image_extensions:
                pattern = f"*{ext}"
                for image_path in images_folder.rglob(pattern):
                    images.append(str(image_path))
            
            # Sort images for consistency (optional: sort by name)
            images.sort()
            return images
            
        except Exception as e:
            logger.error(f"Error getting images from {images_folder}: {e}")
            return []
    
    def _find_images_in_folder(self, folder_path: Path) -> List[str]:
        """Find images directly in the product folder"""
        images = []
        try:
            for ext in self.supported_image_extensions:
                pattern = f"*{ext}"
                for image_path in folder_path.glob(pattern):
                    images.append(str(image_path))
            
            # Sort by filename
            images.sort()
            return images
            
        except Exception as e:
            logger.error(f"Error finding images in {folder_path}: {e}")
            return []
    
    def _validate_price(self, price_str: str) -> str:
        """Validate and format price"""
        try:
            # Remove any currency symbols and whitespace
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
    
    def validate_products(self, products: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Validate a list of products and return statistics"""
        stats = {
            'total': len(products),
            'valid': 0,
            'invalid': 0,
            'with_images': 0,
            'without_images': 0,
            'errors': []
        }
        
        valid_products = []
        
        for product in products:
            # Check required fields
            if not product.get('title'):
                stats['errors'].append(f"{product.get('folder_name', 'Unknown')}: Missing title")
                stats['invalid'] += 1
                continue
            
            if not product.get('price') or product.get('price') == "0":
                stats['errors'].append(f"{product.get('folder_name', 'Unknown')}: Invalid price")
                stats['invalid'] += 1
                continue
            
            # Check if product has images (warning, not error)
            if product.get('has_images'):
                stats['with_images'] += 1
            else:
                stats['without_images'] += 1
                stats['errors'].append(f"{product.get('folder_name', 'Unknown')}: No images found")
            
            valid_products.append(product)
            stats['valid'] += 1
        
        return {
            'stats': stats,
            'valid_products': valid_products
        }
    
    def export_products_to_csv(self, products: List[Dict[str, Any]], output_file: str):
        """Export products to CSV for review"""
        import csv
        
        try:
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                # Write header
                writer.writerow(['Folder', 'Title', 'SKU', 'Price', 'Images', 'Status'])
                
                # Write data
                for product in products:
                    writer.writerow([
                        product.get('folder_name', ''),
                        product.get('title', ''),
                        product.get('sku', ''),
                        product.get('price', ''),
                        product.get('image_count', 0),
                        product.get('status', 'pending')
                    ])
            
            logger.info(f"Exported {len(products)} products to {output_file}")
            
        except Exception as e:
            logger.error(f"Error exporting to CSV: {e}")
    
    def create_batch_summary(self, products: List[Dict[str, Any]]) -> str:
        """Create a summary of the batch"""
        if not products:
            return "No products found"
        
        total_images = sum(p.get('image_count', 0) for p in products)
        with_images = sum(1 for p in products if p.get('has_images'))
        without_images = sum(1 for p in products if not p.get('has_images'))
        
        summary = f"""
        📦 BATCH SUMMARY
        {'='*40}
        Total Products: {len(products)}
        Total Images: {total_images}
        Products with Images: {with_images}
        Products without Images: {without_images}
        
        First 5 Products:
        """
        
        for i, product in enumerate(products[:5]):
            summary += f"\n{i+1}. {product.get('title', 'Unknown')} ({product.get('image_count', 0)} images)"
        
        if len(products) > 5:
            summary += f"\n... and {len(products) - 5} more"
        
        return summary