import os
import logging
from datetime import datetime
from typing import List, Dict, Any


def setup_logging(log_file='uploader.log'):
    """Setup logging configuration"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

def validate_image_paths(image_paths):
    """Validate that image files exist and are accessible"""
    valid_paths = []
    for path in image_paths:
        if os.path.exists(path) and os.path.isfile(path):
            valid_paths.append(path)
        else:
            logging.warning(f"Image not found or inaccessible: {path}")
    return valid_paths

def format_price(price_str):
    """Format price string to WooCommerce format"""
    try:
        price = float(price_str)
        return str(round(price, 2))
    except (ValueError, TypeError):
        return "0"

def build_category_tree(categories):
    """Build hierarchical category tree from flat list"""
    category_dict = {}
    category_list = []
    
    # First pass: store all categories by ID
    for cat in categories:
        category_dict[cat['id']] = cat
    
    # Build tree
    def build_tree(parent_id=0, level=0):
        tree = []
        for cat in categories:
            if cat['parent'] == parent_id:
                indent = "  " * level
                display_name = f"{indent}{cat['name']} (ID: {cat['id']})"
                tree.append({
                    'display': display_name,
                    'id': cat['id'],
                    'children': build_tree(cat['id'], level + 1)
                })
        return tree
    
    # Flatten tree for display
    def flatten_tree(tree, result_list):
        for item in tree:
            result_list.append(item['display'])
            flatten_tree(item['children'], result_list)
    
    tree = build_tree()
    flatten_tree(tree, category_list)
    
    return category_list, tree


def validate_bulk_directory(directory_path: str) -> Dict[str, Any]:
    """Validate a bulk upload directory structure"""
    import os
    from pathlib import Path
    
    validation_result = {
        'valid': False,
        'errors': [],
        'warnings': [],
        'product_count': 0,
        'structure_valid': False
    }
    
    try:
        path = Path(directory_path)
        
        # Check if directory exists
        if not path.exists():
            validation_result['errors'].append(f"Directory does not exist: {directory_path}")
            return validation_result
        
        if not path.is_dir():
            validation_result['errors'].append(f"Path is not a directory: {directory_path}")
            return validation_result
        
        # Check directory structure
        product_folders = []
        for item in path.iterdir():
            if item.is_dir():
                # Check if it's a product folder (has required files)
                required_files_exist = all(
                    (item / file).exists()
                    for file in ['title.txt', 'description.txt', 'price.txt']
                )
                
                if required_files_exist:
                    product_folders.append(item.name)
                else:
                    validation_result['warnings'].append(
                        f"Folder '{item.name}' missing required files"
                    )
        
        if not product_folders:
            validation_result['errors'].append("No valid product folders found")
            return validation_result
        
        validation_result['product_count'] = len(product_folders)
        validation_result['valid'] = True
        validation_result['structure_valid'] = True
        validation_result['product_folders'] = product_folders
        
        return validation_result
        
    except Exception as e:
        validation_result['errors'].append(f"Validation error: {str(e)}")
        return validation_result

def create_batch_log(batch_id: str, products: List[Dict[str, Any]], category_id: int, 
                    output_dir: str = "batch_logs") -> str:
    """Create a log file for the batch upload"""
    import json
    from datetime import datetime
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(output_dir, f"batch_{batch_id}_{timestamp}.json")
    
    log_data = {
        'batch_id': batch_id,
        'timestamp': timestamp,
        'category_id': category_id,
        'total_products': len(products),
        'products': products
    }
    
    try:
        with open(log_file, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, indent=2, ensure_ascii=False)
        
        return log_file
    except Exception as e:
        logging.error(f"Error creating batch log: {e}")
        return ""

def format_bulk_stats(stats: Dict[str, Any]) -> str:
    """Format bulk upload statistics for display"""
    return f"""
📊 BATCH STATISTICS
────────────────────────────────────
Total Products: {stats.get('total', 0)}
✓ Valid: {stats.get('valid', 0)}
✗ Invalid: {stats.get('invalid', 0)}
🖼️ With Images: {stats.get('with_images', 0)}
⚠️ Without Images: {stats.get('without_images', 0)}

Total Images: {sum(p.get('image_count', 0) for p in stats.get('products', []))}
────────────────────────────────────
"""