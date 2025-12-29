import os
import requests
import mimetypes
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import logging

load_dotenv()

logger = logging.getLogger(__name__)

class WooCommerceAPI:
    def __init__(self):
        self.store_url = os.getenv('STORE_URL')
        self.consumer_key = os.getenv('WC_CONSUMER_KEY')
        self.consumer_secret = os.getenv('WC_CONSUMER_SECRET')
        self.api_base = f"{self.store_url}/wp-json/wc/v3"
        
    def test_connection(self):
        """Test connection to WooCommerce API"""
        try:
            response = requests.get(
                f"{self.api_base}/products",
                auth=HTTPBasicAuth(self.consumer_key, self.consumer_secret),
                params={'per_page': 1},
                timeout=10
            )
            return response.status_code == 200
        except Exception as e:
            logger.error(f"Connection test failed: {e}")
            return False
    
    def get_categories(self):
        """Retrieve all categories and subcategories"""
        try:
            categories = []
            page = 1
            per_page = 100
            
            while True:
                response = requests.get(
                    f"{self.api_base}/products/categories",
                    auth=HTTPBasicAuth(self.consumer_key, self.consumer_secret),
                    params={
                        'per_page': per_page,
                        'page': page,
                        'hide_empty': False
                    },
                    timeout=10
                )
                
                if response.status_code == 200:
                    batch = response.json()
                    categories.extend(batch)
                    
                    if len(batch) < per_page:
                        break
                    page += 1
                else:
                    break
            
            return categories
        except Exception as e:
            logger.error(f"Failed to get categories: {e}")
            return []
    
    def create_product(self, product_data):
        """Create a new product in WooCommerce"""
        try:
            response = requests.post(
                f"{self.api_base}/products",
                auth=HTTPBasicAuth(self.consumer_key, self.consumer_secret),
                json=product_data,
                headers={'Content-Type': 'application/json'},
                timeout=30
            )
            
            return {
                'success': response.status_code == 201,
                'status_code': response.status_code,
                'data': response.json() if response.status_code == 201 else response.text,
                'product_data': product_data
            }
        except Exception as e:
            logger.error(f"Failed to create product: {e}")
            return {'success': False, 'error': str(e)}


class WordPressMediaAPI:
    def __init__(self):
        self.store_url = os.getenv('STORE_URL')
        self.username = os.getenv('WP_USERNAME')
        self.password = os.getenv('WP_APP_PASSWORD')
    
    def upload_media(self, image_path):
        """Upload image to WordPress media library with full quality"""
        try:
            # Read image in binary mode (no compression, full quality)
            with open(image_path, 'rb') as img_file:
                img_data = img_file.read()
            
            filename = os.path.basename(image_path)
            mime_type, _ = mimetypes.guess_type(image_path)
            if not mime_type:
                mime_type = 'image/jpeg'
            
            media_url = f"{self.store_url}/wp-json/wp/v2/media"
            
            headers = {
                'Content-Disposition': f'attachment; filename="{filename}"',
                'Content-Type': mime_type
            }
            
            response = requests.post(
                media_url,
                headers=headers,
                data=img_data,  # Send original binary data
                auth=HTTPBasicAuth(self.username, self.password),
                timeout=30
            )
            
            if response.status_code in [200, 201]:
                media_data = response.json()
                return {
                    'success': True,
                    'id': media_data['id'],
                    'url': media_data['source_url']
                }
            else:
                logger.error(f"Media upload failed: {response.status_code} - {response.text}")
                return {
                    'success': False,
                    'error': f"HTTP {response.status_code}",
                    'details': response.text
                }
                
        except Exception as e:
            logger.error(f"Media upload exception: {e}")
            return {'success': False, 'error': str(e)}