## Bulk Products Uploader using WooCommerce REST API

### Description
- This app is used to upload/bulk upload products to WooCommerce based wordpress apps

### Setup and Run
- First you should create a file called <mark>.env</mark> , then put your WooCommerce API Key & Wordpress app password in the .env file

```python
# Store URL (no trailing slash)
STORE_URL=https://yoursite.com

# WordPress User & App Password
WP_USERNAME=username
WP_APP_PASSWORD=password  # NO SPACES!


WC_CONSUMER_KEY=ck_********************************
WC_CONSUMER_SECRET=cs_********************************
# You can keep WP_USERNAME/PASSWORD if you still need image uploads
# OpenAI API Key
OPENAI_API_KEY=sk-your-key-here

# Default category ID (optional)
DEFAULT_CATEGORY=15
```
- To run the app, make sure to be in the directory of the app then use the command in the CMD:
```python
python main.py
```