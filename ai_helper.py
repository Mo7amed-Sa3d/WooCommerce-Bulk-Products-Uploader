import os
from dotenv import load_dotenv

load_dotenv()

class AIHelper:
    def __init__(self):
        self.api_key = os.getenv('OPENAI_API_KEY')
        self.available = False
        
        if self.api_key:
            try:
                import openai
                openai.api_key = self.api_key
                self.available = True
                self.client = openai
            except ImportError:
                pass
    
    def generate_title(self, prompt, num_titles=3):
        """Generate product titles using AI"""
        if not self.available:
            return []
        
        try:
            response = self.client.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a product title generator for e-commerce. Generate compelling product titles."},
                    {"role": "user", "content": f"Generate {num_titles} product titles for: {prompt}"}
                ],
                max_tokens=100,
                temperature=0.7
            )
            
            titles = response.choices[0].message.content.strip().split('\n')
            # Clean up the titles
            titles = [t.split('. ', 1)[1] if '. ' in t else t for t in titles if t.strip()]
            return titles[:num_titles]
        except Exception as e:
            return []
    
    def generate_description(self, product_title, product_type="product"):
        """Generate product description using AI"""
        if not self.available:
            return ""
        
        try:
            response = self.client.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a product description writer for e-commerce. Write SEO-friendly product descriptions."},
                    {"role": "user", "content": f"Write a detailed product description for this {product_type}: {product_title}\nInclude features, benefits, and specifications in a professional tone."}
                ],
                max_tokens=300,
                temperature=0.7
            )
            
            description = response.choices[0].message.content.strip()
            return description
        except Exception as e:
            return ""