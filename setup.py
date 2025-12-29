#!/usr/bin/env python3
import os
import sys
import subprocess

def check_and_install_dependencies():
    """Check and install required packages"""
    required = [
        'requests>=2.28.0',
        'python-dotenv>=0.21.0',
        'Pillow>=9.0.0',
    ]
    
    optional = [
        'openai>=0.27.0'
    ]
    
    print("Checking dependencies...")
    
    for package in required:
        try:
            __import__(package.split('>=')[0].replace('-', '_'))
            print(f"✓ {package}")
        except ImportError:
            print(f"✗ {package} - Installing...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    
    print("\nOptional dependencies for AI features:")
    for package in optional:
        try:
            __import__(package.split('>=')[0].replace('-', '_'))
            print(f"✓ {package} (AI features enabled)")
        except ImportError:
            print(f"✗ {package} (AI features disabled)")
    
    # Create .env.example if it doesn't exist
    if not os.path.exists('.env'):
        if os.path.exists('.env.example'):
            print("\n⚠️  Warning: .env file not found.")
            print("Copy .env.example to .env and update your credentials.")
        else:
            # Create .env.example
            with open('.env.example', 'w') as f:
                f.write("""# WooCommerce Store URL
STORE_URL=https://yourstore.com

# WordPress Credentials (for media upload)
WP_USERNAME=your_email@example.com
WP_APP_PASSWORD=xxxx xxxx xxxx xxxx xxxx xxxx

# WooCommerce API Credentials
WC_CONSUMER_KEY=ck_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
WC_CONSUMER_SECRET=cs_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

# OpenAI API Key (Optional - for AI features)
# OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
""")
            print("\n📄 Created .env.example file. Copy it to .env and update your credentials.")

if __name__ == "__main__":
    check_and_install_dependencies()
    print("\n✅ Setup complete!")
    print("\nTo run the application:")
    print("1. Copy .env.example to .env")
    print("2. Update .env with your credentials")
    print("3. Run: python main.py")