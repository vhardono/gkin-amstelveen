import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # Dropbox configuration
    DROPBOX_APP_KEY = os.getenv('DROPBOX_APP_KEY')
    DROPBOX_APP_SECRET = os.getenv('DROPBOX_APP_SECRET')
    DROPBOX_REFRESH_TOKEN = os.getenv('DROPBOX_REFRESH_TOKEN')
    DROPBOX_EXCEL_PATH = os.getenv('DROPBOX_EXCEL_PATH', '/Church/Excel')
    
    # Website configurations
    PREEKROSTER_URL = os.getenv('PREEKROSTER_URL', 'https://example.com/preekroster')
    SCIPIO_BASE_URL = os.getenv('SCIPIO_BASE_URL', 'https://scipio.example.com')
    SCIPIO_USERNAME = os.getenv('SCIPIO_USERNAME')
    SCIPIO_PASSWORD = os.getenv('SCIPIO_PASSWORD')
    SCIPIO_PIN = os.getenv('SCIPIO_PIN')
    
    # Output configuration
    OUTPUT_DIR = os.getenv('OUTPUT_DIR', './output')
    TEMPLATE_DIR = os.getenv('TEMPLATE_DIR', './doc_templates')
    
    # Bulletin configuration
    CHURCH_NAME = os.getenv('CHURCH_NAME', 'Gemeente')
    BULLETIN_LANGUAGE = os.getenv('BULLETIN_LANGUAGE', 'nl')
