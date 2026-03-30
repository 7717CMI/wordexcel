import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class Config:
    # OpenAI Configuration
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY', '')
    
    # File Upload Configuration
    MAX_FILE_SIZE = int(os.getenv('MAX_FILE_SIZE', 52428800))  # 50MB default
    UPLOAD_DIR = os.getenv('UPLOAD_DIR', './uploads')
    
    # Excel Template Configuration
    EXCEL_TEMPLATE_PATH = os.getenv('EXCEL_TEMPLATE_PATH', './assets/Bonding_Neodymium_Magnet_Market.xlsm')
    
    # Server Configuration
    HOST = os.getenv('HOST', '0.0.0.0')
    PORT = int(os.getenv('PORT', 8000))
    
    # Ensure upload directory exists
    @classmethod
    def ensure_directories(cls):
        """Ensure required directories exist"""
        os.makedirs(cls.UPLOAD_DIR, exist_ok=True)
        os.makedirs('./temp', exist_ok=True)
