from dotenv import load_dotenv
import os

load_dotenv()

PROXY_HOST = os.getenv('PROXY_HOST', '')
PROXY_PORT = os.getenv('PROXY_PORT', '')
PROXY_USER = os.getenv('PROXY_USER', '')
PROXY_PASS = os.getenv('PROXY_PASS', '')
MIN_DELAY = int(os.getenv('MIN_DELAY', 8))
MAX_DELAY = int(os.getenv('MAX_DELAY', 16))
