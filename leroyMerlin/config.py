from dotenv import load_dotenv
import os

load_dotenv()

PROXY_HOST = os.getenv('PROXY_HOST', '127.0.0.0')
PROXY_PORT = os.getenv('PROXY_PORT', '8080')
PROXY_USER = os.getenv('PROXY_USER', 'user')
PROXY_PASS = os.getenv('PROXY_PASS', 'pass')
MIN_DELAY = int(os.getenv('MIN_DELAY', 8))
MAX_DELAY = int(os.getenv('MAX_DELAY', 16))
