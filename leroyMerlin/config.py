from dotenv import load_dotenv
import os

load_dotenv()

PROXY_HOST = os.getenv('PROXY_HOST', None)
PROXY_PORT = os.getenv('PROXY_PORT', None)
PROXY_USER = os.getenv('PROXY_USER', None)
PROXY_PASS = os.getenv('PROXY_PASS', None)
MIN_DELAY = int(os.getenv('MIN_DELAY', None))
MAX_DELAY = int(os.getenv('MAX_DELAY', None))
