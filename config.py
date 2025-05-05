import os
from dotenv import load_dotenv

load_dotenv()

POWERPOINT_PATH = os.getenv("POWERPOINT_PATH")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
DEBUG = os.getenv("DEBUG", "False") == "True"
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", 5000))