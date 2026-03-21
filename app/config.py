import os
from dotenv import load_dotenv

load_dotenv()

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")
AI_MODEL = os.getenv("AI_MODEL", "openai/gpt-4o")
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "")

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates")

os.makedirs(OUTPUT_DIR, exist_ok=True)
