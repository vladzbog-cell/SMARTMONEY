import os

from dotenv import load_dotenv
from google import genai
from google.genai import types


load_dotenv("keys.env")

ARTEMOX_API_KEY = os.getenv("ARTEMOX_API_KEY", "").strip()
ARTEMOX_BASE_URL = os.getenv("ARTEMOX_BASE_URL", "https://api.artemox.com").strip()
ARTEMOX_MODEL = os.getenv("ARTEMOX_MODEL", "gemini-3.1-pro-preview").strip()
ARTEMOX_MEDIA_MODEL = os.getenv("ARTEMOX_MEDIA_MODEL", "gemini-3.1-flash-image-preview").strip()
ARTEMOX_IMAGE_GENERATION_MODEL = os.getenv("ARTEMOX_IMAGE_GENERATION_MODEL", "gemini-2.5-flash-image").strip()
ARTEMOX_QUOTA_EXCEEDED_SENTINEL = "__ARTEMOX_QUOTA_EXCEEDED__"
ARTEMOX_MEDIA_QUOTA_EXCEEDED_SENTINEL = "__ARTEMOX_MEDIA_QUOTA_EXCEEDED__"


def get_artemox_client() -> genai.Client | None:
    if not ARTEMOX_API_KEY:
        return None
    return genai.Client(
        api_key=ARTEMOX_API_KEY,
        http_options=types.HttpOptions(base_url=ARTEMOX_BASE_URL),
    )
