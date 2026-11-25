import io
import json
import os
import time
import streamlit as st
from docx import Document
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pypdf import PdfReader

st.set_page_config(
    page_title="PDF / Word မှ PowerPoint ထုတ်ပေးသည့် စနစ်",
    page_icon="📊",
    layout="wide",
)

st.title("📊 PDF / Word မှ PowerPoint ထုတ်ပေးသည့် စနစ်")
st.caption("PDF နှင့် Word (.docx) ဖိုင်များကို တင်ပြီး Theme, Icon Tag များဖြင့် လှပသော PowerPoint (.pptx) ဖိုင် ထုတ်ယူပါ။")

EXCLUDED_KEYWORDS = ("image", "audio", "embedding", "veo", "imagen", "omni", "tts", "live")

THEMES = {
    "Cyber Security (Dark)": {
        "bg": RGBColor(15, 23, 42),
        "card_bg": RGBColor(30, 41, 59),
        "header_bg": RGBColor(2, 6, 23),
        "title": RGBColor(248, 250, 252),
        "body": RGBColor(226, 232, 240),
        "muted": RGBColor(148, 163, 184),
        "accent": RGBColor(56, 189, 248),
        "badge_bg": RGBColor(239, 68, 68),
        "badge_fg": RGBColor(255, 255, 255),
    },
    "Corporate Navy (Light)": {
        "bg": RGBColor(248, 250, 252),
        "card_bg": RGBColor(255, 255, 255),
        "header_bg": RGBColor(15, 23, 42),
        "title": RGBColor(15, 23, 42),
        "body": RGBColor(30, 41, 59),
        "muted": RGBColor(100, 116, 139),
        "accent": RGBColor(2, 132, 199),
        "badge_bg": RGBColor(217, 119, 6),
        "badge_fg": RGBColor(255, 255, 255),
    },
    "Modern Teal (Light)": {
        "bg": RGBColor(240, 253, 250),
        "card_bg": RGBColor(255, 255, 255),
        "header_bg": RGBColor(19, 78, 74),
        "title": RGBColor(19, 78, 74),
        "body": RGBColor(31, 41, 55),
        "muted": RGBColor(75, 85, 99),
        "accent": RGBColor(13, 148, 136),
        "badge_bg": RGBColor(15, 118, 110),
        "badge_fg": RGBColor(255, 255, 255),
    },
}


def get_api_key() -> str:
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    if "GEMINI_API_KEY" in os.environ:
        return os.environ["GEMINI_API_KEY"]
    return st.sidebar.text_input("Gemini API Key ထည့်ပါ", type="password")


def is_suitable_text_model(model_id: str) -> bool:
    low = model_id.lower()
    if "gemini" not in low:
        return False
    if any(k in low for k in EXCLUDED_KEYWORDS):
        return False
    return True


@st.cache_data(ttl=600, show_spinner=False)
def list_text_gemini_models(api_key: str) -> list[str]:
    client = genai.Client(api_key=api_key)
    models = []
    for m in client.models.list():
        name = getattr(m, "name", "") or ""
        short_name = name.replace("models/", "")
        actions = getattr(m, "supported_actions", None) or []
        if ("generateContent" in actions or not actions) and is_suitable_text_model(short_name):
            if short_name not in models:
                models.append(short_name)

    models.sort(key=lambda x: (0 if "flash" in x else 1, x))
    return models or ["gemini-3.8-flash", "gemini-2.5-flash"]


def detect_slide_icon_and_tag(title: str, bullets: list[str], idx: int, total: int) -> tuple[str, str]:
    text = f"{title} {' '.join(bullets)}".lower()
    if idx == 0:
        return "🛡️", "OVERVIEW"
    if idx == total - 1 and any(w in text for w in ("summary", "conclusion", "နိဂုံး", "အနှစ်ချုပ်", "recommendation", "အကြံပြု")):
        return "✅", "SUMMARY"
