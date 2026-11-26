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
st.caption("PDF / Word ဖိုင်တင်၍ ဖြစ်စေ၊ ခေါင်းစဉ်ရိုက်ထည့်၍ ဖြစ်စေ Template ပုံစံ ရွေးချယ်ကာ PowerPoint (.pptx) ဖိုင် ထုတ်ယူပါ။")

DEFAULT_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.8-pro",
    "gemini-3-flash-preview",
    "gemini-3-pro-preview",
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
]

TEMPLATES = {
    "Cyber Dark Cards": {
        "style": "cards",
        "bg": RGBColor(15, 23, 42),
        "card_bg": RGBColor(30, 41, 59),
        "side_bg": RGBColor(2, 6, 23),
        "title": RGBColor(248, 250, 252),
        "body": RGBColor(226, 232, 240),
        "muted": RGBColor(148, 163, 184),
        "accent": RGBColor(56, 189, 248),
        "badge_bg": RGBColor(239, 68, 68),
        "badge_fg": RGBColor(255, 255, 255),
    },
    "Clean Left-Accent (Light)": {
        "style": "left_accent",
        "bg": RGBColor(248, 250, 252),
        "card_bg": RGBColor(255, 255, 255),
        "side_bg": RGBColor(2, 132, 199),
        "title": RGBColor(15, 23, 42),
        "body": RGBColor(30, 41, 59),
        "muted": RGBColor(100, 116, 139),
        "accent": RGBColor(2, 132, 199),
        "badge_bg": RGBColor(14, 116, 144),
        "badge_fg": RGBColor(255, 255, 255),
    },
    "Executive Split": {
        "style": "split",
        "bg": RGBColor(241, 245, 249),
        "card_bg": RGBColor(255, 255, 255),
        "side_bg": RGBColor(15, 23, 42),
        "title": RGBColor(15, 23, 42),
        "body": RGBColor(30, 41, 59),
        "muted": RGBColor(100, 116, 139),
        "accent": RGBColor(217, 119, 6),
        "badge_bg": RGBColor(217, 119, 6),
        "badge_fg": RGBColor(255, 255, 255),
    },
    "Minimalist Modern": {
        "style": "minimal",
        "bg": RGBColor(255, 255, 255),
        "card_bg": RGBColor(248, 250, 252),
        "side_bg": RGBColor(19, 78, 74),
        "title": RGBColor(19, 78, 74),
        "body": RGBColor(31, 41, 55),
        "muted": RGBColor(107, 114, 128),
        "accent": RGBColor(13, 148, 136),
        "badge_bg": RGBColor(13, 148, 136),
        "badge_fg": RGBColor(255, 255, 255),
    },
}


def get_api_key() -> str:
    try:
        if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass
    if os.environ.get("GEMINI_API_KEY"):
        return os.environ["GEMINI_API_KEY"].strip()
    return st.sidebar.text_input("Gemini API Key ထည့်ပါ", type="password").strip()


def fetch_all_gemini_models(api_key: str) -> list[str]:
    client = genai.Client(api_key=api_key)
    models = []
    for m in client.models.list():
        name = getattr(m, "name", "") or ""
        short_name = name.replace("models/", "")
        if "gemini" in short_name.lower() and short_name not in models:
            models.append(short_name)
    for dm in DEFAULT_MODELS:
        if dm not in models:
            models.append(dm)
    return models


def detect_slide_icon_and_tag(title: str, bullets: list[str], idx: int, total: int) -> tuple[str, str]:
    text = f"{title} {' '.join(bullets)}".lower()
    if idx == 0:
        return "🛡️", "OVERVIEW"
    if idx == total - 1 and any(w in text for w in ("summary", "conclusion", "နိဂုံး", "အနှစ်ချုပ်", "recommendation", "အကြံပြု")):
        return "✅", "SUMMARY"
    if any(w in text for w in ("vpn", "phishing", "link", "email", "အတုအယောင်", "လင့်ခ်", "အီးမေးလ်")):
        return "🎣", "PHISHING / VPN"
    if any(w in text for w in ("malware", "installer", "infect", "ကူးစက်", "ထည့်သွင်း", "trojan", "ransomware")):
        return "🦠", "MALWARE"
    if any(w in text for w in ("persistence", "registry", "scheduled task", "c2", "command and control", "server", "ဆက်သွယ်")):
        return "⚙️", "PERSISTENCE / C2"
    if any(w in text for w in ("ioc", "hash", "indicator", "အမှတ်အသား", "သံသယ")):
        return "🔍", "IoCs"
    if any(w in text for w in ("credential", "password", "access", "account", "အကောင့်", "ဝင်ရောက်")):
        return "🔐", "ACCESS / CREDENTIALS"
    return "📌", f"SLIDE {idx + 1:02d}"


def extract_docx_text(uploaded_file) -> str:
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    parts = [f"=== Word ဖိုင်: {uploaded_file.name} ==="]
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())
    for idx, table in enumerate(doc.tables, start=1):
        parts.append(f"[ဇယား {idx}]")
        for row in table.rows:
            row_text = " | ".join(cell.text.strip().replace("\n", " ") for cell in row.cells)
            if row_text.strip():
                parts.append
