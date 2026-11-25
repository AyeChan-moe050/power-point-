import io
import json
import os
import streamlit as st
from docx import Document
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pypdf import PdfReader

st.set_page_config(
    page_title="PDF / Word မှ PowerPoint (မြန်မာ + English) ထုတ်ပေးသည့် စနစ်",
    page_icon="📊",
    layout="wide",
)

st.title("📊 PDF / Word မှ PowerPoint (မြန်မာ + English) ထုတ်ပေးသည့် စနစ်")
st.caption("PDF နှင့် Word (.docx) ဖိုင်များကို တင်ပြီး မြန်မာဘာသာ PowerPoint ၁ ဖိုင် နှင့် English PowerPoint ၁ ဖိုင် ထုတ်ယူပါ။")


def get_api_key() -> str:
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    if "GEMINI_API_KEY" in os.environ:
        return os.environ["GEMINI_API_KEY"]
    return st.sidebar.text_input("Gemini API Key ထည့်ပါ", type="password")


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
                parts.append(row_text)
    return "\n".join(parts)


def extract_pdf_text(uploaded_file) -> str:
    reader = PdfReader(io.BytesIO(uploaded_file.getvalue()))
    parts = [f"=== PDF ဖိုင်: {uploaded_file.name} ==="]
    for page_num, page in enumerate(reader.pages, start=1):
        txt = (page.extract_text() or "").strip()
        if txt:
            parts.append(f"--- စာမျက်နှာ {page_num} ---\n{txt}")
    return "\n\n".join(parts)


def generate_bilingual_slides_json(
    api_key: str,
    source_text: str,
    slide_count: int,
    extra_note: str,
) -> dict:
    client = genai.Client(api_key=api_key)

    prompt = f"""
သင်သည် အစိုးရဌာန/အဖွဲ့အစည်းများနှင့် ပရော်ဖက်ရှင်နယ် တင်ပြချက်များအတွက် PowerPoint presentation ဖွဲ့စည်းပုံကို စနစ်တကျ ရေးသားပေးနိုင်သော ကျွမ်းကျင်သူတစ်ဦး ဖြစ်သည်။

ပေးထားသော PDF / Word စာရွက်စာတမ်းများမှ အဓိက အချက်အလက်များကို အခြေခံ၍ Slide အရေအတွက် ခန့်မှန်းခြေ {slide_count} slides ခန့် ပါဝင်သော PowerPoint outline ကို **မြန်မာဘာသာ (myanmar)** နှင့် **အင်္ဂလိပ်ဘာသာ (english)** နှစ်မျိုးလုံးအတွက် ပြုစုပေးပါ။

### လိုက်နာရမည့် စည်းကမ်းများ
1. `myanmar` array ထဲရှိ slide များ၏ title နှင့် bullets များကို **မြန်မာဘာသာ** သီးသန့်ဖြင့် ရှင်းလင်းတိုတုတ်စွာ ရေးပါ (Slide တစ်ခုလျှင် bullet ၃ မှ ၅ ချက်ခန့်၊ စာကြောင်းတစ်ကြောင်းချင်းစီကို Slide မှာ အလွယ်တကူ ဖတ်နိုင်အောင် ကျစ်လျစ်စွာ ရေးပါ)။
2. `english` array ထဲရှိ slide များ၏ title နှင့် bullets များကို **အင်္ဂလိပ်ဘာသာ** သီးသန့်ဖြင့် အဓိပ္ပာယ်ကိုက်ညီစွာ ရေးပါ (Slide တစ်ခုလျှင် bullet ၃ မှ ၅ ချက်ခန့်)။
3. ပထမဆုံး Slide သည် Presentation ခေါင်းစဉ် (Title Slide) ဖြစ်သင့်ပြီး နောက်ဆုံး Slide သည် နိဂုံး/အနှစ်ချုပ် (Summary / Conclusion) ဖြစ်သင့်သည်။
4. Source text ထဲတွင် မပါဝင်သော နေ့ရက်၊ ကိန်းဂဏန်း၊ အမည်များကို မတီထွင်ပါနှင့်။

### ထပ်မံမှာကြားချက် (ရှိပါက)
{extra_note.strip() if extra_note.strip() else "မရှိပါ"}

### စာရွက်စာတမ်းများမှ ထုတ်ယူထားသော အချက်အလက်များ
{source_text}
""".strip()

    response_schema = {
        "type": "OBJECT",
        "properties": {
            "myanmar": {
                "type": "ARRAY",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "title": {"type": "STRING"},
                        "bullets": {"type": "ARRAY", "items": {"type": "STRING"}},
                    },
                    "required": ["title", "bullets"],
                },
            },
            "english": {
                "type": "ARRAY",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "title": {"type": "STRING"},
                        "bullets": {"type": "ARRAY", "items": {"type": "STRING"}},
                    },
                    "required": ["title", "bullets"],
                },
            },
        },
        "required": ["myanmar", "english"],
    }

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response
