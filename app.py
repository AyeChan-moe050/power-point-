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
        "bg": RGBColor(15, 23, 42),          # Slate 900
        "card_bg": RGBColor(30, 41, 59),     # Slate 800
        "header_bg": RGBColor(2, 6, 23),     # Slate 950
        "title": RGBColor(248, 250, 252),    # White-ish
        "body": RGBColor(226, 232, 240),     # Light slate
        "muted": RGBColor(148, 163, 184),    # Muted slate
        "accent": RGBColor(56, 189, 248),    # Sky 400
        "badge_bg": RGBColor(239, 68, 68),   # Red 500
        "badge_fg": RGBColor(255, 255, 255),
    },
    "Corporate Navy (Light)": {
        "bg": RGBColor(248, 250, 252),       # Light background
        "card_bg": RGBColor(255, 255, 255),  # White card
        "header_bg": RGBColor(15, 23, 42),   # Dark navy banner
        "title": RGBColor(15, 23, 42),
        "body": RGBColor(30, 41, 59),
        "muted": RGBColor(100, 116, 139),
        "accent": RGBColor(2, 132, 199),     # Sky 600
        "badge_bg": RGBColor(217, 119, 6),   # Amber 600
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


def generate_slides_json(
    api_key: str,
    source_text: str,
    slide_count: int,
    language_choice: str,
    extra_note: str,
    model_name: str,
    max_chars: int = 15000,
) -> list[dict]:
    client = genai.Client(api_key=api_key)
    trimmed_text = source_text[:max_chars]

    if language_choice == "မြန်မာဘာသာ (Myanmar)":
        lang_instruction = (
            "Slide များ၏ title နှင့် bullets အားလုံးကို **မြန်မာဘာသာ** သီးသန့်ဖြင့် ရှင်းလင်းတိုတုတ်စွာ ရေးပါ "
            "(Slide တစ်ခုလျှင် bullet ၃ မှ ၄ ချက်ခန့်၊ ကျစ်လျစ်စွာ ရေးပါ)။"
        )
    else:
        lang_instruction = (
            "Write all slide titles and bullet points strictly in **English**, concise and clear "
            "(around 3 to 4 bullet points per slide)."
        )

    prompt = f"""
သင်သည် အစိုးရဌာန/အဖွဲ့အစည်းများနှင့် ပရော်ဖက်ရှင်နယ် တင်ပြချက်များအတွက် PowerPoint presentation ဖွဲ့စည်းပုံကို စနစ်တကျ ရေးသားပေးနိုင်သော ကျွမ်းကျင်သူတစ်ဦး ဖြစ်သည်။

ပေးထားသော PDF / Word စာရွက်စာတမ်းများမှ အဓိက အချက်အလက်များကို အခြေခံ၍ Slide အရေအတွက် ခန့်မှန်းခြေ {slide_count} slides ခန့် ပါဝင်သော PowerPoint outline ကို ပြုစုပေးပါ။

### လိုက်နာရမည့် စည်းကမ်းများ
1. {lang_instruction}
2. ပထမဆုံး Slide သည် Presentation ခေါင်းစဉ် (Title Slide) ဖြစ်သင့်ပြီး နောက်ဆုံး Slide သည် နိဂုံး/အနှစ်ချုပ် (Summary / Conclusion) ဖြစ်သင့်သည်။
3. Source text ထဲတွင် မပါဝင်သော နေ့ရက်၊ ကိန်းဂဏန်း၊ အမည်များကို မတီထွင်ပါနှင့်။

### ထပ်မံမှာကြားချက် (ရှိပါက)
{extra_note.strip() if extra_note.strip() else "မရှိပါ"}

### စာရွက်စာတမ်းများမှ ထုတ်ယူထားသော အချက်အလက်များ
{trimmed_text}
""".strip()

    response_schema = {
        "type": "OBJECT",
        "properties": {
            "slides": {
                "type": "ARRAY",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "title": {"type": "STRING"},
                        "bullets": {"type": "ARRAY", "items": {"type": "STRING"}},
                    },
                    "required": ["title", "bullets"],
                },
            }
        },
        "required": ["slides"],
    }

    last_err = None
    for attempt in range(2):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=response_schema,
                ),
            )
            raw_text = response.text or "{}"
            result = json.loads(raw_text)
            return result.get("slides", [])
        except Exception as exc:
            last_err = exc
            err_str = str(exc)
            if ("503" in err_str or "UNAVAILABLE" in err_str) and attempt == 0:
                time.sleep(2)
                continue
            raise

    raise RuntimeError(f"Error: {last_err}")


def _set_run_font(run, font_name: str, size_pt: int, color: RGBColor, bold: bool = False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = color
    run.font.bold = bold


def build_styled_pptx_bytes(
    slides: list[dict],
    font_name: str,
    theme_name: str,
    template_file=None,
) -> bytes:
    theme = THEMES.get(theme_name, THEMES["Cyber Security (Dark)"])
    blank_layout_idx = 6

    if template_file is not None:
        prs = Presentation(io.BytesIO(template_file.getvalue()))
        if len(prs.slide_layouts) <= blank_layout_idx:
            blank_layout_idx = len(prs.slide_layouts) - 1
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    total_slides = len(slides)

    for idx, slide_data in enumerate(slides):
        title_text = str(slide_data.get("title", "")).strip()
        bullets = [str(b).strip() for b in slide_data.get("bullets", []) if str(b).strip()]
        icon_emoji, tag_label = detect_slide_icon_and_tag(title_text, bullets, idx, total_slides)

        layout = prs.slide_layouts[min(blank_layout_idx, len(prs.slide_layouts) - 1)]
        slide = prs.slides.add_slide(layout)

        # Background fill if not using custom template
        if template_file is None:
            bg_fill = slide.background.fill
            bg_fill.solid()
            bg_fill.fore_color.rgb = theme["bg"]

        # Top accent bar
        top_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = theme["accent"]
        top_bar.line.fill.background()

        # Icon badge pill on top-left
        badge = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(0.38), Inches(2.55), Inches(0.45)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = theme["badge_bg"]
        badge.line.fill.background()
        b_tf = badge.text_frame
        b_tf.word_wrap = True
        b_tf.margin_left = Inches(0.12)
        b_tf.margin_right = Inches(0.12)
        b_tf.margin_top = Inches(0.04)
        b_tf.margin_bottom = Inches(0.04)
        bp = b_tf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        brun = bp.add_run()
        brun.text = f"{icon_emoji}  {tag_label}"
        _set_run_font(brun, font_name, 12, theme["badge_fg"], bold=True)

        # Slide Title
        title_box = slide.shapes.add_textbox(Inches(0.65), Inches(0.95), Inches(12.0), Inches(0.95))
        t_tf = title_box.text_frame
        t_tf.word_wrap = True
        tp = t_tf.paragraphs[0]
        trun = tp.add_run()
        trun.text = f"{icon_emoji} {title_text}"
        _set_run_font(trun, font_name, 26 if idx > 0 else 32, theme["title"], bold=True)

        # Divider line under title
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(1.95), Inches(12.03), Inches(0.03)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = theme["accent"]
        line.line.fill.background()

        # Content Card
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(2.15), Inches(12.03), Inches(4.65)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = theme["card_bg"]
        card.line.color.rgb = theme["accent"]
        card.line.width = Pt(1)

        c_tf = card.text_frame
        c_tf.word_wrap = True
        c_tf.margin_left = Inches(0.35)
        c_tf.margin_right = Inches(0.35)
        c_tf.margin_top = Inches(0.28)
        c_tf.margin_bottom = Inches(0.25)

        if bullets:
            for b_idx, b_text in enumerate(bullets):
                p = c_tf.paragraphs[0] if b_idx == 0 else c_tf.add_paragraph()
                p.space_after = Pt(12)
                p.line_spacing = 1.25

                dot_run = p.add_run()
                dot_run.text = f"▸  "
                _set_run_font(dot_run, font_name, 18, theme["accent"], bold=True)

                txt_run = p.add_run()
                txt_run.text = b_text
                _set_run_font(txt_run, font_name, 18, theme["body"], bold=False)

        # Footer with slide number
        footer_box = slide.shapes.add_textbox(Inches(0.65), Inches(6.95), Inches(12.03), Inches(0.35))
        f_tf = footer_box.text_frame
        fp = f_tf.paragraphs[0]
        fp.alignment = PP_ALIGN.RIGHT
        frun = fp.add_run()
        frun.text = f"Slide {idx + 1} / {total_slides}"
        _set_run_font(frun, font_name, 11, theme["muted"], bold=False)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


api_key = get_api_key()

model_options = ["gemini-3.8-flash", "gemini-2.5-flash"]
if api_key:
    try:
        model_options = list_text_gemini_models(api_key)
    except Exception:
        pass

col_u1, col_u2 = st.columns([2, 1])
with col_u1:
    uploaded_files = st.file_uploader(
        "PDF (.pdf) သို့မဟုတ် Word (.docx) ဖိုင်များ တင်ပါ",
        type=["pdf
