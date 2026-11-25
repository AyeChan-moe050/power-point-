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
st.caption("PDF နှင့် Word (.docx) ဖိုင်များကို တင်ပြီး Template ပုံစံ ရွေးချယ်ကာ PowerPoint (.pptx) ဖိုင် ထုတ်ယူပါ။")

EXCLUDED_KEYWORDS = ("image", "audio", "embedding", "veo", "imagen", "omni", "tts", "live")
FALLBACK_MODELS = ["gemini-3.8-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]

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
    template_name: str,
    template_file=None,
) -> bytes:
    tpl = TEMPLATES.get(template_name, TEMPLATES["Cyber Dark Cards"])
    style = tpl["style"]
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

        if template_file is None:
            bg_fill = slide.background.fill
            bg_fill.solid()
            bg_fill.fore_color.rgb = tpl["bg"]

        if style == "split":
            sidebar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(3.2), Inches(7.5)
            )
            sidebar.fill.solid()
            sidebar.fill.fore_color.rgb = tpl["side_bg"]
            sidebar.line.fill.background()

            side_tag = slide.shapes.add_textbox(Inches(0.3), Inches(0.6), Inches(2.6), Inches(1.2))
            st_tf = side_tag.text_frame
            st_tf.word_wrap = True
            sp = st_tf.paragraphs[0]
            sp.alignment = PP_ALIGN.LEFT
            srun = sp.add_run()
            srun.text = f"{icon_emoji}\n{tag_label}"
            _set_run_font(srun, font_name, 16, tpl["accent"], bold=True)

            title_box = slide.shapes.add_textbox(Inches(3.55), Inches(0.5), Inches(9.2), Inches(1.15))
            t_tf = title_box.text_frame
            t_tf.word_wrap = True
            tp = t_tf.paragraphs[0]
            tp.alignment = PP_ALIGN.LEFT
            trun = tp.add_run()
            trun.text = title_text
            _set_run_font(trun, font_name, 26 if idx > 0 else 30, tpl["title"], bold=True)

            line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(3.55), Inches(1.72), Inches(9.2), Inches(0.03)
            )
            line.fill.solid()
            line.fill.fore_color.rgb = tpl["accent"]
            line.line.fill.background()

            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.55), Inches(1.95), Inches(9.2), Inches(4.85)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = tpl["card_bg"]
            card.line.color.rgb = tpl["accent"]
            card.line.width = Pt(1)

            c_tf = card.text_frame
            c_tf.word_wrap = True
            c_tf.margin_left = Inches(0.35)
            c_tf.margin_right = Inches(0.35)
            c_tf.margin_top = Inches(0.28)
            c_tf.margin_bottom = Inches(0.25)

            for b_idx, b_text in enumerate(bullets):
                p = c_tf.paragraphs[0] if b_idx == 0 else c_tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(12)
                p.line_spacing = 1.25

                dot_run = p.add_run()
                dot_run.text = "▸  "
                _set_run_font(dot_run, font_name, 18, tpl["accent"], bold=True)

                txt_run = p.add_run()
                txt_run.text = b_text
                _set_run_font(txt_run, font_name, 18, tpl["body"], bold=False)

        elif style == "left_accent":
            left_bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.28), Inches(7.5)
            )
            left_bar.fill.solid()
            left_bar.fill.fore_color.rgb = tpl["accent"]
            left_bar.line.fill.background()

            badge = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(0.42), Inches(2.6), Inches(0.42)
            )
            badge.fill.solid()
            badge.fill.fore_color.rgb = tpl["badge_bg"]
            badge.line.fill.background()
            b_tf = badge.text_frame
            b_tf.word_wrap = True
            bp = b_tf.paragraphs[0]
            bp.alignment = PP_ALIGN.LEFT
            brun = bp.add_run()
            brun.text = f"  {icon_emoji}  {tag_label}"
            _set_run_font(brun, font_name, 12, tpl["badge_fg"], bold=True)

            title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.95), Inches(12.0), Inches(1.0))
            t_tf = title_box.text_frame
            t_tf.word_wrap = True
            tp = t_tf.paragraphs[0]
            tp.alignment = PP_ALIGN.LEFT
            trun = tp.add_run()
            trun.text = f"{icon_emoji} {title_text}"
            _set_run_font(trun, font_name, 26 if idx > 0 else 32, tpl["title"], bold=True)

            card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(2.05), Inches(12.0), Inches(4.7)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = tpl["card_bg"]
            card.line.color.rgb = RGBColor(203, 213, 225)
            card.line.width = Pt(1)

            c_tf = card.text_frame
            c_tf.word_wrap = True
            c_tf.margin_left = Inches(0.35)
            c_tf.margin_right = Inches(0.35)
            c_tf.margin_top = Inches(0.3)
            c_tf.margin_bottom = Inches(0.25)

            for b_idx, b_text in enumerate(bullets):
                p = c_tf.paragraphs[0] if b_idx == 0 else c_tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(12)
                p.line_spacing = 1.25

                dot_run = p.add_run()
                dot_run.text = "•  "
                _set_run_font(dot_run, font_name, 18, tpl["accent"], bold=True)

                txt_run = p.add_run()
                txt_run.text = b_text
                _set_run_font(txt_run, font_name, 18, tpl["body"], bold=False)

        elif style == "minimal":
            badge = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(0.45), Inches(2.5), Inches(0.4)
            )
            badge.fill.solid()
            badge.fill.fore_color.rgb = tpl["badge_bg"]
            badge.line.fill.background()
            b_tf = badge.text_frame
            b_tf.word_wrap = True
            bp = b_tf.paragraphs[0]
            bp.alignment = PP_ALIGN.LEFT
            brun = bp.add_run()
            brun.text = f"  {icon_emoji}  {tag_label}"
            _set_run_font(brun, font_name, 11, tpl["badge_fg"], bold=True)

            title_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.98), Inches(11.8), Inches(1.0))
            t_tf = title_box.text_frame
            t_tf.word_wrap = True
            tp = t_tf.paragraphs[0]
            tp.alignment = PP_ALIGN.LEFT
            trun = tp.add_run()
            trun.text = f"{icon_emoji} {title_text}"
            _set_run_font(trun, font_name, 26 if idx > 0 else 32, tpl["title"], bold=True)

            line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(2.02), Inches(11.8), Inches(0.025)
            )
            line.fill.solid()
            line.fill.fore_color.rgb = tpl["accent"]
            line.line.fill.background()

            body_box = slide.shapes.add_textbox(Inches(0.75), Inches(2.22), Inches(11.8), Inches(4.55))
            c_tf = body_box.text_frame
            c_tf.word_wrap = True

            for b_idx, b_text in enumerate(bullets):
                p = c_tf.paragraphs[0] if b_idx == 0 else c_tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(14)
                p.line_spacing = 1.25

                dot_run = p.add_run()
                dot_run.text = "▸  "
                _set_run_font(dot_run, font_name, 18, tpl["accent"], bold=True)

                txt_run = p.add_run()
                txt_run.text = b_text
                _set_run_font(txt_run, font_name, 18, tpl["body"], bold=False)

        else:
            top_bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12)
            )
            top_bar.fill.solid()
            top_bar.fill.fore_color.rgb = tpl["accent"]
            top_bar.line.fill.background()

            badge = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(0.38), Inches(2.55), Inches(0.45)
            )
            badge.fill.solid()
            badge.fill.fore_color.rgb = tpl["badge_bg"]
            badge.line.fill.background()
            b_tf = badge.text_frame
            b_tf.word_wrap = True
            bp = b_tf.paragraphs[0]
            bp.alignment = PP_ALIGN.LEFT
            brun = bp.add_run()
            brun.text = f"  {icon_emoji}  {tag_label}"
            _set_run_font(brun, font_name, 12, tpl["badge_fg"], bold=True)

            title_box = slide.shapes.add_textbox(Inches(0.65), Inches(0.95), Inches(12.0), Inches(0.95))
            t_tf = title_box.text_frame
            t_tf.word_wrap = True
            tp = t_tf.paragraphs[0]
            tp.alignment = PP_ALIGN.LEFT
            trun = tp.add_run()
            trun.text = f"{icon_emoji} {title_text}"
            _set_run_font(trun, font_name, 26 if idx > 0 else 32, tpl["title"], bold=True)

            line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(1.95), Inches(12.03), Inches(0.03)
            )
            line.fill.solid()
            line.fill.fore_color.rgb = tpl["accent"]
            line.line.fill.background()

            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(2.15), Inches(12.03), Inches(4.65)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = tpl["card_bg"]
            card.line.color.rgb = tpl["accent"]
            card.line.width = Pt(1)

            c_tf = card.text_frame
            c_tf.word_wrap = True
            c_tf.margin_left = Inches(0.35)
            c_tf.margin_right = Inches(0.35)
            c_tf.margin_top = Inches(0.28)
            c_tf.margin_bottom = Inches(0.25)

            for b_idx, b_text in enumerate(bullets):
                p = c_tf.paragraphs[0] if b_idx == 0 else c_tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(12)
                p.line_spacing = 1.25

                dot_run = p.add_run()
                dot_run.text = "▸  "
                _set_run_font(dot_run, font_name, 18, tpl["accent"], bold=True)

                txt_run = p.add_run()
                txt_run.text = b_text
                _set_run_font(txt_run, font_name, 18, tpl["body"], bold=False)

        footer_box = slide.shapes.add_textbox(Inches(0.65), Inches(6.95), Inches(12.03), Inches(0.35))
        f_tf = footer_box.text_frame
        fp = f_tf.paragraphs[0]
        fp.alignment = PP_ALIGN.RIGHT
        frun = fp.add_run()
        frun.text = f"Slide {idx + 1} / {total_slides}"
        _set_run_font(frun, font_name, 11, tpl["muted"], bold=False)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


api_key = get_api_key()

col_u1, col_u2 = st.columns([2, 1])
with col_u1:
    uploaded_files = st.file_uploader(
        "PDF (.pdf) သို့မဟုတ် Word (.docx) ဖိုင်များ တင်ပါ",
        type=["pdf", "docx"],
        accept_multiple_files=True,
    )
with col_u2:
    custom_template = st.file_uploader(
        "ကိုယ်ပိုင် PPT Template (.pptx) - Optional",
        type=["pptx"],
        accept_multiple_files=False,
    )

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    language_choice = st.radio(
        "ဘာသာစကား",
        options=["မြန်မာဘာသာ (Myanmar)", "English"],
        index=0,
    )
with c2:
    slide_count = st.slider("Slide အရေအတွက်", min_value=4, max_value=15, value=7)
with c3:
    template_choice = st.selectbox("Template ရွေးရန်", options=list(TEMPLATES.keys()), index=0)
with c4:
    mm_font = st.text_input("မြန်မာ PPT ဖောင့်", value="Pyidaungsu")
with c5:
    selected_model = st.selectbox("Gemini Model", options=MODEL_OPTIONS, index=0)

extra_note = st.text_area(
    "ထပ်မံထည့်သွင်းလိုသည့် ညွှန်ကြားချက် (Optional)",
    placeholder="ဥပမာ - Security awareness / knowledge sharing presentation ပုံစံ၊ အဓိက အလေးပေးလိုသော အချက်များ...",
    height=75,
)

if st.button("🎨 လှပသော PowerPoint ထုတ်မည်", type="primary", use_container_width=True):
    if not api_key:
        st.error("Gemini API Key ထည့်သွင်းရန် လိုအပ်ပါသည်။")
    elif not uploaded_files:
        st.error("PDF သို့မဟုတ် Word ဖိုင် အနည်းဆုံးတစ်ခု Upload တင်ပါ။")
    else:
        with st.spinner(f"'{selected_model}' ဖြင့် Slide များနှင့် '{template_choice}' Template ပြုစုနေပါသည်..."):
            try:
                extracted_chunks = []
                for f in uploaded_files:
                    fname = f.name.lower()
                    if fname.endswith(".pdf"):
                        extracted_chunks.append(extract_pdf_text(f))
                    elif fname.endswith(".docx"):
                        extracted_chunks.append(extract_docx_text(f))

                combined_text = "\n\n".join(extracted_chunks).strip()
                if not combined_text:
                    st.warning("တင်ထားသော ဖိုင်များထဲမှ စာသား ရှာမတွေ့ပါ (Scanned PDF ပုံများသာ ဖြစ်နေနိုင်ပါသည်)။")
                else:
                    slides = generate_slides_json(
                        api_key=api_key,
                        source_text=combined_text,
                        slide_count=slide_count,
                        language_choice=language_choice,
                        extra_note=extra_note,
                        model_name=selected_model,
                    )

                    is_mm = language_choice == "မြန်မာဘာသာ (Myanmar)"
                    font_to_use = (mm_font.strip() or "Pyidaungsu") if is_mm else "Calibri"
                    file_name = "presentation_myanmar.pptx" if is_mm else "presentation_english.pptx"

                    pptx_bytes = build_styled_pptx_bytes(
                        slides=slides,
                        font_name=font_to_use,
                        template_name=template_choice,
                        template_file=custom_template,
                    )

                    st.session_state["slides_result"] = slides
                    st.session_state["pptx_bytes"] = pptx_bytes
                    st.session_state["pptx_filename"] = file_name
                    st.session_state["pptx_lang_label"] = language_choice
                    st.success(f"{language_choice} ({template_choice}) PowerPoint ဖိုင် ထုတ်ပြီးပါပြီ။")
            except Exception as e:
                st.error(f"Error ဖြစ်ပွားခဲ့ပါသည်: {e}")

if "pptx_bytes" in st.session_state and "slides_result" in st.session_state:
    st.divider()
    st.download_button(
        label=f"⬇️ {st.session_state.get('pptx_lang_label', '')} PowerPoint (.pptx) ဒေါင်းလုဒ်ဆွဲရန်",
        data=st.session_state["pptx_bytes"],
        file_name=st.session_state.get("pptx_filename", "presentation.pptx"),
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

    st.subheader(f"📝 Slides Preview ({st.session_state.get('pptx_lang_label', '')})")
    all_slides = st.session_state.get("slides_result", [])
    for i, s in enumerate(all_slides):
        icon_emoji, tag_label = detect_slide_icon_and_tag(
            s.get("title", ""), s.get("bullets", []), i, len(all_slides)
        )
        st.markdown(f"**Slide {i + 1} `{tag_label}`: {icon_emoji} {s.get('title', '')}**")
        for b in s.get("bullets", []):
            st.markdown(f"- {b}")
