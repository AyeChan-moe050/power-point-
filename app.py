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
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=response_schema,
        ),
    )

    raw_text = response.text or "{}"
    return json.loads(raw_text)


def build_pptx_bytes(slides: list[dict], font_name: str) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for idx, slide_data in enumerate(slides):
        title_text = str(slide_data.get("title", "")).strip()
        bullets = [str(b).strip() for b in slide_data.get("bullets", []) if str(b).strip()]

        if idx == 0:
            layout = prs.slide_layouts[0]  # Title Slide
            slide = prs.slides.add_slide(layout)
            if slide.shapes.title:
                slide.shapes.title.text = title_text
                for p in slide.shapes.title.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.name = font_name
                        r.font.size = Pt(36)
            if len(slide.placeholders) > 1 and slide.placeholders[1].has_text_frame:
                sub_tf = slide.placeholders[1].text_frame
                sub_tf.clear()
                for i, b in enumerate(bullets[:3]):
                    p = sub_tf.paragraphs[0] if i == 0 else sub_tf.add_paragraph()
                    p.text = b
                    for r in p.runs:
                        r.font.name = font_name
                        r.font.size = Pt(20)
        else:
            layout = prs.slide_layouts[1]  # Title and Content
            slide = prs.slides.add_slide(layout)
            if slide.shapes.title:
                slide.shapes.title.text = title_text
                for p in slide.shapes.title.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.name = font_name
                        r.font.size = Pt(28)

            body_placeholder = None
            for shape in slide.placeholders:
                if shape.placeholder_format.idx == 1 and shape.has_text_frame:
                    body_placeholder = shape
                    break

            if body_placeholder:
                tf = body_placeholder.text_frame
                tf.clear()
                for i, b in enumerate(bullets):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = b
                    p.level = 0
                    for r in p.runs:
                        r.font.name = font_name
                        r.font.size = Pt(20)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


# --- UI ---
api_key = get_api_key()

uploaded_files = st.file_uploader(
    "PDF (.pdf) သို့မဟုတ် Word (.docx) ဖိုင်များ တင်ပါ",
    type=["pdf", "docx"],
    accept_multiple_files=True,
)

c1, c2 = st.columns(2)
with c1:
    slide_count = st.slider("ခန့်မှန်း Slide အရေအတွက်", min_value=4, max_value=20, value=8)
with c2:
    mm_font = st.text_input("မြန်မာ PPT ဖောင့်အမည် (PowerPoint တွင် ဖွင့်ရန်)", value="Pyidaungsu")

extra_note = st.text_area(
    "ထပ်မံထည့်သွင်းလိုသည့် ညွှန်ကြားချက် (Optional)",
    placeholder="ဥပမာ - အစိုးရဌာန အစည်းအဝေး/သင်တန်း အနှစ်ချုပ်ပုံစံ၊ အဓိက အလေးပေးလိုသော အချက်များ...",
    height=85,
)

if st.button("🚀 မြန်မာ + English PowerPoint ၂ ဖိုင် ထုတ်မည်", type="primary", use_container_width=True):
    if not api_key:
        st.error("Gemini API Key ထည့်သွင်းရန် လိုအပ်ပါသည်။")
    elif not uploaded_files:
        st.error("PDF သို့မဟုတ် Word ဖိုင် အနည်းဆုံးတစ်ခု Upload တင်ပါ။")
    else:
        with st.spinner("ဖိုင်များကို ဖတ်ရှုပြီး မြန်မာ နှင့် English PowerPoint ဖိုင်များ ပြုစုနေပါသည်..."):
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
                    slide_data = generate_bilingual_slides_json(
                        api_key=api_key,
                        source_text=combined_text,
                        slide_count=slide_count,
                        extra_note=extra_note,
                    )
                    mm_slides = slide_data.get("myanmar", [])
                    en_slides = slide_data.get("english", [])

                    mm_pptx_bytes = build_pptx_bytes(mm_slides, font_name=mm_font.strip() or "Pyidaungsu")
                    en_pptx_bytes = build_pptx_bytes(en_slides, font_name="Calibri")

                    st.session_state["mm_slides"] = mm_slides
                    st.session_state["en_slides"] = en_slides
                    st.session_state["mm_pptx"] = mm_pptx_bytes
                    st.session_state["en_pptx"] = en_pptx_bytes
                    st.success("မြန်မာ နှင့် English PowerPoint ဖိုင် ၂ ခု ထုတ်ပြီးပါပြီ။")
            except Exception as e:
                st.error(f"Error ဖြစ်ပွားခဲ့ပါသည်: {e}")

if "mm_pptx" in st.session_state and "en_pptx" in st.session_state:
    st.divider()
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label="⬇️ မြန်မာ PowerPoint (.pptx) ဒေါင်းလုဒ်ဆွဲရန်",
            data=st.session_state["mm_pptx"],
            file_name="presentation_myanmar.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            label="⬇️ English PowerPoint (.pptx) ဒေါင်းလုဒ်ဆွဲရန်",
            data=st.session_state["en_pptx"],
            file_name="presentation_english.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    tab_mm, tab_en = st.tabs(["🇲🇲 မြန်မာ Slides Preview", "🇬🇧 English Slides Preview"])
    with tab_mm:
        for i, s in enumerate(st.session_state.get("mm_slides", []), start=1):
            st.markdown(f"**Slide {i}: {s.get('title', '')}**")
            for b in s.get("bullets", []):
                st.markdown(f"- {b}")
    with tab_en:
        for i, s in enumerate(st.session_state.get("en_slides", []), start=1):
            st.markdown(f"**Slide {i}: {s.get('title', '')}**")
            for b in s.get("bullets", []):
                st.markdown(f"- {b}")
