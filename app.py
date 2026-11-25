import io
import json
import os
import time
import streamlit as st
from docx import Document
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pypdf import PdfReader

st.set_page_config(
    page_title="PDF / Word မှ PowerPoint ထုတ်ပေးသည့် စနစ်",
    page_icon="📊",
    layout="wide",
)

st.title("📊 PDF / Word မှ PowerPoint ထုတ်ပေးသည့် စနစ်")
st.caption("PDF နှင့် Word (.docx) ဖိုင်များကို တင်ပြီး မြန်မာ (သို့မဟုတ်) English PowerPoint ၁ ဖိုင် ရွေးချယ်ထုတ်ယူပါ။")


def get_api_key() -> str:
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    if "GEMINI_API_KEY" in os.environ:
        return os.environ["GEMINI_API_KEY"]
    return st.sidebar.text_input("Gemini API Key ထည့်ပါ", type="password")


@st.cache_data(ttl=600, show_spinner=False)
def list_available_gemini_models(api_key: str) -> list[str]:
    client = genai.Client(api_key=api_key)
    models = []
    for m in client.models.list():
        name = getattr(m, "name", "") or ""
        short_name = name.replace("models/", "")
        actions = getattr(m, "supported_actions", None) or []
        if "generateContent" in actions or not actions:
            if "gemini" in short_name and short_name not in models:
                models.append(short_name)
    # gemini-3.8-flash ကို ဦးစားပေး အပေါ်ဆုံးမှာ ထားရန်
    models.sort(key=lambda x: (0 if "3.8" in x else 1 if "flash" in x else 2, x))
    return models or ["gemini-3.8-flash"]


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
) -> list[dict]:
    client = genai.Client(api_key=api_key)
    trimmed_text = source_text[:30000]

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
    for attempt in range(3):
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
            if ("503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str) and attempt < 2:
                time.sleep(2 * (attempt + 1))
                continue
            raise

    raise RuntimeError(f"Error: {last_err}")


def build_pptx_bytes(slides: list[dict], font_name: str) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for idx, slide_data in enumerate(slides):
        title_text = str(slide_data.get("title", "")).strip()
        bullets = [str(b).strip() for b in slide_data.get("bullets", []) if str(b).strip()]

        if idx == 0:
            layout = prs.slide_layouts[0]
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
            layout = prs.slide_layouts[1]
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


api_key = get_api_key()

model_options = ["gemini-3.8-flash"]
if api_key:
    try:
        model_options = list_available_gemini_models(api_key)
    except Exception:
        pass

uploaded_files = st.file_uploader(
    "PDF (.pdf) သို့မဟုတ် Word (.docx) ဖိုင်များ တင်ပါ",
    type=["pdf", "docx"],
    accept_multiple_files=True,
)

c1, c2, c3, c4 = st.columns(4)
with c1:
    language_choice = st.radio(
        "ထုတ်မည့် PowerPoint ဘာသာစကား",
        options=["မြန်မာဘာသာ (Myanmar)", "English"],
        index=0,
    )
with c2:
    slide_count = st.slider("ခန့်မှန်း Slide အရေအတွက်", min_value=4, max_value=20, value=6)
with c3:
    mm_font = st.text_input("မြန်မာ PPT ဖောင့်အမည်", value="Pyidaungsu")
with c4:
    selected_model = st.selectbox("Gemini Model ရွေးရန်", options=model_options, index=0)

custom_model = st.text_input(
    "Model အမည် ကိုယ်တိုင်ရိုက်ထည့်လိုပါက (မထည့်ပါက အပေါ်က ရွေးထားတာကို သုံးပါမည်)",
    placeholder="ဥပမာ - gemini-3.8-flash",
)
active_model = custom_model.strip() or selected_model

extra_note = st.text_area(
    "ထပ်မံထည့်သွင်းလိုသည့် ညွှန်ကြားချက် (Optional)",
    placeholder="ဥပမာ - အစိုးရဌာန အစည်းအဝေး/သင်တန်း အနှစ်ချုပ်ပုံစံ၊ အဓိက အလေးပေးလိုသော အချက်များ...",
    height=80,
)

if st.button("🚀 PowerPoint ထုတ်မည်", type="primary", use_container_width=True):
    if not api_key:
        st.error("Gemini API Key ထည့်သွင်းရန် လိုအပ်ပါသည်။")
    elif not uploaded_files:
        st.error("PDF သို့မဟုတ် Word ဖိုင် အနည်းဆုံးတစ်ခု Upload တင်ပါ။")
    else:
        with st.spinner(f"'{active_model}' ဖြင့် PowerPoint ဖိုင် ပြုစုနေပါသည်..."):
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
                        model_name=active_model,
                    )

                    is_mm = language_choice == "မြန်မာဘာသာ (Myanmar)"
                    font_to_use = (mm_font.strip() or "Pyidaungsu") if is_mm else "Calibri"
                    file_name = "presentation_myanmar.pptx" if is_mm else "presentation_english.pptx"

                    pptx_bytes = build_pptx_bytes(slides, font_name=font_to_use)

                    st.session_state["slides_result"] = slides
                    st.session_state["pptx_bytes"] = pptx_bytes
                    st.session_state["pptx_filename"] = file_name
                    st.session_state["pptx_lang_label"] = language_choice
                    st.success(f"{language_choice} PowerPoint ဖိုင် ထုတ်ပြီးပါပြီ။")
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
    for i, s in enumerate(st.session_state.get("slides_result", []), start=1):
        st.markdown(f"**Slide {i}: {s.get('title', '')}**")
        for b in s.get("bullets", []):
            st.markdown(f"- {b}")
