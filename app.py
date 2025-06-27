import streamlit as st
import openai
import subprocess
import tempfile
import json
import shutil
import os
from pptx import Presentation

# --- Sidebar: OpenAI key ---
st.set_page_config(page_title="AI PPTX Style Extractor")
st.title("AI PPTX Style Extractor")
api_key = st.sidebar.text_input("OpenAI API Key", type="password")
if not api_key:
    st.sidebar.warning("Enter your OpenAI API key")
    st.stop()
openai.api_key = api_key

# --- Helper functions ---
def parse_pptx(path: str) -> dict:
    prs = Presentation(path)
    slides = []
    for slide in prs.slides:
        layout = slide.slide_layout.name if slide.slide_layout else "Custom"
        bg = slide.background.fill
        bg_color = None
        if bg and bg.type == 1:
            rgb = bg.fore_color.rgb
            bg_color = f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        elements = []
        for shape in slide.shapes:
            if not hasattr(shape, 'text_frame') or not shape.text_frame:
                continue
            # Extract full text safely
            full_text = ''
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    full_text += run.text
            # Font from first run if exists
            font_info = {'name': None, 'size': None, 'bold': False, 'italic': False}
            if shape.text_frame.paragraphs:
                first_para = shape.text_frame.paragraphs[0]
                if first_para.runs:
                    font = first_para.runs[0].font
                    font_info = {
                        'name': font.name,
                        'size': font.size.pt if font.size else None,
                        'bold': bool(font.bold),
                        'italic': bool(font.italic)
                    }
            elements.append({
                "text": full_text,
                "font": font_info,
                "position": {"x": shape.left.pt, "y": shape.top.pt,
                             "w": shape.width.pt, "h": shape.height.pt}
            })
        slides.append({
            "layout": layout,
            "background_color": bg_color,
            "elements": elements
        })
    return {"slides": slides}


def convert_pdf_to_pptx(pdf_path: str) -> str:
    pptx_path = pdf_path.replace('.pdf', '.pptx')
    subprocess.run(['unoconv', '-f', 'pptx', pdf_path], check=True)
    return pptx_path


def call_openai_to_generate_code(slides_json: dict) -> str:
    system = (
        "You are a code generator. Given JSON describing slides, produce Python code using python-pptx"
        " that recreates each slide with fonts, positions, and colors."
    )
    resp = openai.ChatCompletion.create(
        model="gpt-4o-code",
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": json.dumps(slides_json, indent=2)}],
        temperature=0
    )
    return resp.choices[0].message.content


def run_generated_code(code: str) -> str:
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, 'gen.py')
        with open(path, 'w') as f:
            f.write(code)
        subprocess.run(['python', path], cwd=tmp, check=True)
        out = os.path.join(tmp, 'recreated.pptx')
        final = tempfile.mktemp(suffix='.pptx')
        shutil.copy(out, final)
        return final

# --- Main UI ---
uploaded = st.file_uploader("Upload PPTX or PDF", type=["pptx","pdf"])
if uploaded:
    suffix = os.path.splitext(uploaded.name)[1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded.getbuffer())
        tmp_path = tmp.name
    if suffix == '.pdf':
        try:
            tmp_path = convert_pdf_to_pptx(tmp_path)
        except Exception as e:
            st.error(f"PDF→PPTX conversion failed: {e}")
            st.stop()
    try:
        slides_json = parse_pptx(tmp_path)
    except Exception as e:
        st.error(f"Parsing failed: {e}")
        st.stop()

    if st.button("Generate New PPTX"):
        with st.spinner("Calling OpenAI..."):
            try:
                code = call_openai_to_generate_code(slides_json)
            except Exception as e:
                st.error(f"OpenAI call failed: {e}")
                st.stop()
        with st.spinner("Reconstructing slides..."):
            try:
                out_file = run_generated_code(code)
            except Exception as e:
                st.error(f"Reconstruction failed: {e}")
                st.stop()
        st.success("Done! Download below:")
        with open(out_file, 'rb') as f:
            data = f.read()
        st.download_button("Download PPTX", data=data,
                           file_name="recreated.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
