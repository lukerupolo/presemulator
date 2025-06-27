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
def sanitize_code(code: str) -> str:
    """
    Strip markdown fences and narrative prefixes from the AI-generated code.
    """
    lines = code.splitlines()
    # Remove markdown fences
    if lines and lines[0].startswith("```"):
        lines = lines[1:]
    if lines and lines[-1].startswith("```"):
        lines = lines[:-1]
    # Remove narrative lines before actual Python code
    for i, line in enumerate(lines):
        if line.startswith("import ") or line.startswith("from ") or line.startswith("def "):
            lines = lines[i:]
            break
    return "\n".join(lines)


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
            tf = getattr(shape, 'text_frame', None)
            if not tf:
                continue
            text = ''.join(run.text for p in tf.paragraphs for run in p.runs)
            font_info = {'name': None, 'size': None, 'bold': False, 'italic': False}
            runs = tf.paragraphs[0].runs if tf.paragraphs else []
            if runs:
                fnt = runs[0].font
                font_info = {
                    'name': fnt.name,
                    'size': fnt.size.pt if fnt.size else None,
                    'bold': bool(fnt.bold),
                    'italic': bool(fnt.italic)
                }
            elements.append({
                'text': text,
                'font': font_info,
                'position': {'x': shape.left.pt, 'y': shape.top.pt,
                             'w': shape.width.pt, 'h': shape.height.pt}
            })
        slides.append({'layout': layout, 'background_color': bg_color, 'elements': elements})
    return {'slides': slides}


def convert_pdf_to_pptx(pdf_path: str) -> str:
    pptx_path = pdf_path.replace('.pdf', '.pptx')
    subprocess.run(['unoconv', '-f', 'pptx', pdf_path], check=True)
    return pptx_path


def call_openai_to_generate_code(slides_json: dict) -> str:
    system = (
        "You are a code generator. Given JSON describing slides, produce Python code using python-pptx"
        " that recreates each slide with fonts, positions, and colors."
    )
    response = openai.chat.completions.create(
        model="gpt-4",
        messages=[
            {'role': 'system', 'content': system},
            {'role': 'user', 'content': json.dumps(slides_json, indent=2)}
        ],
        temperature=0
    )
    return response.choices[0].message.content


def run_generated_code(code: str) -> str:
    """
    Clean, run the generated script, and return the rebuilt PPTX's path.
    """
    code_clean = sanitize_code(code)
    with tempfile.TemporaryDirectory() as tmp:
        script = os.path.join(tmp, 'gen.py')
        with open(script, 'w') as f:
            f.write(code_clean)
        try:
            result = subprocess.run(
                ['python', script], cwd=tmp,
                capture_output=True, text=True, check=True
            )
        except subprocess.CalledProcessError as e:
            st.error('Error executing generated code:')
            st.code(e.stderr)
            raise
        out = os.path.join(tmp, 'recreated.pptx')
        if not os.path.exists(out):
            st.error("Expected 'recreated.pptx' not found.")
            raise FileNotFoundError(out)
        final = tempfile.mktemp(suffix='.pptx')
        shutil.copy(out, final)
        return final

# --- Main UI ---
uploaded = st.file_uploader('Upload PPTX or PDF', type=['pptx', 'pdf'])
if uploaded:
    suffix = os.path.splitext(uploaded.name)[1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmpf:
        tmpf.write(uploaded.getbuffer())
        tmp_path = tmpf.name
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

    if st.button('Generate New PPTX'):
        with st.spinner('Calling OpenAI...'):
            try:
                code = call_openai_to_generate_code(slides_json)
            except Exception as e:
                st.error(f"OpenAI call failed: {e}")
                st.stop()
        with st.spinner('Reconstructing slides...'):
            try:
                out_file = run_generated_code(code)
            except Exception:
                st.stop()
        st.success('Done! Download below:')
        data = open(out_file, 'rb').read()
        st.download_button(
            'Download PPTX', data=data,
            file_name='recreated.pptx',
            mime='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
