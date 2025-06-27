# app.py
# Streamlit Hybrid Style-Copy PPTX Extractor

import streamlit as st
import subprocess
import tempfile
import os
import shutil
from pptx import Presentation
from pptx.dml.color import RGBColor

st.set_page_config(page_title="Hybrid Style-Copy PPTX Extractor")
st.title("Hybrid Style-Copy PPTX Extractor")

# --- Helper functions ---

def parse_pptx(path: str) -> dict:
    prs = Presentation(path)
    slides = []
    for idx, slide in enumerate(prs.slides):
        elements = []
        for i, shape in enumerate(slide.shapes):
            tf = getattr(shape, 'text_frame', None)
            text = ''
            if tf:
                text = ''.join(run.text for p in tf.paragraphs for run in p.runs)
            elements.append({
                'shape_idx': i,
                'text': text
            })
        slides.append({
            'slide_index': idx,
            'elements': elements
        })
    return {'slides': slides}


def convert_pdf_to_pptx(pdf_path: str) -> str:
    # Requires unoconv installed
    pptx_path = pdf_path.replace('.pdf', '.pptx')
    subprocess.run(['unoconv', '-f', 'pptx', pdf_path], check=True)
    return pptx_path


def copy_shape_style(src_shape, tgt_shape):
    # Copy fill
    if src_shape.fill.type == 1:
        tgt_shape.fill.solid()
        rgb = src_shape.fill.fore_color.rgb
        tgt_shape.fill.fore_color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    # Copy line
    if src_shape.line:
        tgt_shape.line.fill.solid()
        clr = src_shape.line.color.rgb
        tgt_shape.line.color.rgb = RGBColor(clr[0], clr[1], clr[2])
        tgt_shape.line.width = src_shape.line.width
    # Copy text & font
    if src_shape.has_text_frame and tgt_shape.has_text_frame:
        src_tf = src_shape.text_frame
        tgt_tf = tgt_shape.text_frame
        tgt_tf.text = src_tf.text
        for src_run, tgt_run in zip(src_tf.paragraphs[0].runs, tgt_tf.paragraphs[0].runs):
            tgt_run.font.name = src_run.font.name
            tgt_run.font.size = src_run.font.size
            tgt_run.font.bold = src_run.font.bold
            tgt_run.font.italic = src_run.font.italic
            try:
                tgt_run.font.color.rgb = src_run.font.color.rgb
            except:
                pass

# --- Main UI ---
uploaded = st.file_uploader('Upload PPTX or PDF', type=['pptx', 'pdf'])
if uploaded:
    # Save uploaded file
    suffix = os.path.splitext(uploaded.name)[1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
        tmp_file.write(uploaded.getbuffer())
        tmp_path = tmp_file.name
    # Convert PDF if needed
    if suffix == '.pdf':
        try:
            tmp_path = convert_pdf_to_pptx(tmp_path)
        except Exception as e:
            st.error(f"PDF→PPTX conversion failed: {e}")
            st.stop()

    # Preview parsed JSON
    try:
        slides_json = parse_pptx(tmp_path)
        st.json(slides_json)
    except Exception as e:
        st.error(f"Parsing failed: {e}")
        st.stop()

    # Hybrid regenerate
    if st.button('Hybrid Regenerate'):
        prs = Presentation(tmp_path)
        new_prs = Presentation()
        # Copy each slide
        for slide in prs.slides:
            new_slide = new_prs.slides.add_slide(new_prs.slide_layouts[5])
            for shape in slide.shapes:
                # Add same shape type and geometry
                try:
                    new_shape = new_slide.shapes.add_shape(
                        shape.auto_shape_type,
                        shape.left, shape.top,
                        shape.width, shape.height
                    )
                    copy_shape_style(shape, new_shape)
                except Exception:
                    # Skip unsupported shapes
                    continue
        # Save & offer download
        out_path = tempfile.mktemp(suffix='.pptx')
        new_prs.save(out_path)
        with open(out_path, 'rb') as f:
            data = f.read()
        st.download_button('Download Recreated PPTX', data=data,
                           file_name='hybrid_recreated.pptx',
                           mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')
