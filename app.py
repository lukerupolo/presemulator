# app.py
# Hybrid Style-Copy PPTX Extractor with Enhanced Color Handling

import streamlit as st
import subprocess
import tempfile
import os
import shutil
from pptx import Presentation
from pptx.dml.color import RGBColor

st.set_page_config(page_title="Hybrid Style-Copy PPTX Extractor")
st.title("Hybrid Style-Copy PPTX Extractor with Enhanced Color Handling")

# --- Helper Functions ---

def convert_pdf_to_pptx(pdf_path: str) -> str:
    """
    Convert PDF to PPTX using unoconv. Requires libreoffice/unoconv installed.
    """
    pptx_path = pdf_path.replace('.pdf', '.pptx')
    subprocess.run(['unoconv', '-f', 'pptx', pdf_path], check=True)
    return pptx_path


def parse_pptx(path: str) -> dict:
    """
    Parse a PPTX into JSON for preview: slide index and texts of each shape.
    """
    prs = Presentation(path)
    slides = []
    for idx, slide in enumerate(prs.slides):
        elements = []
        for i, shape in enumerate(slide.shapes):
            tf = getattr(shape, 'text_frame', None)
            text = ''
            if tf:
                text = ''.join(run.text for p in tf.paragraphs for run in p.runs)
            elements.append({'shape_idx': i, 'text': text})
        slides.append({'slide_index': idx, 'elements': elements})
    return {'slides': slides}


def copy_shape_style(src_shape, tgt_shape):
    """
    Copy fill, line, and text/font properties from src_shape to tgt_shape,
    safely handling NoneColor and missing attributes.
    """
    # Fill (solid only)
    try:
        if getattr(src_shape, 'fill', None) and src_shape.fill.type == 1:
            tgt_shape.fill.solid()
            rgb = src_shape.fill.fore_color.rgb
            tgt_shape.fill.fore_color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    except Exception:
        pass

    # Line
    try:
        if getattr(src_shape, 'line', None) and getattr(src_shape.line, 'fill', None) and src_shape.line.fill.type == 1:
            tgt_shape.line.fill.solid()
            rgb = src_shape.line.color.rgb
            tgt_shape.line.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
            tgt_shape.line.width = src_shape.line.width
    except Exception:
        pass

    # Text & Font
    if getattr(src_shape, 'has_text_frame', False) and getattr(tgt_shape, 'has_text_frame', False):
        src_tf = src_shape.text_frame
        tgt_tf = tgt_shape.text_frame
        # Copy full text
        tgt_tf.text = src_tf.text
        # Copy first paragraph runs
        for src_run, tgt_run in zip(src_tf.paragraphs[0].runs, tgt_tf.paragraphs[0].runs):
            try:
                tgt_run.font.name = src_run.font.name
                tgt_run.font.size = src_run.font.size
                tgt_run.font.bold = src_run.font.bold
                tgt_run.font.italic = src_run.font.italic
                tgt_run.font.color.rgb = src_run.font.color.rgb
            except Exception:
                pass

# --- Main App Logic ---
uploaded = st.file_uploader('Upload PPTX or PDF', type=['pptx', 'pdf'])
if not uploaded:
    st.info('Please upload a PowerPoint (.pptx) or PDF file.')
    st.stop()

# Save uploaded file to a temp path
suffix = os.path.splitext(uploaded.name)[1].lower()
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
    tmp_file.write(uploaded.getbuffer())
    tmp_path = tmp_file.name

# Convert PDF if necessary
if suffix == '.pdf':
    try:
        tmp_path = convert_pdf_to_pptx(tmp_path)
    except Exception as e:
        st.error(f"PDF→PPTX conversion failed: {e}")
        st.stop()

# Load the presentation
try:
    prs = Presentation(tmp_path)
except Exception as e:
    st.error(f"Failed to load PPTX: {e}")
    st.stop()

# Debug: Show slide and shape counts
st.write(f"Found {len(prs.slides)} slide(s)")
for idx, slide in enumerate(prs.slides):
    st.write(f"Slide {idx}: {len(slide.shapes)} shape(s)")

# JSON preview of parsed slide contents
slides_json = parse_pptx(tmp_path)
st.subheader('Parsed Slide Contents')
st.json(slides_json)

# Hybrid Regenerate
if st.button('Hybrid Regenerate'):
    new_prs = Presentation()
    # Choose a blank layout (no placeholders)
    blank_layout = next((l for l in new_prs.slide_layouts if not l.placeholders), new_prs.slide_layouts[0])

    for slide_idx, slide in enumerate(prs.slides):
        st.write(f"Regenerating slide {slide_idx}")
        new_slide = new_prs.slides.add_slide(blank_layout)
        for shape in slide.shapes:
            # Attempt to get an autoshape type
            try:
                shape_type = shape.auto_shape_type
            except Exception:
                st.warning(f"Skipping non-autoshape (shape_idx={getattr(shape, 'shape_id', 'n/a')})")
                continue
            st.write(f"  Copying shape idx={getattr(shape, 'shape_id', 'n/a')} type={shape_type}")
            try:
                new_shape = new_slide.shapes.add_shape(
                    shape_type,
                    shape.left, shape.top,
                    shape.width, shape.height
                )
                copy_shape_style(shape, new_shape)
            except Exception as e:
                st.error(f"Failed to copy shape (shape_idx={getattr(shape, 'shape_id', 'n/a')}): {e}")
                continue

    # Save and output the regenerated deck
    out_path = tempfile.mktemp(suffix='.pptx')
    new_prs.save(out_path)
    with open(out_path, 'rb') as f:
        data = f.read()
    st.success('Recreated PPTX is ready!')
    st.download_button(
        'Download Recreated PPTX', data=data,
        file_name='hybrid_recreated.pptx',
        mime='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )
