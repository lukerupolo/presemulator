import streamlit as st
import subprocess
import tempfile
import os
import json
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
import openai

# --- App Config ---
st.set_page_config(page_title="AI Slide Generator with Template Style", layout="wide")
st.title("AI Slide Generator with Template Style")

# --- Sidebar: OpenAI API Key ---
api_key = st.sidebar.text_input("OpenAI API Key", type="password")
if not api_key:
    st.sidebar.warning("Please enter your OpenAI API Key to proceed.")
    st.stop()
openai.api_key = api_key

# --- Upload Template ---
uploaded = st.file_uploader("Upload a PPTX or PDF template", type=["pptx", "pdf"])
if not uploaded:
    st.info("Awaiting PPTX or PDF template upload...")
    st.stop()

# Save uploaded file
suffix = os.path.splitext(uploaded.name)[1].lower()
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
    tmp.write(uploaded.getbuffer())
    template_path = tmp.name

# Convert PDF to PPTX if needed
if suffix == ".pdf":
    try:
        converted = template_path.replace('.pdf', '.pptx')
        subprocess.run(["unoconv", "-f", "pptx", template_path], check=True)
        template_path = converted
    except Exception as e:
        st.error(f"PDF → PPTX conversion failed: {e}")
        st.stop()

# --- Slide Generation Parameters ---
prompt = st.text_area("Enter a prompt for slide content generation", height=150)
num_slides = st.number_input("Number of slides to generate", min_value=1, max_value=50, value=5)

if st.button("Generate Slides"):
    if not prompt.strip():
        st.error("Please provide a prompt for content generation.")
        st.stop()

    # --- Generate Outline via OpenAI ---
    with st.spinner("Generating slide outline with OpenAI..."):
        system_msg = (
            "You are an assistant that creates PowerPoint slide outlines. "
            "Given a user prompt, generate a JSON object with a 'slides' array. "
            "Each slide entry: { 'title': string, 'bullets': [string, ...] }."
        )
        user_msg = f"Prompt: {prompt}\nGenerate {num_slides} slides. Return only valid JSON."
        try:
            resp = openai.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "system", "content": system_msg},
                          {"role": "user", "content": user_msg}],
                temperature=0.7
            )
            outline = json.loads(resp.choices[0].message.content)
        except Exception as e:
            st.error(f"OpenAI API or JSON parse error: {e}")
            st.stop()

    # --- Load Template and Select Layout ---
    try:
        prs = Presentation(template_path)
    except Exception as e:
        st.error(f"Failed to load template PPTX: {e}")
        st.stop()

    # find a layout with title and body placeholders
    def select_layout(pres):
        for layout in pres.slide_layouts:
            types = [ph.placeholder_format.type for ph in layout.placeholders]
            if PP_PLACEHOLDER.TITLE in types and PP_PLACEHOLDER.BODY in types:
                return layout
        return pres.slide_layouts[0]

    layout = select_layout(prs)

    # --- Create Slides ---
    try:
        for slide_def in outline.get('slides', []):
            slide = prs.slides.add_slide(layout)
            # fill title
            for ph in slide.placeholders:
                if ph.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                    ph.text = slide_def.get('title', '')
                elif ph.placeholder_format.type == PP_PLACEHOLDER.BODY:
                    tf = ph.text_frame
                    tf.clear()
                    for bullet in slide_def.get('bullets', []):
                        p = tf.add_paragraph()
                        p.text = bullet
                        p.level = 0
    except Exception as e:
        st.error(f"Error generating slides: {e}")
        st.stop()

    # --- Save and Download ---
    out_path = tempfile.mktemp(suffix=".pptx")
    prs.save(out_path)
    with open(out_path, 'rb') as f:
        data = f.read()
    st.success("Slides generated successfully!")
    st.download_button(
        "Download Generated Presentation",
        data=data,
        file_name="generated_presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
