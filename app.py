import streamlit as st
import subprocess
import tempfile
import os
import json
from pptx import Presentation
import openai

# --- App Config ---
st.set_page_config(page_title="AI Slide Generator", layout="wide")
st.title("AI Slide Generator with Template Style")

# --- Sidebar: OpenAI API Key ---
api_key = st.sidebar.text_input("OpenAI API Key", type="password")
if not api_key:
    st.sidebar.warning("Please enter your OpenAI API Key to proceed.")
    st.stop()
openai.api_key = api_key

# --- File Upload ---
uploaded = st.file_uploader("Upload a PPTX or PDF template", type=["pptx", "pdf"])
if not uploaded:
    st.info("Awaiting PPTX or PDF template upload...")
    st.stop()

# Save upload to temp file
suffix = os.path.splitext(uploaded.name)[1].lower()
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
    tmp.write(uploaded.getbuffer())
    template_path = tmp.name

# Convert PDF to PPTX if needed
if suffix == ".pdf":
    try:
        template_path = template_path.replace('.pdf', '.pptx')
        subprocess.run(["unoconv", "-f", "pptx", tmp.name], check=True)
    except Exception as e:
        st.error(f"PDF → PPTX conversion failed: {e}")
        st.stop()

# --- Slide Generation Parameters ---
prompt = st.text_area("Enter a prompt for slide content generation", height=150)
num_slides = st.number_input("Number of slides to generate", min_value=1, max_value=20, value=5)

if st.button("Generate Slides"):
    if not prompt.strip():
        st.error("Please provide a prompt for content generation.")
        st.stop()

    with st.spinner("Generating slide outline with OpenAI..."):
        system_msg = (
            "You are an assistant that creates PowerPoint slide outlines. "
            "Given a user prompt, generate a JSON object with a 'slides' array. "
            "Each slide entry should have 'title' (string) and 'bullets' (list of strings)."
        )
        user_msg = (f"Prompt: {prompt}\nGenerate {num_slides} slides." +
                    "\nReturn only valid JSON.")
        try:
            resp = openai.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": system_msg},
                    {"role": "user", "content": user_msg}
                ],
                temperature=0.7
            )
            outline_json = resp.choices[0].message.content
            outline = json.loads(outline_json)
        except Exception as e:
            st.error(f"OpenAI or JSON parse error: {e}")
            st.stop()

    # --- Build Slides ---
    try:
        prs = Presentation(template_path)
        layout = prs.slide_layouts[1]  # Title & Content
        for slide_def in outline.get("slides", []):
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = slide_def.get("title", "")
            body = slide.placeholders[1].text_frame
            body.clear()
            for b in slide_def.get("bullets", []):
                p = body.add_paragraph()
                p.text = b
                p.level = 0
    except Exception as e:
        st.error(f"Error generating slides: {e}")
        st.stop()

    # --- Download Result ---
    out_path = tempfile.mktemp(suffix=".pptx")
    prs.save(out_path)
    with open(out_path, 'rb') as f:
        data = f.read()
    st.success("Slides generated successfully!")
    st.download_button(
        "Download New Presentation",
        data=data,
        file_name="generated_presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
