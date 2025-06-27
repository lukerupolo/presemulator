import streamlit as st
import subprocess
import tempfile
import os
import json
from pptx import Presentation
import openai
from jinja2 import Template

# HTML template with reveal.js integration and CSS placeholders
HTML_TEMPLATE = """<!doctype html>
<html>
<head>
  <meta charset=\"utf-8\">  
  <title>Generated Presentation</title>
  <link rel=\"stylesheet\" href=\"https://unpkg.com/reveal.js/dist/reveal.css\">  
  <link rel=\"stylesheet\" href=\"https://unpkg.com/reveal.js/dist/theme/white.css\">  
  <style>
  /* Template-derived CSS */
  :root {
    --theme-font: {{ font_family }};
    --theme-font-size: {{ font_size }}px;
  }
  body { font-family: var(--theme-font); background: {{ background }}; }
  h2 { font-size: calc(var(--theme-font-size) * 1.5); color: {{ title_color }}; }
  ul { list-style-type: disc; margin-left: 20px; color: {{ text_color }}; }
  li { margin-bottom: 0.5em; }
  </style>
</head>
<body>
  <div class=\"reveal\"><div class=\"slides\">
  {%- for slide in slides %}
    <section>
      <h2>{{ slide.title }}</h2>
      <ul>
      {%- for bullet in slide.bullets %}
        <li>{{ bullet }}</li>
      {%- endfor %}
      </ul>
    </section>
  {%- endfor %}
  </div></div>
  <script src=\"https://unpkg.com/reveal.js/dist/reveal.js\"></script>
  <script>Reveal.initialize({ hash: true });</script>
</body>
</html>"""

# --- Streamlit UI ---
st.set_page_config(page_title="HTML Slide Generator", layout="wide")
st.title("AI Slide Generator (HTML/Reveal.js)")

# API Key
api_key = st.sidebar.text_input("OpenAI API Key", type="password")
if not api_key:
    st.sidebar.warning("Enter your OpenAI API key to proceed.")
    st.stop()
openai.api_key = api_key

# Upload PPTX/PDF template
uploaded = st.file_uploader("Upload a PPTX or PDF template", type=["pptx", "pdf"])
if not uploaded:
    st.info("Upload a PPTX or PDF to extract style.")
    st.stop()

# Save and (if needed) convert
suffix = os.path.splitext(uploaded.name)[1].lower()
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
    tmp.write(uploaded.getbuffer())
    path = tmp.name
if suffix == ".pdf":
    try:
        conv = path.replace('.pdf', '.pptx')
        subprocess.run(["unoconv", "-f", "pptx", path], check=True)
        path = conv
    except Exception as e:
        st.error(f"PDF → PPTX conversion failed: {e}")
        st.stop()

# Prompt + slide count
prompt = st.text_area("Prompt for slide content", height=150)
num_slides = st.number_input("Number of slides", min_value=1, max_value=20, value=5)

if st.button("Generate HTML Slides"):
    if not prompt.strip():
        st.error("Please enter a prompt.")
        st.stop()

    # 1. Generate outline via OpenAI
    with st.spinner("Generating slide outline..."):
        sys_msg = (
            "You are a slide outline generator. Return JSON with 'slides':"
            " [{ 'title': string, 'bullets': [string] }, ...]."
        )
        usr_msg = f"Prompt: {prompt}\nGenerate {num_slides} slides. Return only valid JSON."
        try:
            resp = openai.chat.completions.create(
                model="gpt-4",
                messages=[{'role':'system','content':sys_msg},
                          {'role':'user','content':usr_msg}],
                temperature=0.7
            )
            outline = json.loads(resp.choices[0].message.content)
        except Exception as e:
            st.error(f"OpenAI/JSON parse error: {e}")
            st.stop()

    # 2. Extract CSS variables from template
    try:
        prs = Presentation(path)
        ms = prs.slide_master
        # Background (color or image)
        bg_css = 'none'
        bg = ms.background.fill
        if bg and bg.type == 1:
            rgb = bg.fore_color.rgb
            bg_css = f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"  
        elif ms.background.fill.type == 6:  # picture fill
            pic = ms.background.fill.picture
            pic_path = pic._blip.filename
            bg_css = f"url('{pic_path}') center/cover"
        # Default font
        text_frame = prs.slides[0].placeholders[1].text_frame if prs.slides[0].placeholders else None
        font = text_frame.paragraphs[0].runs[0].font if text_frame else None
        font_family = font.name or 'sans-serif'
        font_size = int(font.size.pt) if font and font.size else 18
        title_color = f"#{font.color.rgb[0]:02X}{font.color.rgb[1]:02X}{font.color.rgb[2]:02X}" if font and font.color and font.color.rgb else '#000'
        text_color = title_color
    except Exception:
        bg_css = 'none'
        font_family, font_size, title_color, text_color = 'sans-serif', 18, '#000', '#000'

    # 3. Render HTML
    tpl = Template(HTML_TEMPLATE)
    html = tpl.render(
        slides=outline.get('slides', []),
        css='',  # inline handled above
        background=bg_css,
        font_family=font_family,
        font_size=font_size,
        title_color=title_color,
        text_color=text_color
    )

    # Display preview
    st.subheader("Preview:")
    st.components.v1.html(html, height=600, scrolling=True)

    # Download HTML
    st.download_button(
        "Download HTML file",
        data=html,
        file_name="presentation.html",
        mime="text/html"
    )
