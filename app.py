import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor
import openai
import io
import os
import json

# --- Core Functions ---

def find_title_and_body_shapes(slide):
    """
    Finds the title and body shapes on a slide.
    Uses heuristics for slides without standard placeholders.
    """
    title_shape = None
    body_shape = None

    # First, try to find standard placeholders
    for shape in slide.placeholders:
        if 'TITLE' in shape.placeholder_format.type.name:
            title_shape = shape
            continue
        if 'BODY' in shape.placeholder_format.type.name or 'OBJECT' in shape.placeholder_format.type.name:
            body_shape = shape
            continue

    # If placeholders weren't found, use heuristics
    if title_shape is None and body_shape is None:
        # Assume largest text box is body, second largest is title
        # This is a fallback for non-standard slides
        text_boxes = sorted(
            [s for s in slide.shapes if s.has_text_frame and s.text.strip()],
            key=lambda s: len(s.text),
            reverse=True
        )
        if len(text_boxes) >= 2:
            body_shape = text_boxes[0]
            title_shape = text_boxes[1]
        elif len(text_boxes) == 1:
            body_shape = text_boxes[0]
            
    return title_shape, body_shape


def extract_structured_text_from_pptx(prs):
    """Extracts structured text (title and body) from each slide."""
    slide_data = []
    for i, slide in enumerate(prs.slides):
        title_shape, body_shape = find_title_and_body_shapes(slide)
        title_text = title_shape.text if title_shape else ""
        body_text = body_shape.text if body_shape else ""
        slide_data.append({"slide_number": i + 1, "title": title_text, "body": body_text})
    return slide_data

def get_ai_modified_content(api_key, original_slide_data, user_prompt):
    """Sends the structured slide data to OpenAI for modification."""
    try:
        client = openai.OpenAI(api_key=api_key)
    except Exception as e:
        raise ValueError(f"Failed to initialize OpenAI client. Check your API key. Error: {e}")

    system_prompt = """
    You are an expert presentation editor. You will be given a JSON object representing the slides in a presentation.
    Each slide object has a 'title' and a 'body'. Your task is to rewrite the title and body for each slide according to the user's instruction.
    You MUST return a JSON object with a single key 'modified_slides'. This key should contain an array of objects, one for each slide.
    Each object must have 'title' and 'body' keys with the rewritten text.
    The number of slide objects in your response must exactly match the number in the input. Maintain the original slide structure.
    If an original title or body is empty, you should still generate new content for it based on the instruction and the content of the other field.
    Do not add any extra commentary. Only return the JSON object.
    """

    full_user_prompt = f"""
    Instruction: "{user_prompt}"

    Original slide data:
    {json.dumps(original_slide_data, indent=2)}
    """
    
    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_user_prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.7,
        )
        response_content = response.choices[0].message.content
        modified_data = json.loads(response_content)
        
        if 'modified_slides' not in modified_data or not isinstance(modified_data['modified_slides'], list):
            raise ValueError("AI response did not contain 'modified_slides' list.")
            
        return modified_data['modified_slides']
    except Exception as e:
        st.error(f"An error occurred communicating with OpenAI: {e}")
        return None


def preserve_and_set_text(text_frame, new_text):
    """Replaces text in a text_frame while preserving the formatting of the first run."""
    if not text_frame.paragraphs:
        p = text_frame.add_paragraph()
    else:
        p = text_frame.paragraphs[0]
        
    font_name, font_size, font_bold, font_italic, font_color = None, Pt(18), None, None, None
    
    if p.runs:
        original_font = p.runs[0].font
        font_name = original_font.name
        font_size = original_font.size
        font_bold = original_font.bold
        font_italic = original_font.italic
        font_color = original_font.color
    
    # Clear all paragraphs to ensure clean slate
    for para in list(text_frame.paragraphs):
        p = para._p
        p.getparent().remove(p)

    # Add new paragraph with the new text and apply preserved formatting
    p = text_frame.add_paragraph()
    run = p.add_run()
    run.text = new_text
    
    font = run.font
    if font_name:
        font.name = font_name
    if font_size:
        font.size = font_size
    if font_bold is not None:
        font.bold = font_bold
    if font_italic is not None:
        font.italic = font_italic
    if font_color:
        if font_color.type == MSO_THEME_COLOR:
            font.color.theme_color = font_color.theme_color
            font.color.brightness = font_color.brightness
        elif hasattr(font_color, 'rgb'):
            font.color.rgb = font_color.rgb

def update_presentation_with_new_text(prs, modified_slides):
    """Updates presentation with AI-modified text, using robust shape finding."""
    if len(prs.slides) != len(modified_slides):
        st.warning("Mismatch between original slide count and AI response count.")
        return

    for i, slide in enumerate(prs.slides):
        slide_mods = modified_slides[i]
        title_shape, body_shape = find_title_and_body_shapes(slide)

        if title_shape:
            preserve_and_set_text(title_shape.text_frame, slide_mods.get('title', ''))
        
        if body_shape:
            preserve_and_set_text(body_shape.text_frame, slide_mods.get('body', ''))

# --- UI Functions ---
def display_summary(original_data, modified_data):
    """Displays a side-by-side comparison of changes."""
    st.subheader("📝 Summary of Modifications")
    st.write("Review the changes made by the AI below.")
    
    for original, modified in zip(original_data, modified_data):
        st.markdown(f"---")
        slide_num = original['slide_number']
        st.markdown(f"### Slide {slide_num}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Before**")
            st.text_area("Title", value=original['title'], height=70, disabled=True, key=f"orig_title_{slide_num}")
            st.text_area("Body", value=original['body'], height=200, disabled=True, key=f"orig_body_{slide_num}")
        with col2:
            st.markdown("**After**")
            st.text_area("Title", value=modified.get('title', ''), height=70, disabled=True, key=f"mod_title_{slide_num}")
            st.text_area("Body", value=modified.get('body', ''), height=200, disabled=True, key=f"mod_body_{slide_num}")

# --- Streamlit App ---
st.set_page_config(page_title="AI PowerPoint Editor", layout="wide")
st.title("🤖 AI-Powered PowerPoint Content Editor")
st.write("This application intelligently rewrites your presentation's content while preserving its original formatting.")

with st.sidebar:
    st.header("Controls")
    api_key = st.text_input("Enter your OpenAI API Key", type="password")
    st.markdown("---")
    uploaded_file = st.file_uploader("1. Upload a PowerPoint (.pptx)", type=["pptx"])
    st.markdown("---")
    user_prompt = st.text_area(
        "2. Enter your editing instruction",
        height=150,
        placeholder="e.g., 'Rewrite the content to reflect an Australian market perspective.' or 'Summarize the key points into bullet points.'"
    )

if uploaded_file is not None:
    if st.button("✨ Process Presentation with AI", type="primary"):
        if not api_key: st.error("Please enter your OpenAI API key.")
        elif not user_prompt: st.error("Please enter an instruction for the AI.")
        else:
            with st.spinner("Processing... This may take a moment."):
                try:
                    file_content = uploaded_file.getvalue()
                    prs = Presentation(io.BytesIO(file_content))
                    
                    st.write("Step 1/4: Reading slide titles and content...")
                    original_data = extract_structured_text_from_pptx(prs)
                    
                    st.write("Step 2/4: Asking AI to rewrite content...")
                    modified_data = get_ai_modified_content(api_key, original_data, user_prompt)

                    if modified_data:
                        st.write("Step 3/4: Updating slides while preserving formatting...")
                        prs_to_edit = Presentation(io.BytesIO(file_content))
                        update_presentation_with_new_text(prs_to_edit, modified_data)

                        st.write("Step 4/4: Preparing your download...")
                        output_buffer = io.BytesIO()
                        prs_to_edit.save(output_buffer)
                        output_buffer.seek(0)
                        
                        base, ext = os.path.splitext(uploaded_file.name)
                        new_filename = f"{base}_ai_modified.pptx"

                        st.success("🎉 Your presentation has been successfully modified!")
                        display_summary(original_data, modified_data)
                        st.download_button(
                            label="Download Modified PowerPoint",
                            data=output_buffer,
                            file_name=new_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
                except Exception as e:
                    st.error(f"A critical error occurred: {e}")
                    st.exception(e) # Provides a more detailed traceback for debugging
else:
    st.info("Upload a PowerPoint, provide an API key and instructions in the sidebar to begin.")
