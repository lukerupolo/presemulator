import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io
import os
import openai
import copy
from lxml.etree import ElementBase

# --- Core PowerPoint Functions ---

def clone_slide(pres, slide_to_clone):
    """
    Duplicates a slide from a source presentation into a target presentation.
    This is a complex but necessary function to ensure a high-fidelity copy.
    """
    # 1. Add a blank slide to the target presentation
    blank_slide_layout = pres.slide_layouts[6] # Index 6 is typically a blank layout
    new_slide = pres.slides.add_slide(blank_slide_layout)

    # 2. Copy background
    new_slide.background.fill.solid()
    new_slide.background.fill.fore_color.rgb = slide_to_clone.background.fill.fore_color.rgb
    
    # 3. Copy shapes from the original slide to the new slide
    for shape in slide_to_clone.shapes:
        if shape.is_placeholder:
            # Handle placeholders by creating a new one
            ph = shape.placeholder_format
            new_ph = new_slide.placeholders.add_placeholder(ph.idx, ph.type, shape.left, shape.top, shape.width, shape.height)
            # Copy text and formatting
            if shape.has_text_frame:
                new_ph.text_frame.text = shape.text_frame.text
        else:
            # For non-placeholder shapes, we copy them by creating a new element
            new_el = copy.deepcopy(shape.element)
            new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    return new_slide

def find_slide_by_title(prs, title):
    """Finds the first slide in a presentation that contains a specific title."""
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame and title.lower() in shape.text.lower():
                return slide
    return None

def populate_title_slide(slide, project_name, subtitle):
    """Populates the placeholders on a title slide with new text."""
    title_placeholder = None
    subtitle_placeholder = None

    # Find title and subtitle placeholders
    for shape in slide.placeholders:
        if shape.placeholder_format.type.name in ('TITLE', 'CENTER_TITLE'):
            title_placeholder = shape
        elif shape.placeholder_format.type.name == 'SUBTITLE':
            subtitle_placeholder = shape
            
    # If standard placeholders aren't found, use heuristics
    if not title_placeholder or not subtitle_placeholder:
         text_boxes = sorted([s for s in slide.shapes if s.has_text_frame], key=lambda s: s.top)
         if len(text_boxes) >= 2:
             title_placeholder = text_boxes[0]
             subtitle_placeholder = text_boxes[1]

    if title_placeholder:
        title_placeholder.text = project_name
    if subtitle_placeholder:
        subtitle_placeholder.text = subtitle

# --- Streamlit App ---

st.set_page_config(page_title="AI Presentation Generator", layout="wide")
st.title("🤖 AI Presentation Generator from Template")
st.write("This tool uses a template presentation as a 'slide bank' to build new decks, ensuring perfect style emulation.")

# --- UI for File Uploads ---
st.header("1. Upload Your Files")
col1, col2 = st.columns(2)

with col1:
    template_file = st.file_uploader("Upload Template Deck (.pptx)", type=["pptx"], help="This presentation will be used as a 'slide bank' for styles and layouts.")

with col2:
    # In the future, this could be a content deck or just a prompt
    content_prompt = st.text_input("Enter the main topic for the presentation", value="Quarterly Marketing Plan", help="For this demo, this will be used to generate a title.")


# --- Main Logic ---
if template_file and content_prompt:
    if st.button("🚀 Generate Presentation", type="primary"):
        with st.spinner("Building your presentation..."):
            try:
                # --- Step 1: Load the template and find the title slide ---
                st.write("Step 1/3: Analyzing template and finding 'Title Page'...")
                template_bytes = template_file.getvalue()
                template_prs = Presentation(io.BytesIO(template_bytes))
                
                title_slide_from_template = find_slide_by_title(template_prs, "TITLE PAGE")
                
                if not title_slide_from_template:
                    st.error("Could not find a slide with 'TITLE PAGE' in your template. Please ensure one exists.")
                    st.stop()
                
                st.success("Found 'TITLE PAGE' in the template.")

                # --- Step 2: Create a new presentation and copy the slide ---
                st.write("Step 2/3: Creating new deck and copying the title slide...")
                new_prs = Presentation()
                
                # We need to make sure the slide size of the new presentation matches the template
                new_prs.slide_width = template_prs.slide_width
                new_prs.slide_height = template_prs.slide_height

                copied_slide = clone_slide(new_prs, title_slide_from_template)
                
                # --- Step 3: Populate the new slide with AI-generated (or placeholder) content ---
                st.write("Step 3/3: Populating the new slide with content...")
                
                # For this demo, we'll use simple placeholder text as requested.
                # In a real scenario, this would come from another AI call or content file.
                project_name = content_prompt
                subtitle_text = "AI-Generated Content | Regional Strategy"

                populate_title_slide(copied_slide, project_name, subtitle_text)
                
                # --- Final Step: Save and provide download ---
                output_buffer = io.BytesIO()
                new_prs.save(output_buffer)
                output_buffer.seek(0)

                st.success("🎉 Your new presentation has been generated!")

                new_filename = f"{project_name.replace(' ', '_')}_presentation.pptx"
                st.download_button(
                    label="Download Generated PowerPoint",
                    data=output_buffer,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

            except Exception as e:
                st.error(f"A critical error occurred: {e}")
                st.exception(e)

else:
    st.info("Please upload a template deck and provide a topic to begin.")
