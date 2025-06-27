import streamlit as st
from pptx import Presentation
import io
import os
import openai
import copy
from lxml.etree import ElementBase

# --- Core PowerPoint Functions ---

def clone_slide(pres_to, slide_from):
    """
    Duplicates a slide from a source presentation into a target presentation.
    This is the most critical function for a high-fidelity copy.
    """
    # Using a blank layout is a common starting point
    blank_slide_layout = pres_to.slide_layouts[6] 
    new_slide = pres_to.slides.add_slide(blank_slide_layout)

    sl = slide_from.element
    new_sl = copy.deepcopy(sl)

    rId_map = {}

    for r in slide_from.part.rels:
        rel = slide_from.part.rels[r]
        if "image" in rel.target_ref:
            image_part = rel.target_part
            new_part = pres_to.part.package.get_part(image_part.partname)
            if new_part is None:
                new_part = pres_to.part.package.add_part(image_part.partname, image_part.content_type, image_part.blob)
            
            new_rId = new_slide.part.relate_to(new_part, rel.reltype)
            rId_map[rel.rId] = new_rId

    new_slide.element.getparent().replace(new_slide.element, new_sl)
    
    for old_rId, new_rId in rId_map.items():
        for elem in new_slide.element.xpath(f'.//*[@r:id="{old_rId}"]', namespaces={'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}):
            elem.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', new_rId)
            
    return new_slide


def find_slides_by_title(prs, title_keyword):
    """Finds all slides in a single presentation that contain a specific keyword in their title."""
    found_slides = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and title_keyword.lower() in shape.text.lower():
                if shape.top < Pt(150): 
                    found_slides.append(slide)
                    break 
    return found_slides

def find_slide_in_templates(template_prs_list, title_keyword):
    """Searches through a list of template presentations to find the first slide with a matching title."""
    for prs in template_prs_list:
        found_slides = find_slides_by_title(prs, title_keyword)
        if found_slides:
            return found_slides[0] # Return the first slide found
    return None

def populate_text_in_shape(shape, text):
    """Populates a shape with new text, preserving the first run's formatting."""
    if not shape.has_text_frame:
        return
        
    tf = shape.text_frame
    for para in tf.paragraphs:
        for run in para.runs:
            run.text = ''
    
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text

# --- Streamlit App ---

st.set_page_config(page_title="AI Presentation Assembler", layout="wide")
st.title("🤖 AI Presentation Assembler")
st.write("This tool builds a new regional presentation by combining slides from a global GTM deck and a template slide bank.")

# --- UI for File Uploads ---
st.header("1. Upload Your Decks")
col1, col2 = st.columns(2)

with col1:
    # UPDATED: Now accepts multiple files for the template bank
    template_files = st.file_uploader("Upload Template Deck(s) (.pptx)", type=["pptx"], accept_multiple_files=True, help="The 'slide bank' with approved layouts (e.g., Activation slide).")

with col2:
    gtm_file = st.file_uploader("Upload GTM Global Deck (.pptx)", type=["pptx"], help="The source of core content to be copied directly (e.g., Objectives slides).")


# --- Main Logic ---
if template_files and gtm_file:
    if st.button("🚀 Assemble Regional Deck", type="primary"):
        with st.spinner("Assembling your new presentation..."):
            try:
                # --- Step 1: Load all presentations ---
                st.write("Step 1/4: Loading and analyzing decks...")
                # Load all uploaded template files into a list
                template_prs_list = [Presentation(io.BytesIO(f.getvalue())) for f in template_files]
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))

                # Create the new presentation, matching the first template's size
                new_prs = Presentation()
                new_prs.slide_width = template_prs_list[0].slide_width
                new_prs.slide_height = template_prs_list[0].slide_height

                # --- Step 2: Find and copy "Objectives" slides from GTM deck ---
                st.write("Step 2/4: Finding and copying 'Franchise Objectives' slides...")
                objective_slides_from_gtm = find_slides_by_title(gtm_prs, "objectives")

                if not objective_slides_from_gtm:
                    st.warning("Could not find any slides with 'Objectives' in the GTM deck title.")
                else:
                    for slide in objective_slides_from_gtm:
                        clone_slide(new_prs, slide)
                    st.success(f"Copied {len(objective_slides_from_gtm)} 'Objectives' slide(s) from the GTM deck.")

                # --- Step 3: Find "Activation" slide in ALL templates, copy, then populate ---
                st.write("Step 3/4: Finding 'Activation' slide in templates and populating...")
                # Search across all provided template decks
                activation_slide_from_template = find_slide_in_templates(template_prs_list, "activation")

                if not activation_slide_from_template:
                    st.warning("Could not find any slides with 'Activation' in any of the Template decks.")
                else:
                    copied_activation_slide = clone_slide(new_prs, activation_slide_from_template)
                    
                    for shape in copied_activation_slide.shapes:
                        if shape.has_text_frame and "Lorem Ipsum" in shape.text:
                            populate_text_in_shape(shape, "Placeholder for regional activation details.\n- Tactic 1: [INSERT REGIONAL TACTIC]\n- Tactic 2: [INSERT REGIONAL TACTIC]\n- Budget: [INSERT REGIONAL BUDGET]")
                    
                    st.success("Added and populated 1 'Activation' slide from the template bank.")
                
                # --- Final Step: Save and provide download ---
                st.write("Step 4/4: Finalizing and preparing download...")
                output_buffer = io.BytesIO()
                new_prs.save(output_buffer)
                output_buffer.seek(0)

                st.success("🎉 Your new regional presentation has been assembled!")

                new_filename = "Regional_Deck_Assembled.pptx"
                st.download_button(
                    label="Download Assembled PowerPoint",
                    data=output_buffer,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

            except Exception as e:
                st.error(f"A critical error occurred: {e}")
                st.exception(e)

else:
    st.info("Please upload both a GTM Global Deck and at least one Template Deck to begin.")
