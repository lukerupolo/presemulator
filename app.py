import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import os
import openai
import copy
from lxml.etree import ElementBase

# --- Core PowerPoint Functions ---

def clone_slide(pres, slide):
    """
    Duplicates a slide from a source presentation and adds it to the end of
    the slides in the destination presentation `pres`.
    Ensures all parts (like images) are copied correctly.
    """
    # 1. Get the source slide's part
    src_slide_part = slide.part

    # 2. Add a new slide part to the destination presentation
    # The partname is derived from the source. python-pptx handles uniqueness.
    # We copy the 'blob' (binary content) of the source slide part.
    # This is a low-level operation on the underlying package.
    new_part = pres.part.package.add_part(
        src_slide_part.partname, src_slide_part.content_type, src_slide_part.blob
    )

    # 3. Add a relationship from the presentation's main part to the new slide part.
    # This makes the new slide "appear" in the slide list.
    pres.slides.add_slide(new_part)

    # 4. Copy relationships from the source slide part to the new slide part
    for rel in src_slide_part.rels:
        # If the relationship is external, copy it as is
        if rel.is_external:
            new_part.rels.add_relationship(
                rel.reltype, rel.target_ref, rel.rId, is_external=True
            )
            continue

        # If the relationship target part (e.g., an image) isn't already in the destination package...
        target_part = rel.target_part
        if not pres.part.package.has_part(target_part.partname):
            # ...copy the target part to the destination package
            pres.part.package.add_part(
                target_part.partname, target_part.content_type, target_part.blob
            )

        # Add the relationship from the new slide to the new target part
        new_part.rels.add_relationship(rel.reltype, target_part, rel.rId)
    
    # The new slide is the last one in the presentation
    return pres.slides[-1]


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
            return found_slides[0]
    return None

def populate_text_in_shape(shape, text):
    """Populates a shape with new text, clearing old content."""
    if not shape.has_text_frame:
        return
        
    tf = shape.text_frame
    # Clear all existing paragraphs by removing their XML elements, which is more robust.
    for p in tf.paragraphs:
        p._p.getparent().remove(p._p)
    
    # Add a new paragraph and run for the new text
    p = tf.add_paragraph()
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
                template_prs_list = [Presentation(io.BytesIO(f.getvalue())) for f in template_files]
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))

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
                activation_slide_from_template = find_slide_in_templates(template_prs_list, "activation")

                if not activation_slide_from_template:
                    st.warning("Could not find any slides with 'Activation' in any of the Template decks.")
                else:
                    copied_activation_slide = clone_slide(new_prs, activation_slide_from_template)
                    
                    for shape in copied_activation_slide.shapes:
                        # Using a simple heuristic to find the main body text placeholder
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
