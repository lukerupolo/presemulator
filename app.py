import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import os
import copy

# --- Core PowerPoint Functions ---

def clone_slide(pres, slide_to_clone):
    """
    Duplicates a slide from a source presentation and adds it to the end of
    the slides in the destination presentation `pres`. This is a robust method
    that correctly handles all slide parts and relationships (like images).
    """
    # 1. Get the source slide's part (the XML representation of the slide).
    src_part = slide_to_clone.part

    # 2. Add a new slide part to the destination presentation's package.
    # This copies the raw XML and binary content of the slide.
    # The package handles assigning a unique name if a conflict exists.
    new_part = pres.part.package.add_part(
        src_part.partname, src_part.content_type, src_part.blob
    )

    # 3. Add the new slide part to the presentation's main slide list.
    # This makes the slide "visible" in the slide sequence.
    pres.slides.add_slide(new_part)

    # 4. Copy relationships from the source slide to the new slide.
    # This is CRITICAL for images, charts, etc.
    for rel in src_part.rels:
        # If the relationship is external (e.g., a hyperlink), copy it as is.
        if rel.is_external:
            new_part.rels.add_relationship(
                rel.reltype, rel.target_ref, rel.rId, is_external=True
            )
            continue

        # If the target of the relationship (e.g., an image file) isn't already
        # in the destination package...
        target_part = rel.target_part
        if not pres.part.package.has_part(target_part.partname):
            # ...add it to the destination package.
            pres.part.package.add_part(
                target_part.partname, target_part.content_type, target_part.blob
            )

        # Create the relationship from the new slide to the now-guaranteed-to-exist target part.
        new_part.relate_to(target_part, rel.reltype, rId=rel.rId)

    return pres.slides[-1]

def find_slides_by_title(prs, title_keyword):
    """Finds all slides in a presentation containing a keyword in their title area."""
    found_slides = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and title_keyword.lower() in shape.text.lower():
                # Heuristic: A title shape is usually near the top of the slide.
                if shape.top < Pt(150):
                    found_slides.append(slide)
                    break
    return found_slides

def find_slide_in_templates(template_prs_list, title_keyword):
    """Searches through a list of template presentations to find the first matching slide."""
    for prs in template_prs_list:
        found_slides = find_slides_by_title(prs, title_keyword)
        if found_slides:
            return found_slides[0]  # Return the first one found
    return None

def populate_text_in_shape(shape, text):
    """Populates a shape with new text, clearing old content first."""
    if not shape.has_text_frame:
        return
    
    tf = shape.text_frame
    # Clear all existing paragraphs by removing their underlying XML elements.
    for p in tf.paragraphs:
        p._p.getparent().remove(p._p)
    
    # Add a new paragraph and run for the new text.
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = text

# --- Streamlit App ---
st.set_page_config(page_title="AI Presentation Assembler", layout="wide")
st.title("🤖 AI Presentation Assembler")
st.write("This tool builds a new regional presentation by combining slides from a global GTM deck and a template slide bank.")

st.header("1. Upload Your Decks")
col1, col2 = st.columns(2)

with col1:
    template_files = st.file_uploader(
        "Upload Template Deck(s) (.pptx)",
        type=["pptx"],
        accept_multiple_files=True,
        help="The 'slide bank' with approved layouts (e.g., Activation slide)."
    )
with col2:
    gtm_file = st.file_uploader(
        "Upload GTM Global Deck (.pptx)",
        type=["pptx"],
        help="The source of core content to be copied directly (e.g., Objectives slides)."
    )

if template_files and gtm_file:
    if st.button("🚀 Assemble Regional Deck", type="primary"):
        with st.spinner("Assembling your new presentation..."):
            try:
                # --- Step 1: Load decks and create a stable base presentation ---
                st.write("Step 1/4: Loading decks and preparing a clean base...")
                template_prs_list = [Presentation(io.BytesIO(f.getvalue())) for f in template_files]
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))

                # CRITICAL FIX: Use the first template as the base for the new presentation.
                new_prs = Presentation(io.BytesIO(template_files[0].getvalue()))

                # Delete all slides from the base to create a clean, styled canvas.
                for i in range(len(new_prs.slides) - 1, -1, -1):
                    rId = new_prs.slides._sldIdLst[i].rId
                    new_prs.part.drop_rel(rId)
                    del new_prs.slides._sldIdLst[i]

                # --- Step 2: Find and copy "Objectives" slides from GTM deck ---
                st.write("Step 2/4: Finding and copying 'Franchise Objectives' slides...")
                objective_slides_from_gtm = find_slides_by_title(gtm_prs, "objectives")

                if not objective_slides_from_gtm:
                    st.warning("Could not find any slides with 'Objectives' in the GTM deck title.")
                else:
                    for slide in objective_slides_from_gtm:
                        clone_slide(new_prs, slide)
                    st.success(f"Copied {len(objective_slides_from_gtm)} 'Objectives' slide(s) from the GTM deck.")

                # --- Step 3: Find "Activation" slide in templates, copy, then populate ---
                st.write("Step 3/4: Finding 'Activation' slide in templates and populating...")
                activation_slide_from_template = find_slide_in_templates(template_prs_list, "activation")

                if not activation_slide_from_template:
                    st.warning("Could not find any slides with 'Activation' in any of the Template decks.")
                else:
                    copied_activation_slide = clone_slide(new_prs, activation_slide_from_template)
                    
                    for shape in copied_activation_slide.shapes:
                        if shape.has_text_frame and "Lorem Ipsum" in shape.text:
                            populate_text_in_shape(shape, "Placeholder for regional activation details.\n- Tactic 1: [INSERT REGIONAL TACTIC]\n- Tactic 2: [INSERT REGIONAL TACTIC]\n- Budget: [INSERT REGIONAL BUDGET]")
                    st.success("Added and populated 1 'Activation' slide from the template bank.")
                
                # --- Step 4: Finalize and provide download ---
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

