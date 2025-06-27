import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
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
    # We access the package through the presentation's part.
    package = pres.part.package
    new_part = package.get_or_create_part(
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
        
        target_part = rel.target_part
        # If the target part of the relationship (e.g., an image file) isn't already
        # in the destination package...
        if not package.has_part(target_part.partname):
            # ...add it to the destination package.
            package.get_or_create_part(
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
                if shape.top < Pt(150):
                    found_slides.append(slide)
                    break
    return found_slides

def is_slide_of_type(slide, title_keyword):
    """Checks if a single slide's title contains a keyword."""
    for shape in slide.shapes:
        if shape.has_text_frame and title_keyword.lower() in shape.text.lower():
            if shape.top < Pt(150):
                return True
    return False

def populate_text_in_shape(shape, text):
    """Populates a shape with new text, clearing old content first and making it bold."""
    if not shape.has_text_frame:
        return
    
    tf = shape.text_frame
    for p in tf.paragraphs:
        p._p.getparent().remove(p._p)
    
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = text
    run.font.bold = True

# --- Streamlit App ---
st.set_page_config(page_title="AI Presentation Assembler", layout="wide")
st.title("🤖 AI Presentation Assembler")
st.write("This tool builds a new presentation using a Template Deck as the structure and a GTM Deck as the content source.")

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
                # Use the first uploaded template as the base for our new presentat
