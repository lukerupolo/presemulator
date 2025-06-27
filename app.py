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
    src_part = slide_to_clone.part
    new_part = pres.part.package.add_part(
        src_part.partname, src_part.content_type, src_part.blob
    )
    pres.slides.add_slide(new_part)

    for rel in src_part.rels:
        if rel.is_external:
            new_part.rels.add_relationship(
                rel.reltype, rel.target_ref, rel.rId, is_external=True
            )
            continue
        
        target_part = rel.target_part
        if not pres.part.package.has_part(target_part.partname):
            pres.part.package.add_part(
                target_part.partname, target_part.content_type, target_part.blob
            )
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
                # --- Step 1: Load all decks ---
                st.write("Step 1/4: Loading and analyzing decks...")
                base_template_prs = Presentation(io.BytesIO(template_files[0].getvalue()))
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                
                # Create a new presentation to build into
                new_prs = Presentation()
                new_prs.slide_width = base_template_prs.slide_width
                new_prs.slide_height = base_template_prs.slide_height

                # --- Step 2: Get the list of "Objectives" slides from the GTM deck ---
                st.write("Step 2/4: Identifying content slides from GTM Deck...")
                gtm_objectives_slides = find_slides_by_title(gtm_prs, "objectives")
                gtm_objectives_queue = list(gtm_objectives_slides) # Create a queue to draw from
                st.success(f"Found {len(gtm_objectives_queue)} 'Objectives' slides in the GTM deck to use as content.")

                # --- Step 3: Iterate through the template and build the new deck ---
                st.write("Step 3/4: Building new presentation from template structure...")
                for template_slide in base_template_prs.slides:
                    # Check if the slide is an 'Objectives' slide
                    if is_slide_of_type(template_slide, "objectives"):
                        if gtm_objectives_queue:
                            # If it is, take the next available slide from the GTM deck and clone it
                            gtm_slide_to_clone = gtm_objectives_queue.pop(0)
                            clone_slide(new_prs, gtm_slide_to_clone)
                        else:
                            st.warning("Template has an 'Objectives' slide, but no more content slides were found in the GTM deck. Skipping.")
                    
                    # Check if the slide is an 'Activation' slide
                    elif is_slide_of_type(template_slide, "activation"):
                        # If it is, clone the template slide
                        newly_cloned_slide = clone_slide(new_prs, template_slide)
                        # And then populate it with placeholder text
                        for shape in newly_cloned_slide.shapes:
                            if shape.has_text_frame and "Lorem Ipsum" in shape.text:
                                populate_text_in_shape(shape, "Placeholder for regional activation details.\n- Tactic 1: [INSERT REGIONAL TACTIC]\n- Tactic 2: [INSERT REGIONAL TACTIC]\n- Budget: [INSERT REGIONAL BUDGET]")
                                break # Assume we only populate one placeholder
                    
                    # Otherwise, it's a standard slide to be copied as-is
                    else:
                        clone_slide(new_prs, template_slide)
                
                st.success("Successfully built the new presentation structure.")

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
