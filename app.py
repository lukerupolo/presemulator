import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import copy
import uuid

# --- Core PowerPoint Functions ---

def clone_slide(pres, slide_to_clone):
    """
    Duplicates a slide from a source presentation and adds it to the end of
    the slides in the destination presentation `pres`.
    """
    src_part = slide_to_clone.part
    package = pres.part.package
    
    # Use get_or_create_part for robustness
    new_part = package.get_or_create_part(
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
        if not package.has_part(target_part.partname):
            package.get_or_create_part(
                target_part.partname, target_part.content_type, target_part.blob
            )
        new_part.relate_to(target_part, rel.reltype, rId=rel.rId)

    return pres.slides[-1]

def find_slide_by_title(prs, title_keyword):
    """Finds the first slide in a presentation containing a keyword in its title area."""
    if not title_keyword: return None
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and title_keyword.lower() in shape.text.lower():
                if shape.top < Pt(150):
                    return slide
    return None

def get_slide_content(slide):
    """Extracts title and body text from a slide."""
    title, body = "", ""
    title_shape = None
    # Find title first
    for shape in slide.shapes:
        if shape.has_text_frame and shape.top < Pt(150):
            title = shape.text
            title_shape = shape
            break
    
    # Find the largest text box that isn't the title for the body
    text_boxes = sorted(
        [s for s in slide.shapes if s.has_text_frame and s.text and s != title_shape],
        key=lambda s: len(s.text),
        reverse=True
    )
    if text_boxes:
        body = text_boxes[0].text
        
    return {"title": title, "body": body}


def populate_slide(slide, content):
    """Populates a slide's placeholders with new content, making it bold."""
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        
        # Simple heuristic to find placeholders in template
        if "title" in shape.name.lower() or (shape.top < Pt(150) and len(shape.text_frame.paragraphs) > 0):
             tf = shape.text_frame
             tf.clear()
             p = tf.paragraphs[0]
             run = p.add_run()
             run.text = content.get("title", "")
             run.font.bold = True

        elif "body" in shape.name.lower() or "content" in shape.name.lower() or "lorem ipsum" in shape.text.lower():
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = content.get("body", "")
            run.font.bold = True


# --- Streamlit App ---
st.set_page_config(page_title="Dynamic Presentation Assembler", layout="wide")
st.title("🤖 Dynamic Presentation Assembler")

# --- Initialize Session State ---
if 'structure' not in st.session_state:
    st.session_state.structure = []

# --- UI Sidebar ---
with st.sidebar:
    st.header("1. Upload Your Decks")
    template_file = st.file_uploader(
        "Upload Template Deck (.pptx)",
        type=["pptx"],
        help="The 'slide bank' with approved layouts."
    )
    gtm_file = st.file_uploader(
        "Upload GTM Global Deck (.pptx)",
        type=["pptx"],
        help="The source of core content."
    )
    st.markdown("---")
    st.header("2. Define Presentation Structure")
    
    # --- Structure Editor UI ---
    def add_step():
        st.session_state.structure.append(
            {"id": str(uuid.uuid4()), "keyword": "", "action": "Copy from GTM (as is)"}
        )

    st.button("Add New Step", on_click=add_step, use_container_width=True)

    for i, step in enumerate(st.session_state.structure):
        with st.container():
            st.markdown(f"**Step {i+1}**")
            cols = st.columns([3, 3, 1])
            step["keyword"] = cols[0].text_input("Slide Keyword (Title)", value=step["keyword"], key=f"keyword_{step['id']}")
            step["action"] = cols[1].selectbox(
                "Action",
                ["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"],
                index=["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"].index(step["action"]),
                key=f"action_{step['id']}"
            )
            def remove_step(index):
                st.session_state.structure.pop(index)
            cols[2].button("🗑️", on_click=remove_step, args=[i], key=f"del_{step['id']}")
    
    st.markdown("---")
    if st.button("Clear Structure", use_container_width=True):
        st.session_state.structure = []


# --- Main Content Area ---
if template_file and gtm_file:
    if st.session_state.structure:
        if st.button("🚀 Assemble Presentation", type="primary"):
            with st.spinner("Assembling your new presentation..."):
                try:
                    # --- Step 1: Load decks ---
                    st.write("Step 1/3: Loading and analyzing decks...")
                    template_prs = Presentation(io.BytesIO(template_file.getvalue()))
                    gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                    
                    new_prs = Presentation()
                    new_prs.slide_width = template_prs.slide_width
                    new_prs.slide_height = template_prs.slide_height

                    # --- Step 2: Iterate through the user-defined structure ---
                    st.write("Step 2/3: Building new presentation from your defined structure...")
                    for i, step in enumerate(st.session_state.structure):
                        keyword = step["keyword"]
                        action = step["action"]
                        st.write(f"  - Step {i+1}: Processing '{keyword}' with action '{action}'")

                        if action == "Copy from GTM (as is)":
                            content_slide = find_slide_by_title(gtm_prs, keyword)
                            if content_slide:
                                clone_slide(new_prs, content_slide)
                            else:
                                st.warning(f"  - Could not find a slide with keyword '{keyword}' in the GTM deck. Skipping.")
                        
                        elif action == "Merge: Template Layout + GTM Content":
                            layout_slide = find_slide_by_title(template_prs, keyword)
                            content_slide = find_slide_by_title(gtm_prs, keyword)

                            if layout_slide and content_slide:
                                content = get_slide_content(content_slide)
                                new_slide = clone_slide(new_prs, layout_slide)
                                populate_slide(new_slide, content)
                            else:
                                if not layout_slide: st.warning(f"  - Could not find layout for '{keyword}' in Template. Skipping.")
                                if not content_slide: st.warning(f"  - Could not find content for '{keyword}' in GTM Deck. Skipping.")
                    
                    st.success("Successfully built the new presentation structure.")

                    # --- Step 3: Finalize and provide download ---
                    st.write("Step 3/3: Finalizing and preparing download...")
                    output_buffer = io.BytesIO()
                    new_prs.save(output_buffer)
                    output_buffer.seek(0)

                    st.success("🎉 Your new regional presentation has been assembled!")
                    st.download_button(
                        label="Download Assembled PowerPoint",
                        data=output_buffer,
                        file_name="Dynamic_Assembled_Deck.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

                except Exception as e:
                    st.error(f"A critical error occurred: {e}")
                    st.exception(e)
    else:
        st.info("Define the presentation structure in the sidebar to begin.")
else:
    st.info("Please upload both a Template Deck and a GTM Global Deck to begin.")

