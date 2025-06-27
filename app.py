import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import copy
import uuid
import openai
import json

# --- Core PowerPoint Functions ---

def deep_copy_slide(dest_slide, src_slide):
    """
    Performs a deep copy of shapes and content from a source slide to a destination slide.
    This is a stable alternative to cloning parts between presentations.
    """
    # Clear all shapes from the destination slide first
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)

    # Copy shapes from source to destination
    for shape in src_slide.shapes:
        new_shape_el = copy.deepcopy(shape.element)
        dest_slide.shapes._spTree.insert_element_before(new_shape_el, 'p:extLst')

def find_slide_by_ai(api_key, prs, slide_type_prompt):
    """Uses OpenAI to intelligently find the best matching slide in a presentation."""
    if not slide_type_prompt: return None
    client = openai.OpenAI(api_key=api_key)
    
    slides_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(prs.slides)]

    system_prompt = f"""
    You are an expert presentation analyst. Given a JSON list of slide contents and a description ('{slide_type_prompt}'), identify the index of the single best-matching slide.
    Analyze text for purpose (e.g., "Timeline" has dates; "Objectives" has goal language).
    Return a JSON object with a single key 'best_match_index'. If no good match is found, return -1.
    """
    full_user_prompt = f"Find the best slide for '{slide_type_prompt}' in: {json.dumps(slides_content, indent=2)}"

    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo",
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": full_user_prompt}],
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        best_index = result.get("best_match_index", -1)
        if best_index != -1 and best_index < len(prs.slides):
            return prs.slides[best_index]
        return None
    except Exception as e:
        st.error(f"AI slide analysis failed for '{slide_type_prompt}': {e}")
        return None

def get_slide_content(slide):
    """Extracts title and body text from a slide."""
    text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and s.text.strip()], key=lambda s: s.top)
    title = text_boxes[0].text if text_boxes else ""
    body = "\n".join(s.text for s in text_boxes[1:]) if len(text_boxes) > 1 else ""
    return {"title": title, "body": body}

def populate_slide(slide, content):
    """Populates a slide's placeholders with new content, making it bold."""
    text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and "lorem ipsum" in s.text.lower()], key=lambda s: s.top)
    if not text_boxes:
        text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and len(s.text_frame.paragraphs) > 0], key=lambda s: s.top)
        if len(text_boxes) > 1:
            text_boxes = text_boxes[1:]
        elif not text_boxes:
            st.warning("Could not find a placeholder to populate on a merged slide.")
            return
    
    # Heuristic: combine title and body into the first placeholder found
    target_shape = text_boxes[0]
    tf = target_shape.text_frame
    tf.clear()
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = content.get('title', '') + '\n\n' + content.get('body', '')
    run.font.bold = True

# --- Streamlit App ---
st.set_page_config(page_title="Dynamic AI Presentation Assembler", layout="wide")
st.title("🤖 Dynamic AI Presentation Assembler")

with st.sidebar:
    st.header("1. API Key")
    api_key = st.text_input("OpenAI API Key", type="password")
    st.markdown("---")
    st.header("2. Upload Decks")
    template_files = st.file_uploader("Upload Template Deck(s)", type=["pptx"], accept_multiple_files=True)
    gtm_file = st.file_uploader("Upload GTM Global Deck", type=["pptx"])
    st.markdown("---")
    st.header("3. Define Presentation Structure")
    
    if 'structure' not in st.session_state:
        st.session_state.structure = []
    
    if st.button("Add New Step", use_container_width=True):
        st.session_state.structure.append({"id": str(uuid.uuid4()), "keyword": "", "action": "Copy from GTM (as is)"})

    for i, step in enumerate(st.session_state.structure):
        with st.container(border=True):
            cols = st.columns([3, 3, 1])
            step["keyword"] = cols[0].text_input("Slide Type (e.g., 'Objectives')", step["keyword"], key=f"keyword_{step['id']}")
            step["action"] = cols[1].selectbox("Action", ["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"], index=["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"].index(step["action"]), key=f"action_{step['id']}")
            if cols[2].button("🗑️", key=f"del_{step['id']}"):
                st.session_state.structure.pop(i)
                st.rerun()

    if st.button("Clear Structure", use_container_width=True):
        st.session_state.structure = []
        st.rerun()

# --- Main App Logic ---
if template_files and gtm_file and api_key and st.session_state.structure:
    if st.button("🚀 Assemble Presentation", type="primary"):
        with st.spinner("Assembling your new presentation... This may take a moment."):
            try:
                st.write("Step 1/3: Loading decks...")
                # Load all decks into memory
                template_prs_list = [Presentation(io.BytesIO(f.getvalue())) for f in template_files]
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                
                # CRITICAL FIX: Use the first template as the base for the new presentation.
                # This ensures the final deck has the correct master slides and themes.
                new_prs = Presentation(io.BytesIO(template_files[0].getvalue()))

                # Create a build plan based on the user's defined structure
                build_plan = []
                for step in st.session_state.structure:
                    keyword = step["keyword"]
                    action = step["action"]
                    
                    if action == "Copy from GTM (as is)":
                        slide = find_slide_by_ai(api_key, gtm_prs, keyword)
                        if slide:
                            build_plan.append({"action": "clone", "source_slide": slide})
                        else:
                            st.warning(f"AI could not find '{keyword}' in GTM Deck. Skipping.")
                    
                    elif action == "Merge: Template Layout + GTM Content":
                        layout_slide = find_slide_by_ai(api_key, template_prs_list[0], keyword)
                        content_slide = find_slide_by_ai(api_key, gtm_prs, keyword)
                        if layout_slide and content_slide:
                            content = get_slide_content(content_slide)
                            build_plan.append({"action": "merge", "layout_slide": layout_slide, "content": content})
                        else:
                            if not layout_slide: st.warning(f"AI could not find layout for '{keyword}' in Template. Skipping.")
                            if not content_slide: st.warning(f"AI could not find content for '{keyword}' in GTM Deck. Skipping.")

                # Clear all slides from the base presentation to create a clean canvas
                for i in range(len(new_prs.slides) - 1, -1, -1):
                    rId = new_prs.slides._sldIdLst[i].rId
                    new_prs.part.drop_rel(rId)
                    del new_prs.slides._sldIdLst[i]

                st.write("Step 2/3: Building new presentation from plan...")
                # Execute the build plan
                for item in build_plan:
                    if item["action"] == "clone":
                        deep_copy_slide(new_prs.slides.add_slide(new_prs.slide_layouts[6]), item["source_slide"])
                    elif item["action"] == "merge":
                        new_slide = new_prs.slides.add_slide(item["layout_slide"].slide_layout)
                        populate_slide(new_slide, item["content"])

                st.success("Successfully built the new presentation structure.")
                st.write("Step 3/3: Finalizing and preparing download...")
                output_buffer = io.BytesIO()
                new_prs.save(output_buffer)
                output_buffer.seek(0)

                st.success("🎉 Your new regional presentation has been assembled!")
                st.download_button("Download Assembled PowerPoint", data=output_buffer, file_name="Dynamic_AI_Assembled_Deck.pptx")

            except Exception as e:
                st.error(f"A critical error occurred: {e}")
                st.exception(e)
else:
    st.info("Please provide an API Key, upload both a GTM Deck and at least one Template Deck, and define the structure in the sidebar to begin.")
