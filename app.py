import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import copy
import uuid
import openai
import json
import requests
from lxml.etree import QName

# --- Core PowerPoint Functions ---

def deep_copy_slide(dest_pres, dest_slide, src_slide):
    """
    Performs a stable, deep copy of all shapes and content from a source slide
    to a destination slide. This is the most robust method for "Copy from GTM".
    It now handles linked images by downloading and embedding them.
    """
    # Clear all shapes from the destination slide first to prepare it.
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)

    # Iterate through shapes in the source slide and copy them to the destination.
    for shape in src_slide.shapes:
        # If it's a linked picture, we need to download and embed it.
        if shape.shape_type == 13 and hasattr(shape, 'image'): # 13 is the shape type for Picture
            # This is a heuristic to identify linked pictures. A more robust solution might
            # need to parse XML to find the external relationship ID.
            try:
                # Get the URL of the linked image
                image_url = shape.image.ext_uri
                response = requests.get(image_url, stream=True)
                response.raise_for_status()
                image_stream = io.BytesIO(response.content)
                # Add the downloaded image to the destination slide
                dest_slide.shapes.add_picture(image_stream, shape.left, shape.top, width=shape.width, height=shape.height)
                continue # Skip the generic element copy
            except (AttributeError, requests.exceptions.RequestException):
                # Fallback for embedded images or if download fails
                pass
        
        # For all other shapes (or fallback), copy the element
        new_el = copy.deepcopy(shape.element)
        dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

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
    """
    Populates a slide's placeholders with new content, preserving formatting by
    replacing text in existing runs.
    """
    title_populated, body_populated = False, False
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        
        is_placeholder = shape.is_placeholder
        
        if not title_populated and ((is_placeholder and shape.placeholder_format.type in ('TITLE', 'CENTER_TITLE')) or (not is_placeholder and shape.top < Pt(150))):
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = content.get("title", "")
            run.font.bold = True
            title_populated = True

        elif not body_populated and ((is_placeholder and shape.placeholder_format.type in ('BODY', 'OBJECT')) or "lorem ipsum" in shape.text.lower()):
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = content.get("body", "")
            run.font.bold = True
            body_populated = True

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
                new_prs = Presentation(io.BytesIO(template_files[0].getvalue()))
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                
                st.write("Step 2/3: Building new presentation from your structure...")
                
                if len(st.session_state.structure) != len(new_prs.slides):
                    st.warning(f"Warning: Your structure has {len(st.session_state.structure)} steps, but the template has {len(new_prs.slides)} slides. Output will match the template's slide count.")

                for i, dest_slide in enumerate(new_prs.slides):
                    if i >= len(st.session_state.structure): break 
                    
                    step = st.session_state.structure[i]
                    keyword = step["keyword"]
                    action = step["action"]
                    st.write(f"  - Modifying slide {i+1} for '{keyword}' with action '{action}'")

                    if action == "Copy from GTM (as is)":
                        src_slide = find_slide_by_ai(api_key, gtm_prs, keyword)
                        if src_slide:
                            deep_copy_slide(new_prs, dest_slide, src_slide)
                            st.success(f"  - Deep copied content for '{keyword}'.")
                        else:
                            st.warning(f"  - AI could not find '{keyword}' in GTM Deck. Leaving template slide as is.")

                    elif action == "Merge: Template Layout + GTM Content":
                        content_slide = find_slide_by_ai(api_key, gtm_prs, keyword)
                        if content_slide:
                            content = get_slide_content(content_slide)
                            populate_slide(dest_slide, content)
                            st.success(f"  - Merged content for '{keyword}'.")
                        else:
                            st.warning(f"  - AI could not find content for '{keyword}' in GTM Deck. Leaving template slide as is.")

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
