import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import copy
import uuid
import openai
import json
from lxml import etree

# --- Core PowerPoint Functions ---

def clone_slide(pres, slide_to_clone):
    """
    Duplicates a slide from a source presentation and adds it to the end of
    the slides in the destination presentation `pres`. This is the most robust method.
    """
    # 1. Create a dictionary to map relationship IDs from the source to the target
    rId_map = {}
    
    # 2. Iterate through the relationships of the source slide
    for r in slide_to_clone.part.rels:
        rel = slide_to_clone.part.rels[r]
        # If the relationship is to an external resource (like a hyperlink), skip it
        if rel.is_external:
            continue
        
        # Get the target part of the relationship (e.g., an image)
        target_part = rel.target_part
        # Add the target part to the destination presentation's package
        # This will handle images, charts, etc., ensuring they are available for the new slide
        pres.part.package.get_part(target_part.partname)
        
        # Store the relationship ID mapping
        rId_map[rel.rId] = rel.rId

    # 3. Create a new blank slide in the destination presentation
    # We use a blank layout (usually index 6) as a starting point
    blank_slide_layout = pres.slide_layouts[6]
    new_slide = pres.slides.add_slide(blank_slide_layout)
    
    # 4. Replace the content of the new blank slide with the content of the slide to clone
    # This is done by directly manipulating the underlying XML elements
    new_slide.shapes.element.getparent().replace(new_slide.shapes.element, slide_to_clone.shapes.element)
    
    # 5. Copy the background from the source slide to the new slide
    if slide_to_clone.has_notes_slide:
        new_slide.has_notes_slide
        notes_slide = new_slide.notes_slide
        notes_slide.notes_text_frame.text = slide_to_clone.notes_slide.notes_text_frame.text

    return new_slide


def find_slide_by_ai(api_key, prs, slide_type_prompt):
    """
    Uses OpenAI to intelligently find the best matching slide in a presentation.
    """
    if not slide_type_prompt: return None
    client = openai.OpenAI(api_key=api_key)
    
    slides_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(prs.slides)]

    system_prompt = f"""
    You are an expert presentation analyst. You will be given a JSON list of slide contents and a user's description of a slide type.
    Your task is to identify the index of the single best-matching slide from the list.
    The user is looking for a slide that represents: '{slide_type_prompt}'.
    Analyze the text of each slide to understand its purpose. For example, a "Timeline" might contain dates, quarters, or sequential phases. "Objectives" might contain goal-oriented language.
    You MUST return a JSON object with a single key 'best_match_index', which is the integer index of the best-matching slide.
    If no suitable slide is found, return -1.
    """

    full_user_prompt = f"Find the best slide for '{slide_type_prompt}' in the following content:\n{json.dumps(slides_content, indent=2)}"

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
    title, body = "", ""
    text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and s.text.strip()], key=lambda s: s.top)
    if text_boxes:
        title = text_boxes[0].text
        if len(text_boxes) > 1:
            body = "\n".join(s.text for s in text_boxes[1:])
    return {"title": title, "body": body}

def populate_slide(slide, content):
    """Populates a slide's placeholders with new content, making it bold."""
    text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and "lorem ipsum" in s.text.lower()], key=lambda s: s.top)
    if not text_boxes: return

    for shape in text_boxes:
        shape.text_frame.clear()
        
    p = text_boxes[0].text_frame.add_paragraph()
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
                template_prs_list = [Presentation(io.BytesIO(f.getvalue())) for f in template_files]
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                
                # Use the first template to set the dimensions for the new presentation
                new_prs = Presentation()
                new_prs.slide_width = template_prs_list[0].slide_width
                new_prs.slide_height = template_prs_list[0].slide_height

                st.write("Step 2/3: Building presentation from your structure...")
                for i, step in enumerate(st.session_state.structure):
                    keyword = step["keyword"]
                    action = step["action"]
                    st.write(f"  - Processing '{keyword}' with action '{action}'")

                    if action == "Copy from GTM (as is)":
                        content_slide = find_slide_by_ai(api_key, gtm_prs, keyword)
                        if content_slide:
                            clone_slide(new_prs, content_slide)
                            st.success(f"  - Found and copied '{keyword}' slide.")
                        else:
                            st.warning(f"  - AI could not find a '{keyword}' slide in GTM deck. Skipping.")
                    
                    elif action == "Merge: Template Layout + GTM Content":
                        layout_slide = find_slide_by_ai(api_key, template_prs_list[0], keyword)
                        content_slide = find_slide_by_ai(api_key, gtm_prs, keyword)

                        if layout_slide and content_slide:
                            content = get_slide_content(content_slide)
                            new_slide = clone_slide(new_prs, layout_slide)
                            populate_slide(new_slide, content)
                            st.success(f"  - Found and merged '{keyword}' slide.")
                        else:
                            if not layout_slide: st.warning(f"  - AI could not find layout for '{keyword}' in Template. Skipping.")
                            if not content_slide: st.warning(f"  - AI could not find content for '{keyword}' in GTM Deck. Skipping.")
                
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
