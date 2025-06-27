import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import copy
import uuid
import openai
import json

# --- Core PowerPoint Functions ---

def deep_copy_slide(dest_pres, src_slide):
    """
    Performs a stable, deep copy of a source slide into a destination presentation.
    This is the most robust method for the "Copy from GTM" action.
    """
    # Use a blank layout as a base for the new slide
    blank_layout = dest_pres.slide_layouts[6] 
    dest_slide = dest_pres.slides.add_slide(blank_layout)

    # Clear any default shapes from the blank layout
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)

    # Copy shapes from the source slide to the destination slide by duplicating their XML elements.
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

def find_slide_by_ai(api_key, prs, slide_type_prompt, deck_name):
    """
    Uses OpenAI to intelligently find the best matching slide and get a justification.
    Returns a dictionary with the slide object, its index, and the AI's justification.
    """
    if not slide_type_prompt: return {"slide": None, "index": -1, "justification": "No keyword provided."}
    client = openai.OpenAI(api_key=api_key)
    
    slides_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(prs.slides)]

    system_prompt = f"""
    You are an expert presentation analyst. Your task is to find the best slide in a presentation that matches a user's description.
    The user is looking for a slide representing: '{slide_type_prompt}'.
    Analyze the text of each slide to understand its purpose. A "Timeline" slide VISUALLY represents a schedule with dates, quarters, or sequential phases (Phase 1, Phase 2); it is NOT just a list in a table of contents. An "Objectives" slide will contain goal-oriented language. You must prioritize actual content slides over simple divider or table of contents pages.
    You MUST return a JSON object with two keys: 'best_match_index' (an integer, or -1 if no match) and 'justification' (a brief, one-sentence justification for your choice).
    """
    full_user_prompt = f"Find the best slide for '{slide_type_prompt}' in the '{deck_name}' deck with the following contents:\n{json.dumps(slides_content, indent=2)}"

    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo", messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": full_user_prompt}],
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        best_index = result.get("best_match_index", -1)
        justification = result.get("justification", "No justification provided.")
        if best_index != -1 and best_index < len(prs.slides):
            return {"slide": prs.slides[best_index], "index": best_index, "justification": justification}
        return {"slide": None, "index": -1, "justification": "AI could not find a suitable slide."}
    except Exception as e:
        return {"slide": None, "index": -1, "justification": f"An error occurred during analysis: {e}"}

def find_slide_in_templates(api_key, template_prs_list, slide_type_prompt):
    """Searches through all template presentations to find the best layout slide."""
    for i, prs in enumerate(template_prs_list):
        result = find_slide_by_ai(api_key, prs, slide_type_prompt, f"Template Deck {i+1}")
        if result and result["slide"]:
            # Return the first good match found across all templates
            return result
    return {"slide": None, "index": -1, "justification": "Could not find a suitable layout in any template deck."}

def get_slide_content(slide):
    """Extracts title and body text from a slide."""
    if not slide: return {"title": "", "body": ""}
    text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and s.text.strip()], key=lambda s: s.top)
    title = text_boxes[0].text if text_boxes else ""
    body = "\n".join(s.text for s in text_boxes[1:]) if len(text_boxes) > 1 else ""
    return {"title": title, "body": body}

def populate_slide(slide, content):
    """Populates a slide's placeholders with new content, making it bold."""
    title_shape, body_shape = None, None
    for shape in slide.shapes:
        if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
            if shape.placeholder_format.type in ('TITLE', 'CENTER_TITLE'): title_shape = shape
            elif shape.placeholder_format.type in ('BODY', 'OBJECT'): body_shape = shape
    if not body_shape: # Fallback for non-standard templates
         text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and "lorem ipsum" in s.text.lower()], key=lambda s: s.top)
         if text_boxes: body_shape = text_boxes[0]
    
    if title_shape:
        tf = title_shape.text_frame; tf.clear(); run = tf.add_paragraph().add_run(); run.text = content.get("title", ""); run.font.bold = True
    if body_shape:
        tf = body_shape.text_frame; tf.clear(); run = tf.add_paragraph().add_run(); run.text = content.get("body", ""); run.font.bold = True

# --- Streamlit App ---
st.set_page_config(page_title="Dynamic AI Presentation Assembler", layout="wide")
st.title("🤖 Dynamic AI Presentation Assembler")

with st.sidebar:
    st.header("1. API Key & Decks")
    api_key = st.text_input("OpenAI API Key", type="password")
    st.markdown("---")
    st.header("2. Upload Decks")
    template_files = st.file_uploader("Upload Template Deck(s)", type=["pptx"], accept_multiple_files=True)
    gtm_file = st.file_uploader("Upload GTM Global Deck", type=["pptx"])
    st.markdown("---")
    st.header("3. Define Presentation Structure")
    
    if 'structure' not in st.session_state: st.session_state.structure = []
    
    if st.button("Add New Step", use_container_width=True):
        st.session_state.structure.append({"id": str(uuid.uuid4()), "keyword": "", "action": "Copy from GTM (as is)"})

    for i, step in enumerate(st.session_state.structure):
        with st.container(border=True):
            cols = st.columns([3, 3, 1]); step["keyword"] = cols[0].text_input("Slide Type", step["keyword"], key=f"keyword_{step['id']}")
            step["action"] = cols[1].selectbox("Action", ["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"], index=["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"].index(step["action"]), key=f"action_{step['id']}")
            if cols[2].button("🗑️", key=f"del_{step['id']}"): st.session_state.structure.pop(i); st.rerun()
    if st.button("Clear Structure", use_container_width=True): st.session_state.structure = []; st.rerun()

# --- Main App Logic ---
if template_files and gtm_file and api_key and st.session_state.structure:
    if st.button("🚀 Assemble Presentation", type="primary"):
        with st.spinner("Assembling your new presentation..."):
            try:
                st.write("Step 1/3: Loading decks...")
                template_prs_list = [Presentation(io.BytesIO(f.getvalue())) for f in template_files]
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                
                # Create a new, blank presentation and set its dimensions from the first template
                new_prs = Presentation()
                new_prs.slide_width = template_prs_list[0].slide_width
                new_prs.slide_height = template_prs_list[0].slide_height
                
                process_log = []
                st.write("Step 2/3: Building new presentation from your defined structure...")
                
                for i, step in enumerate(st.session_state.structure):
                    keyword, action = step["keyword"], step["action"]
                    log_entry = {"step": i + 1, "keyword": keyword, "action": action, "log": []}
                    
                    if action == "Copy from GTM (as is)":
                        result = find_slide_by_ai(api_key, gtm_prs, keyword, "GTM Deck")
                        log_entry["log"].append(f"**GTM Content Choice Justification:** {result['justification']}")
                        if result["slide"]:
                            deep_copy_slide(new_prs, result["slide"])
                            log_entry["log"].append(f"**Action:** Copied slide {result['index'] + 1} from GTM Deck into the new presentation.")
                        else:
                            log_entry["log"].append("**Action:** No suitable slide found in GTM deck. This step was skipped.")
                    
                    elif action == "Merge: Template Layout + GTM Content":
                        layout_result = find_slide_in_templates(api_key, template_prs_list, keyword)
                        content_result = find_slide_by_ai(api_key, gtm_prs, keyword, "GTM Deck")
                        
                        log_entry["log"].append(f"**Template Layout Choice Justification:** {layout_result['justification']}")
                        log_entry["log"].append(f"**GTM Content Choice Justification:** {content_result['justification']}")

                        if layout_result["slide"] and content_result["slide"]:
                            content = get_slide_content(content_result["slide"])
                            new_slide = new_prs.slides.add_slide(layout_result["slide"].slide_layout)
                            populate_slide(new_slide, content)
                            log_entry["log"].append(f"**Action:** Merged content from GTM slide {content_result['index'] + 1} into a new slide using the layout from Template slide {layout_result['index'] + 1}.")
                        else:
                             log_entry["log"].append("**Action:** Could not find both a suitable layout and content. This step was skipped.")
                    
                    process_log.append(log_entry)

                st.success("Successfully built the new presentation structure.")
                
                st.write("Step 3/3: Finalizing...")
                st.subheader("📋 Process Log")
                for entry in process_log:
                    with st.expander(f"Step {entry['step']}: '{entry['keyword']}' ({entry['action']})"):
                        for line in entry['log']: st.markdown(f"- {line}")
                
                output_buffer = io.BytesIO(); new_prs.save(output_buffer); output_buffer.seek(0)
                st.success("🎉 Your new regional presentation has been assembled!")
                st.download_button("Download Assembled PowerPoint", data=output_buffer, file_name="Dynamic_AI_Assembled_Deck.pptx")
            except Exception as e:
                st.error(f"A critical error occurred: {e}"); st.exception(e)
else:
    st.info("Please provide an API Key, upload at least one Template Deck and a GTM Deck, and define the structure in the sidebar to begin.")
