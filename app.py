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
    Performs a stable, deep copy of all shapes and content from a source slide
    to a destination slide.
    """
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

def find_slide_by_ai(api_key, prs, slide_type_prompt):
    """
    Uses OpenAI to intelligently find the best matching slide and get a justification.
    Returns a dictionary with the slide object, its index, and the AI's justification.
    """
    if not slide_type_prompt: return None
    client = openai.OpenAI(api_key=api_key)
    
    slides_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(prs.slides)]

    system_prompt = f"""
    You are an expert presentation analyst. Given a JSON list of slide contents and a description of a slide type ('{slide_type_prompt}'), your task is twofold:
    1. Identify the index of the single best-matching slide. A "Timeline" might contain dates, quarters, or sequential phases. "Objectives" might contain goal-oriented language.
    2. Provide a brief, one-sentence justification for your choice based on the slide's text content.
    You MUST return a JSON object with two keys: 'best_match_index' (an integer, or -1 if no match) and 'justification' (a string).
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
        justification = result.get("justification", "No justification provided.")

        if best_index != -1 and best_index < len(prs.slides):
            return {"slide": prs.slides[best_index], "index": best_index, "justification": justification}
        return {"slide": None, "index": -1, "justification": "AI could not find a suitable slide."}
    except Exception as e:
        return {"slide": None, "index": -1, "justification": f"An error occurred during analysis: {e}"}

def get_slide_content(slide):
    """Extracts title and body text from a slide."""
    text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and s.text.strip()], key=lambda s: s.top)
    title = text_boxes[0].text if text_boxes else ""
    body = "\n".join(s.text for s in text_boxes[1:]) if len(text_boxes) > 1 else ""
    return {"title": title, "body": body}

def populate_slide(slide, content):
    """Populates a slide's placeholders with new content, making it bold."""
    title_populated, body_populated = False, False
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        is_placeholder = shape.is_placeholder
        
        if not title_populated and ((is_placeholder and shape.placeholder_format.type in ('TITLE', 'CENTER_TITLE')) or (not is_placeholder and shape.top < Pt(150))):
            tf = shape.text_frame
            tf.clear()
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = content.get("title", "")
            run.font.bold = True
            title_populated = True
        elif not body_populated and ((is_placeholder and shape.placeholder_format.type in ('BODY', 'OBJECT')) or "lorem ipsum" in shape.text.lower()):
            tf = shape.text_frame
            tf.clear()
            p = tf.add_paragraph()
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
    
    if 'structure' not in st.session_state: st.session_state.structure = []
    
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
        with st.spinner("Assembling your new presentation..."):
            try:
                st.write("Step 1/4: Loading decks...")
                new_prs = Presentation(io.BytesIO(template_files[0].getvalue()))
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                template_prs = Presentation(io.BytesIO(template_files[0].getvalue())) # Re-load for iteration
                
                process_log = []
                build_plan = []

                st.write("Step 2/4: Planning the build...")
                for i, step in enumerate(st.session_state.structure):
                    keyword = step["keyword"]
                    action = step["action"]
                    
                    if action == "Copy from GTM (as is)":
                        result = find_slide_by_ai(api_key, gtm_prs, keyword)
                        build_plan.append({"action": "copy", "result": result, "keyword": keyword})
                    
                    elif action == "Merge: Template Layout + GTM Content":
                        layout_result = find_slide_by_ai(api_key, template_prs, keyword)
                        content_result = find_slide_by_ai(api_key, gtm_prs, keyword)
                        build_plan.append({"action": "merge", "layout_result": layout_result, "content_result": content_result, "keyword": keyword})

                # Clear the base presentation to build fresh
                for i in range(len(new_prs.slides) - 1, -1, -1):
                    rId = new_prs.slides._sldIdLst[i].rId
                    new_prs.part.drop_rel(rId)
                    del new_prs.slides._sldIdLst[i]

                st.write("Step 3/4: Executing the build plan...")
                for i, item in enumerate(build_plan):
                    keyword = item["keyword"]
                    log_entry = {"step": i + 1, "keyword": keyword, "action": item["action"], "log": []}

                    if item["action"] == "copy":
                        result = item["result"]
                        log_entry["log"].append(f"Searched GTM deck for '{keyword}'.")
                        log_entry["log"].append(f"AI Justification: {result['justification']}")
                        if result["slide"]:
                            # For Copy, we need a different function that adds a new slide
                            # For simplicity, we'll deep copy into a new slide added to the end
                            new_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6]) # Add blank slide
                            deep_copy_slide(new_slide, result["slide"])
                            log_entry["log"].append(f"Action: Copied slide {result['index'] + 1} from GTM Deck.")
                        else:
                            log_entry["log"].append("Action: No suitable slide found. Skipped.")
                    
                    elif item["action"] == "merge":
                        layout_result = item["layout_result"]
                        content_result = item["content_result"]
                        log_entry["log"].append(f"Searched Template for '{keyword}' layout. Justification: {layout_result['justification']}")
                        log_entry["log"].append(f"Searched GTM Deck for '{keyword}' content. Justification: {content_result['justification']}")
                        
                        if layout_result["slide"] and content_result["slide"]:
                            # Add a new slide using the found layout
                            new_slide = new_prs.slides.add_slide(layout_result["slide"].slide_layout)
                            content = get_slide_content(content_result["slide"])
                            populate_slide(new_slide, content)
                            log_entry["log"].append(f"Action: Merged content from GTM slide {content_result['index'] + 1} into a new slide using layout from Template slide {layout_result['index'] + 1}.")
                        else:
                             log_entry["log"].append("Action: Could not find both layout and content. Skipped.")
                    
                    process_log.append(log_entry)

                st.success("Successfully built the new presentation.")
                
                # --- Step 4: Display Log and Download ---
                st.write("Step 4/4: Finalizing...")
                st.subheader("📋 Process Log")
                for entry in process_log:
                    with st.expander(f"Step {entry['step']}: '{entry['keyword']}' ({entry['action']})"):
                        for line in entry['log']:
                            if "Justification:" in line: st.info(line)
                            elif "Action:" in line: st.success(line)
                            else: st.write(line)
                
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
