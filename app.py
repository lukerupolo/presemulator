import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import copy
import uuid
import openai
import json

# --- Core PowerPoint Functions ---

def clone_slide(pres, slide_to_clone):
    """
    Duplicates a slide from a source presentation and adds it to the end of
    the slides in the destination presentation `pres`. This is the most robust method.
    """
    src_part = slide_to_clone.part
    package = pres.part.package
    
    if package.has_part(src_part.partname):
        # This part is tricky. If a part (like an image) is already in the package
        # we should ideally reuse it. For simplicity and stability, we'll allow
        # python-pptx to handle part naming by just adding it.
        pass

    new_part = package.add_part(
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
            package.add_part(
                target_part.partname, target_part.content_type, target_part.blob
            )
        new_part.relate_to(target_part, rel.reltype, rId=rel.rId)

    return pres.slides[-1]

def find_slide_by_ai(api_key, prs, slide_type_prompt):
    """
    Uses OpenAI to intelligently find the best matching slide and get a justification.
    Returns a dictionary with the slide object, its index, and the AI's justification.
    """
    if not slide_type_prompt: return None
    client = openai.OpenAI(api_key=api_key)
    
    slides_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(prs.slides)]

    system_prompt = f"""
    You are an expert presentation analyst. Given a JSON list of slide contents and a user's description ('{slide_type_prompt}'), your task is twofold:
    1. Identify the index of the single best-matching slide.
    2. Provide a brief justification for your choice based on the slide's text content.
    Analyze the text for purpose. For example, a "Timeline" might contain dates, quarters, or sequential phases. "Objectives" might contain goal-oriented language.
    You MUST return a JSON object with two keys: 'best_match_index' (an integer, or -1 if no match) and 'justification' (a string explaining your choice).
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
        st.error(f"AI slide analysis failed for '{slide_type_prompt}': {e}")
        return {"slide": None, "index": -1, "justification": f"An error occurred during analysis: {e}"}

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
        if len(text_boxes) > 1: text_boxes = text_boxes[1:]
        elif not text_boxes:
            st.warning("Could not find a placeholder to populate on a merged slide.")
            return

    for shape in text_boxes: shape.text_frame.clear()
        
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
                template_prs_list = [Presentation(io.BytesIO(f.getvalue())) for f in template_files]
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                
                # CRITICAL FIX: Start with a completely blank presentation, but set its dimensions
                new_prs = Presentation()
                new_prs.slide_width = template_prs_list[0].slide_width
                new_prs.slide_height = template_prs_list[0].slide_height

                process_log = []

                st.write("Step 2/4: Building presentation from your defined structure...")
                for i, step in enumerate(st.session_state.structure):
                    keyword = step["keyword"]
                    action = step["action"]
                    log_entry = {"step": i + 1, "keyword": keyword, "action": action, "log": []}

                    if action == "Copy from GTM (as is)":
                        result = find_slide_by_ai(api_key, gtm_prs, keyword)
                        log_entry["log"].append(f"AI searched GTM deck for '{keyword}'.")
                        log_entry["log"].append(f"AI Justification: {result['justification']}")
                        if result["slide"]:
                            clone_slide(new_prs, result["slide"])
                            log_entry["log"].append(f"Action: Copied slide {result['index'] + 1} from GTM Deck.")
                        else:
                            log_entry["log"].append("Action: No suitable slide found. Skipped.")
                    
                    elif action == "Merge: Template Layout + GTM Content":
                        log_entry["log"].append(f"AI searched Template deck for '{keyword}' layout.")
                        layout_result = find_slide_by_ai(api_key, template_prs_list[0], keyword)
                        log_entry["log"].append(f"Layout Justification: {layout_result['justification']}")

                        log_entry["log"].append(f"AI searched GTM deck for '{keyword}' content.")
                        content_result = find_slide_by_ai(api_key, gtm_prs, keyword)
                        log_entry["log"].append(f"Content Justification: {content_result['justification']}")

                        if layout_result["slide"] and content_result["slide"]:
                            content = get_slide_content(content_result["slide"])
                            new_slide = clone_slide(new_prs, layout_result["slide"])
                            populate_slide(new_slide, content)
                            log_entry["log"].append(f"Action: Merged content from GTM slide {content_result['index'] + 1} into layout from Template slide {layout_result['index'] + 1}.")
                        else:
                             log_entry["log"].append("Action: Could not find both layout and content. Skipped.")

                    process_log.append(log_entry)
                
                st.success("Successfully built the new presentation structure.")
                
                # --- Step 3: Display the Process Log ---
                st.write("Step 3/4: Displaying Process Log...")
                st.subheader("📋 Process Log")
                for entry in process_log:
                    with st.expander(f"Step {entry['step']}: '{entry['keyword']}' ({entry['action']})"):
                        for line in entry['log']:
                            if "Justification:" in line:
                                st.info(line)
                            elif "Action:" in line:
                                st.success(line)
                            else:
                                st.write(line)
                
                # --- Step 4: Finalize and download ---
                st.write("Step 4/4: Finalizing and preparing download...")
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
