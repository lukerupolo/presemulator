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
    Performs a stable, deep copy of all shapes and content from a source slide
    to a destination presentation by adding a new slide and then copying shapes.
    """
    blank_layout = dest_pres.slide_layouts[6] 
    dest_slide = dest_pres.slides.add_slide(blank_layout)
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

def find_slide_candidates_by_ai(api_key, prs, slide_type_prompt, all_slides_content):
    """
    Uses OpenAI to intelligently find the top 3 best matching slides and get justifications.
    """
    client = openai.OpenAI(api_key=api_key)
    system_prompt = f"""
    You are an expert presentation analyst. Given a JSON list of slide contents and a description ('{slide_type_prompt}'), your task is to identify the top 3 best-matching slides.
    A "Timeline" slide VISUALLY represents a schedule with dates, quarters, or sequential phases (Phase 1, Phase 2), it is NOT just a list in a table of contents. An "Objectives" slide contains goal-oriented language. You must prioritize slides that are actual content slides over divider or table of contents pages.
    Return a JSON object with a single key 'best_matches', which is a list of objects. Each object must have two keys: 'index' (the integer index of the slide) and 'justification' (a brief explanation for your choice). Return up to 3 matches, sorted from best to worst. If no matches are found, return an empty list.
    """
    full_user_prompt = f"Find the best slides for '{slide_type_prompt}' in: {json.dumps(all_slides_content, indent=2)}"
    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo", messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": full_user_prompt}],
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        matches = result.get("best_matches", [])
        
        valid_matches = []
        for match in matches:
            idx = match.get("index")
            if idx is not None and 0 <= idx < len(prs.slides):
                match["slide"] = prs.slides[idx]
                valid_matches.append(match)
        return valid_matches
    except Exception as e:
        st.error(f"AI slide analysis failed: {e}")
        return []

def get_slide_content(slide):
    """Extracts title and all other text from a slide."""
    if not slide: return {"title": "", "body": ""}
    text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and s.text and s.text.strip()], key=lambda s: s.top)
    title = text_boxes[0].text if text_boxes else ""
    body = "\n---\n".join(s.text for s in text_boxes[1:]) if len(text_boxes) > 1 else ""
    return {"title": title, "body": body}

def populate_slide(slide, content):
    """Populates a slide's placeholders with new content, making it bold."""
    title_shape, body_shape = None, None
    for shape in slide.shapes:
        if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
            if shape.placeholder_format.type in ('TITLE', 'CENTER_TITLE'): title_shape = shape
            elif shape.placeholder_format.type in ('BODY', 'OBJECT'): body_shape = shape
    if not body_shape:
         text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and "lorem ipsum" in s.text.lower()], key=lambda s: s.top)
         if text_boxes: body_shape = text_boxes[0]
    if title_shape:
        tf = title_shape.text_frame; tf.clear(); run = tf.add_paragraph().add_run(); run.text = content.get("title", ""); run.font.bold = True
    if body_shape:
        tf = body_shape.text_frame; tf.clear(); run = tf.add_paragraph().add_run(); run.text = content.get("body", ""); run.font.bold = True

def generate_slide_html_preview(slide, slide_width, slide_height):
    """Generates an HTML representation of the slide's text layout."""
    # Start the HTML with a container div and basic styles
    html = """
    <div style="border: 1px solid #ddd; background-color: #f9f9f9; padding: 10px; height: 300px; position: relative; font-family: sans-serif; font-size: 10px;">
    """
    # Iterate through shapes and create divs for them
    for shape in slide.shapes:
        if not shape.has_text_frame or not shape.text.strip():
            continue
        
        # Calculate proportional positions and dimensions
        left_perc = (shape.left / slide_width) * 100
        top_perc = (shape.top / slide_height) * 100
        width_perc = (shape.width / slide_width) * 100
        height_perc = (shape.height / slide_height) * 100

        # Truncate text for display
        text = (shape.text[:150] + '...') if len(shape.text) > 150 else shape.text
        
        # Add a div for each shape
        html += f"""
        <div style="position: absolute; left: {left_perc}%; top: {top_perc}%; width: {width_perc}%; height: {height_perc}%;
                    border: 1px dashed #ccc; background-color: rgba(200, 200, 200, 0.1); overflow: hidden; padding: 5px;">
            {text}
        </div>
        """
    html += "</div>"
    return html


# --- Streamlit App ---
st.set_page_config(page_title="Interactive AI Presentation Assembler", layout="wide")
st.title("🤖 Interactive AI Presentation Assembler")

if 'structure' not in st.session_state: st.session_state.structure = []
if 'build_plan' not in st.session_state: st.session_state.build_plan = None

with st.sidebar:
    st.header("1. API Key & Decks")
    api_key = st.text_input("OpenAI API Key", type="password")
    template_files = st.file_uploader("Upload Template Deck(s)", type=["pptx"], accept_multiple_files=True)
    gtm_file = st.file_uploader("Upload GTM Global Deck", type=["pptx"])
    st.markdown("---")
    st.header("2. Define Presentation Structure")
    if st.button("Add New Step", use_container_width=True):
        st.session_state.structure.append({"id": str(uuid.uuid4()), "keyword": "", "action": "Copy from GTM (as is)", "user_selection": 0})
    for i, step in enumerate(st.session_state.structure):
        with st.container(border=True):
            cols = st.columns([3, 3, 1]); step["keyword"] = cols[0].text_input("Slide Type", step["keyword"], key=f"keyword_{step['id']}")
            step["action"] = cols[1].selectbox("Action", ["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"], index=["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"].index(step["action"]), key=f"action_{step['id']}")
            if cols[2].button("🗑️", key=f"del_{step['id']}"): st.session_state.structure.pop(i); st.rerun()
    if st.button("Clear Structure", use_container_width=True): st.session_state.structure = []; st.session_state.build_plan = None; st.rerun()

# --- Main App Logic ---
if template_files and gtm_file and api_key and st.session_state.structure:
    if st.button("1. Generate Build Plan", type="primary"):
        with st.spinner("Analyzing decks and creating initial plan..."):
            template_prs = Presentation(io.BytesIO(template_files[0].getvalue()))
            gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
            template_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(template_prs.slides)]
            gtm_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(gtm_prs.slides)]
            plan = []
            for step in st.session_state.structure:
                keyword, action = step["keyword"], step["action"]
                item = {"keyword": keyword, "action": action, "template_choices": [], "gtm_choice": None, "user_selection": 0}
                if action == "Copy from GTM (as is)":
                    result = find_slide_candidates_by_ai(api_key, gtm_prs, keyword, gtm_content)
                    if result: item["gtm_choice"] = result[0]
                elif action == "Merge: Template Layout + GTM Content":
                    item["template_choices"] = find_slide_candidates_by_ai(api_key, template_prs, keyword, template_content)
                    content_result = find_slide_candidates_by_ai(api_key, gtm_prs, keyword, gtm_content)
                    if content_result: item["gtm_choice"] = content_result[0]
                plan.append(item)
            st.session_state.build_plan = plan
            st.session_state.template_prs = template_prs; st.session_state.gtm_prs = gtm_prs

if st.session_state.build_plan:
    st.markdown("---"); st.header("3. Review and Approve Build Plan")
    for i, item in enumerate(st.session_state.build_plan):
        with st.container(border=True):
            st.subheader(f"Step {i+1}: '{item['keyword']}' ({item['action']})")
            if item['action'] == "Merge: Template Layout + GTM Content":
                st.markdown("**Select a Layout from the Template Deck:**")
                if item["template_choices"]:
                    options = [f"Option {j+1}: Slide {choice['index']+1}" for j, choice in enumerate(item["template_choices"])]
                    selection_index = st.radio("Layout Options", range(len(options)), format_func=lambda x: options[x], key=f"select_{item['keyword']}", horizontal=True)
                    item['user_selection'] = selection_index
                    selected_choice = item["template_choices"][selection_index]
                    st.info(f"AI Justification: {selected_choice['justification']}")
                    # NEW: Display HTML preview instead of text expander
                    st.html(generate_slide_html_preview(selected_choice['slide'], st.session_state.template_prs.slide_width, st.session_state.template_prs.slide_height))
                else:
                    st.warning("AI found no suitable layouts in the Template Deck.")
            
            if item['gtm_choice'] and item['gtm_choice']['slide']:
                 st.markdown("**Content Source from GTM Deck:**")
                 st.info(f"AI Justification: {item['gtm_choice']['justification']}")
                 st.html(generate_slide_html_preview(item['gtm_choice']['slide'], st.session_state.gtm_prs.slide_width, st.session_state.gtm_prs.slide_height))
    
    st.markdown("---")
    if st.button("2. Assemble Final Presentation", type="primary"):
        with st.spinner("Executing final assembly..."):
            final_prs = Presentation(); final_prs.slide_width = st.session_state.template_prs.slide_width; final_prs.slide_height = st.session_state.template_prs.slide_height
            for item in st.session_state.build_plan:
                if item['action'] == "Copy from GTM (as is)":
                    if item['gtm_choice'] and item['gtm_choice']['slide']:
                        deep_copy_slide(final_prs, item['gtm_choice']['slide'])
                elif item['action'] == "Merge: Template Layout + GTM Content":
                    if item['template_choices'] and item['gtm_choice'] and item['gtm_choice']['slide']:
                        selected_layout_slide = item['template_choices'][item['user_selection']]['slide']
                        content = get_slide_content(item['gtm_choice']['slide'])
                        new_slide = final_prs.slides.add_slide(selected_layout_slide.slide_layout)
                        populate_slide(new_slide, content)
            
            output_buffer = io.BytesIO(); final_prs.save(output_buffer); output_buffer.seek(0)
            st.success("🎉 Your presentation has been assembled!"); st.download_button("Download Assembled PowerPoint", data=output_buffer, file_name="Interactive_AI_Assembled_Deck.pptx")
            st.session_state.build_plan = None
