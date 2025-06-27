import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import copy
import uuid
import openai
import json
from PIL import Image, ImageDraw

# --- Core PowerPoint Functions ---

def deep_copy_slide(dest_pres, src_slide):
    """Deep copies a slide from source to destination presentation."""
    dest_layout = dest_pres.slide_layouts[6]
    dest_slide = dest_pres.slides.add_slide(dest_layout)
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
    Analyze the text for purpose. A "Timeline" slide often contains dates, quarters, or sequential phases (Phase 1, Phase 2). It is a visual representation of a schedule, not just a list. An "Objectives" slide will have goal-oriented language. Be careful not to select divider or table of contents pages.
    You MUST return a JSON object with a single key 'best_matches', which is a list of objects. Each object must have two keys: 'index' (the integer index of the slide) and 'justification' (a brief explanation for your choice). Return up to 3 matches, sorted from best to worst. If no matches are found, return an empty list.
    """
    full_user_prompt = f"Find the best slides for '{slide_type_prompt}' in: {json.dumps(all_slides_content, indent=2)}"
    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo", messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": full_user_prompt}],
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        matches = result.get("best_matches", [])
        
        # Validate indices and create full result objects
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
    """Extracts title and body text from a slide."""
    text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and s.text.strip()], key=lambda s: s.top)
    title = text_boxes[0].text if text_boxes else ""
    body = "\n".join(s.text for s in text_boxes[1:]) if len(text_boxes) > 1 else ""
    return {"title": title, "body": body}

def populate_slide(slide, content):
    """Populates a slide's placeholders with new content, making it bold."""
    title_shape, body_shape = None, None
    for shape in slide.shapes:
        if shape.is_placeholder:
            if shape.placeholder_format.type in ('TITLE', 'CENTER_TITLE'): title_shape = shape
            elif shape.placeholder_format.type in ('BODY', 'OBJECT'): body_shape = shape
    if not body_shape:
         text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and "lorem ipsum" in s.text.lower()], key=lambda s: s.top)
         if text_boxes: body_shape = text_boxes[0]
    if title_shape:
        tf = title_shape.text_frame; tf.clear(); run = tf.add_paragraph().add_run(); run.text = content.get("title", ""); run.font.bold = True
    if body_shape:
        tf = body_shape.text_frame; tf.clear(); run = tf.add_paragraph().add_run(); run.text = content.get("body", ""); run.font.bold = True

def get_slide_thumbnail(slide, index):
    """Creates a basic placeholder thumbnail image for a slide."""
    title = get_slide_content(slide).get("title", f"Slide {index + 1}")
    img = Image.new('RGB', (300, 225), color = '#F0F2F6')
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 15)
    except IOError:
        font = ImageFont.load_default()
    d.text((10,10), f"PREVIEW: SLIDE {index + 1}\n\n{title[:80]}...", fill=(0,0,0), font=font)
    return img


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
        st.session_state.structure.append({"id": str(uuid.uuid4()), "keyword": "", "action": "Copy from GTM (as is)"})
    for i, step in enumerate(st.session_state.structure):
        with st.container(border=True):
            cols = st.columns([3, 3, 1])
            step["keyword"] = cols[0].text_input("Slide Type", step["keyword"], key=f"keyword_{step['id']}")
            step["action"] = cols[1].selectbox("Action", ["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"], index=["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"].index(step["action"]), key=f"action_{step['id']}")
            if cols[2].button("🗑️", key=f"del_{step['id']}"):
                st.session_state.structure.pop(i); st.rerun()
    if st.button("Clear Structure", use_container_width=True):
        st.session_state.structure = []; st.session_state.build_plan = None; st.rerun()

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
                if action == "Copy from GTM (as is)":
                    result = find_slide_candidates_by_ai(api_key, gtm_prs, keyword, gtm_content)
                    plan.append({"keyword": keyword, "action": action, "gtm_choices": result, "user_selection": 0 if result else -1})
                elif action == "Merge: Template Layout + GTM Content":
                    layout_results = find_slide_candidates_by_ai(api_key, template_prs, keyword, template_content)
                    content_result = find_slide_candidates_by_ai(api_key, gtm_prs, keyword, gtm_content)
                    plan.append({"keyword": keyword, "action": action, "template_choices": layout_results, "gtm_choice": content_result[0] if content_result else None, "user_selection": 0 if layout_results else -1})
            st.session_state.build_plan = plan
            st.session_state.template_prs = template_prs
            st.session_state.gtm_prs = gtm_prs

if st.session_state.build_plan:
    st.markdown("---"); st.header("3. Review and Approve Build Plan")
    for i, item in enumerate(st.session_state.build_plan):
        with st.container(border=True):
            st.subheader(f"Step {i+1}: '{item['keyword']}' ({item['action']})")
            if item['action'] == "Merge: Template Layout + GTM Content":
                st.write("**Select a Layout from the Template Deck:**")
                if item["template_choices"]:
                    cols = st.columns(len(item["template_choices"]))
                    for j, choice in enumerate(item["template_choices"]):
                        with cols[j]:
                            st.image(get_slide_thumbnail(choice['slide'], choice['index']), use_column_width=True)
                            st.info(f"AI says: {choice['justification']}")
                            if st.button(f"Use Layout {j+1}", key=f"select_{item['keyword']}_{j}"):
                                item["user_selection"] = j
                                st.rerun()
                else:
                    st.warning("AI found no suitable layouts in the Template Deck.")
            
            if item["user_selection"] != -1:
                st.success(f"**Selected Layout:** Option {item['user_selection'] + 1}")
    
    st.markdown("---")
    if st.button("2. Assemble Final Presentation", type="primary"):
        with st.spinner("Executing final assembly..."):
            final_prs = Presentation()
            final_prs.slide_width = st.session_state.template_prs.slide_width
            final_prs.slide_height = st.session_state.template_prs.slide_height
            for item in st.session_state.build_plan:
                if item['action'] == "Copy from GTM (as is)":
                    if item['gtm_choices']: deep_copy_slide(final_prs, item['gtm_choices'][0]['slide'])
                elif item['action'] == "Merge: Template Layout + GTM Content":
                    if item['template_choices'] and item['gtm_choice']:
                        selected_layout_slide = item['template_choices'][item['user_selection']]['slide']
                        content = get_slide_content(item['gtm_choice']['slide'])
                        new_slide = final_prs.slides.add_slide(selected_layout_slide.slide_layout)
                        populate_slide(new_slide, content)
            output_buffer = io.BytesIO()
            final_prs.save(output_buffer)
            output_buffer.seek(0)
            st.success("🎉 Your presentation has been assembled!")
            st.download_button("Download Assembled PowerPoint", data=output_buffer, file_name="Interactive_AI_Assembled_Deck.pptx")
            st.session_state.build_plan = None
