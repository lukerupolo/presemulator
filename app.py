import streamlit as st
from pptx import Presentation
import io
import copy
import uuid
import openai
import json
import requests
from PIL import Image
import os

# --- Core PowerPoint Functions ---

def find_slide_by_ai(api_key, prs, slide_type_prompt, all_slides_content):
    """Uses OpenAI to intelligently find the best matching slide and get a justification."""
    client = openai.OpenAI(api_key=api_key)
    system_prompt = f"""
    You are an expert presentation analyst. Given a JSON list of slide contents and a description ('{slide_type_prompt}'), identify the index of the single best-matching slide.
    Analyze the text for purpose (e.g., a "Timeline" has dates or sequential phases; "Objectives" has goal-oriented language).
    Return a JSON object with two keys: 'best_match_index' (an integer, or -1 if no match) and 'justification' (a brief explanation for your choice).
    """
    full_user_prompt = f"Find the best slide for '{slide_type_prompt}' in: {json.dumps(all_slides_content, indent=2)}"
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
        return {"slide": None, "index": -1, "justification": f"An error occurred: {e}"}

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
        tf = title_shape.text_frame
        tf.clear()
        run = tf.add_paragraph().add_run()
        run.text = content.get("title", "")
        run.font.bold = True
    if body_shape:
        tf = body_shape.text_frame
        tf.clear()
        run = tf.add_paragraph().add_run()
        run.text = content.get("body", "")
        run.font.bold = True

def deep_copy_slide(dest_pres, src_slide):
    """Deep copies a slide from source to destination presentation."""
    dest_layout = dest_pres.slide_layouts[6] # Using a blank layout
    dest_slide = dest_pres.slides.add_slide(dest_layout)
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

def get_slide_thumbnail(slide):
    """Creates a PIL Image thumbnail of a slide."""
    title = get_slide_content(slide).get("title", f"Slide {slide.slide_id}")
    img = Image.new('RGB', (400, 300), color = 'white')
    from PIL import ImageDraw
    d = ImageDraw.Draw(img)
    d.text((10,10), f"Thumbnail for:\n{title[:50]}...", fill=(0,0,0))
    return img

# --- Streamlit App ---
st.set_page_config(page_title="Interactive AI Presentation Assembler", layout="wide")
st.title("🤖 Interactive AI Presentation Assembler")

# --- Initialize Session State ---
if 'structure' not in st.session_state: st.session_state.structure = []
if 'build_plan' not in st.session_state: st.session_state.build_plan = None
if 'assembly_started' not in st.session_state: st.session_state.assembly_started = False

with st.sidebar:
    st.header("1. API Key & Decks")
    api_key = st.text_input("OpenAI API Key", type="password")
    template_file = st.file_uploader("Upload Template Deck", type=["pptx"])
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
                st.session_state.structure.pop(i)
                st.rerun()
    
    if st.button("Clear Structure", use_container_width=True):
        st.session_state.structure = []
        st.session_state.build_plan = None
        st.session_state.assembly_started = False
        st.rerun()

# --- Main App Logic ---
# FIX: Corrected typo from gtm_.file to gtm_file
if template_file and gtm_file and api_key and st.session_state.structure:
    # Generate Plan Button
    if st.button("1. Generate Build Plan", type="primary"):
        st.session_state.assembly_started = True
        with st.spinner("Analyzing decks and creating initial plan..."):
            template_prs = Presentation(io.BytesIO(template_file.getvalue()))
            gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
            
            template_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(template_prs.slides)]
            gtm_content = [{"slide_index": i, "text": " ".join(s.text for s in slide.shapes if s.has_text_frame)[:1000]} for i, slide in enumerate(gtm_prs.slides)]

            plan = []
            for step in st.session_state.structure:
                keyword = step["keyword"]
                action = step["action"]
                if action == "Copy from GTM (as is)":
                    result = find_slide_by_ai(api_key, gtm_prs, keyword, gtm_content)
                    plan.append({"keyword": keyword, "action": action, "gtm_choice": result})
                elif action == "Merge: Template Layout + GTM Content":
                    layout_result = find_slide_by_ai(api_key, template_prs, keyword, template_content)
                    content_result = find_slide_by_ai(api_key, gtm_prs, keyword, gtm_content)
                    plan.append({"keyword": keyword, "action": action, "template_choice": layout_result, "gtm_choice": content_result})
            st.session_state.build_plan = plan
            st.session_state.template_prs = template_prs
            st.session_state.gtm_prs = gtm_prs


if st.session_state.assembly_started and st.session_state.build_plan:
    st.markdown("---")
    st.header("3. Review and Approve Build Plan")

    # Display the interactive build plan
    for i, item in enumerate(st.session_state.build_plan):
        st.subheader(f"Step {i+1}: '{item['keyword']}'")
        
        if item['action'] == "Copy from GTM (as is)":
            gtm_choice = item['gtm_choice']
            st.info(f"AI Justification (Content): {gtm_choice['justification']}")
            if gtm_choice['slide']:
                img = get_slide_thumbnail(gtm_choice['slide'])
                st.image(img, caption=f"Selected GTM Slide {gtm_choice['index']+1}")
        
        elif item['action'] == "Merge: Template Layout + GTM Content":
            template_choice = item['template_choice']
            gtm_choice = item['gtm_choice']
            
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"AI Justification (Layout): {template_choice['justification']}")
                if template_choice['slide']:
                    img = get_slide_thumbnail(template_choice['slide'])
                    st.image(img, caption=f"Selected Template Layout Slide {template_choice['index']+1}")
            with col2:
                st.info(f"AI Justification (Content): {gtm_choice['justification']}")
                if gtm_choice['slide']:
                    img = get_slide_thumbnail(gtm_choice['slide'])
                    st.image(img, caption=f"Selected GTM Content Slide {gtm_choice['index']+1}")

    st.markdown("---")
    # Assemble Button
    if st.button("2. Assemble Final Presentation", type="primary"):
        with st.spinner("Executing final assembly..."):
            final_prs = Presentation()
            final_prs.slide_width = st.session_state.template_prs.slide_width
            final_prs.slide_height = st.session_state.template_prs.slide_height

            for item in st.session_state.build_plan:
                if item['action'] == "Copy from GTM (as is)":
                    if item['gtm_choice']['slide']:
                        deep_copy_slide(final_prs, item['gtm_choice']['slide'])
                elif item['action'] == "Merge: Template Layout + GTM Content":
                    if item['template_choice']['slide'] and item['gtm_choice']['slide']:
                        content = get_slide_content(item['gtm_choice']['slide'])
                        new_slide = final_prs.slides.add_slide(item['template_choice']['slide'].slide_layout)
                        populate_slide(new_slide, content)
            
            output_buffer = io.BytesIO()
            final_prs.save(output_buffer)
            output_buffer.seek(0)
            st.success("🎉 Your new regional presentation has been assembled!")
            st.download_button("Download Assembled PowerPoint", data=output_buffer, file_name="Interactive_AI_Assembled_Deck.pptx")
            
            # Clean up session state
            st.session_state.build_plan = None
            st.session_state.assembly_started = False
