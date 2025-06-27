import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import copy
import uuid
import openai
import json
import requests

# --- Core PowerPoint Functions ---

def deep_copy_slide(dest_pres, dest_slide, src_slide):
    """
    Performs a stable, deep copy of all shapes and content from a source slide
    to a destination slide. This new version correctly handles linked images
    by downloading and embedding them.
    """
    # Clear all shapes from the destination slide first to prepare it.
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)

    # Iterate through shapes in the source slide and copy them to the destination.
    for shape in src_slide.shapes:
        # Check if the shape is a picture and if it is linked externally.
        if shape.shape_type == 13: # 13 is the shape type for Picture
            # The relationship ID for the image is in the 'r:link' attribute for linked images
            # and 'r:embed' for embedded images. We need to parse the XML to find it.
            blip_element = shape.element.xpath('.//a:blip', namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
            if blip_element:
                blip = blip_element[0]
                rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link')
                
                if rId:
                    # It's a linked image. We need to find the URL in the relationships file.
                    try:
                        image_url = src_slide.part.rels[rId].target_ref
                        # Download the image
                        response = requests.get(image_url, stream=True)
                        response.raise_for_status()
                        image_stream = io.BytesIO(response.content)
                        
                        # Add the downloaded image to the destination slide
                        dest_slide.shapes.add_picture(image_stream, shape.left, shape.top, width=shape.width, height=shape.height)
                        continue # Skip the generic element copy for this shape
                    except (KeyError, requests.exceptions.RequestException) as e:
                        # If download fails or relationship is not found, we fall back to copying the shape element
                        # This will result in a broken link, but it's better than crashing.
                        st.warning(f"Could not download linked image for shape on slide. Error: {e}")
                        pass
        
        # For all other shapes (or as a fallback for pictures), copy the element
        new_el = copy.deepcopy(shape.element)
        dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')


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

# --- Streamlit App ---
st.set_page_config(page_title="Interactive AI Presentation Assembler", layout="wide")
st.title("🤖 Interactive AI Presentation Assembler")

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

if template_file and gtm_file and api_key and st.session_state.structure:
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

    for i, item in enumerate(st.session_state.build_plan):
        st.subheader(f"Step {i+1}: '{item['keyword']}'")
        # In a real app, you would show thumbnails here. For now, we show justifications.
        if item['action'] == "Copy from GTM (as is)":
            st.info(f"AI Justification (Content): {item['gtm_choice']['justification']}")
        elif item['action'] == "Merge: Template Layout + GTM Content":
            st.info(f"AI Justification (Layout): {item['template_choice']['justification']}")
            st.info(f"AI Justification (Content): {item['gtm_choice']['justification']}")

    st.markdown("---")
    if st.button("2. Assemble Final Presentation", type="primary"):
        with st.spinner("Executing final assembly..."):
            final_prs = Presentation(io.BytesIO(template_file.getvalue()))
            # Clear all slides to start fresh but keep the master styles
            for i in range(len(final_prs.slides) - 1, -1, -1):
                rId = final_prs.slides._sldIdLst[i].rId
                final_prs.part.drop_rel(rId)
                del final_prs.slides._sldIdLst[i]

            for item in st.session_state.build_plan:
                if item['action'] == "Copy from GTM (as is)":
                    if item['gtm_choice']['slide']:
                        new_slide = final_prs.slides.add_slide(item['gtm_choice']['slide'].slide_layout)
                        deep_copy_slide(final_prs, new_slide, item['gtm_choice']['slide'])
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
            
            st.session_state.build_plan = None
            st.session_state.assembly_started = False
