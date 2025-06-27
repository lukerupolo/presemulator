import streamlit as st
from pptx import Presentation # Still used for generating the output PPTX
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.dml.color import RGBColor
import io
import copy
import uuid
import openai
import json
import base64 # Required for encoding/decoding image data

# --- Placeholder for Visual Data ---
# IMPORTANT: This is a placeholder. In a real deployment, you would need
# external tools (like LibreOffice/unoconv for PPTX, or PyMuPDF/poppler-utils for PDF)
# to convert PPTX slides or PDF pages into actual Base64-encoded images.
# This dummy image allows the multimodal AI API calls to be structured correctly,
# but the AI will not receive real visual information from your uploaded files
# in this sandbox environment.
Base64_PLACEHOLDER_IMAGE = "R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==" # 1x1 transparent GIF

# --- Helper Function for Copying Background (PPTX-specific) ---
def copy_slide_background(src_slide, dest_slide):
    """
    Copies the background properties (fill type, color, and image if present)
    from the src_slide to the dest_slide. This involves low-level XML manipulation
    for image backgrounds to ensure correct embedding and relationships.
    """
    src_slide_elm = src_slide.element
    dest_slide_elm = dest_slide.element

    src_bg_pr = src_slide_elm.find('.//p:bgPr', namespaces=src_slide_elm.nsmap)
    
    if src_bg_pr is None:
        return

    src_blip_fill = src_bg_pr.find('.//a:blipFill', namespaces=src_slide_elm.nsmap)
    
    if src_blip_fill is not None:
        src_blip = src_blip_fill.find('.//a:blip', namespaces=src_slide_elm.nsmap)
        if src_blip is not None and '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed' in src_blip.attrib:
            rId = src_blip.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
            
            try:
                src_image_part = src_slide.part.related_part(rId)
                image_bytes = src_image_part.blob
                
                new_image_part = dest_slide.part.get_or_add_image_part(image_bytes, src_image_part.content_type)
                new_rId = dest_slide.part.relate_to(new_image_part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')

                new_bg_pr = copy.deepcopy(src_bg_pr)
                new_blip = new_bg_pr.find('.//a:blip', namespaces=new_bg_pr.nsmap)
                if new_blip is not None:
                    new_blip.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'] = new_rId

                current_bg = dest_slide_elm.find('.//p:bg', namespaces=dest_slide_elm.nsmap)
                if current_bg is not None:
                    current_bg.getparent().remove(current_bg)
                
                dest_slide_elm.append(new_bg_pr)
                
            except Exception as e:
                print(f"Warning: Could not copy background image. Error: {e}")
                copy_solid_or_gradient_background(src_slide, dest_slide)

    else:
        copy_solid_or_gradient_background(src_slide, dest_slide)

def copy_solid_or_gradient_background(src_slide, dest_slide):
    """
    Helper to copy solid or gradient background fills using direct XML copy.
    """
    src_slide_elm = src_slide.element
    dest_slide_elm = dest_slide.element
    src_bg_pr = src_slide_elm.find('.//p:bgPr', namespaces=src_slide_elm.nsmap)

    if src_bg_pr is not None:
        new_bg_pr = copy.deepcopy(src_bg_pr)

        current_bg = dest_slide_elm.find('.//p:bg', namespaces=dest_slide_elm.nsmap)
        if current_bg is not None:
            current_bg.getparent().remove(current_bg)
        
        dest_slide_elm.append(new_bg_pr)

# --- Core PowerPoint Functions (for PPTX output generation) ---
def deep_copy_slide_content(dest_slide, src_slide):
    """
    Performs a stable deep copy of all shapes from a source PPTX slide to a
    destination PPTX slide, handling different shape types robustly.
    """
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)

    for shape in src_slide.shapes:
        left, top, width, height = shape.left, shape.top, shape.width, shape.height

        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                image_bytes = shape.image.blob
                dest_slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width, height)
            except Exception as e:
                print(f"Warning: Could not copy picture from source slide. Error: {e}")
                if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                    new_el = copy.deepcopy(shape.element)
                    dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
                
        elif shape.has_text_frame:
            new_shape = dest_slide.shapes.add_textbox(left, top, width, height)
            new_text_frame = new_shape.text_frame
            new_text_frame.clear()

            for paragraph in shape.text_frame.paragraphs:
                new_paragraph = new_text_frame.add_paragraph()
                new_paragraph.alignment = paragraph.alignment
                if hasattr(paragraph, 'level'):
                    new_paragraph.level = paragraph.level
                
                for run in paragraph.runs:
                    new_run = new_paragraph.add_run()
                    new_run.text = run.text
                    
                    new_run.font.bold = run.font.bold
                    new_run.font.italic = run.font.italic
                    new_run.font.underline = run.font.underline
                    if run.font.size:
                        new_run.font.size = run.font.size
                    
                    if run.font.fill.type == MSO_FILL_TYPE.SOLID:
                        new_run.font.fill.solid()
                        try:
                            if isinstance(run.font.fill.fore_color.rgb, RGBColor):
                                new_run.font.fill.fore_color.rgb = run.font.fill.fore_color.rgb
                            else: 
                                rgb_tuple = run.font.fill.fore_color.rgb
                                new_run.font.fill.fore_color.rgb = RGBColor(rgb_tuple[0], rgb_tuple[1], rgb_tuple[2])
                        except Exception as color_e:
                            print(f"Warning: Could not copy font color. Error: {color_e}")
                            pass

            new_text_frame.word_wrap = shape.text_frame.word_wrap
            new_text_frame.margin_left = shape.text_frame.margin_left
            new_text_frame.margin_right = shape.text_frame.margin_right
            new_text_frame.margin_top = shape.text_frame.margin_top
            new_text_frame.margin_bottom = shape.text_frame.margin_bottom

        else:
            new_el = copy.deepcopy(shape.element)
            dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    
    copy_slide_background(src_slide, dest_slide)


def get_all_slide_data(file_bytes: bytes, file_type: str):
    """
    Extracts text content and includes a placeholder for visual data from all
    slides/pages of a given file (PPTX or PDF).
    """
    all_slides_data = []

    if file_type == 'pptx':
        prs = Presentation(io.BytesIO(file_bytes))
        for i, slide in enumerate(prs.slides):
            slide_text_content = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    slide_text_content.append(shape.text)
            
            all_slides_data.append({
                "slide_index": i, 
                "text": " ".join(slide_text_content)[:2000], # Limit text length for AI tokens
                "image_data": Base64_PLACEHOLDER_IMAGE # Placeholder for visual data
            })
    elif file_type == 'pdf':
        # IMPORTANT: In a real application, you would use a library like PyMuPDF (fitz)
        # to actually extract text AND render images from PDF pages.
        # This current implementation provides dummy text and a placeholder image
        # for demonstration purposes only.
        
        # Simulate a few pages and dummy text/image
        simulated_page_count = 5 
        
        for i in range(simulated_page_count):
            all_slides_data.append({
                "slide_index": i,
                "text": f"Simulated text from PDF page {i+1}. This would be actual text extracted from the PDF page. It contains global sales figures for Q1 2024 and marketing initiatives for EMEA region. Also, key milestones for product launch in Q3.",
                "image_data": Base64_PLACEHOLDER_IMAGE # Placeholder for visual data
            })
    
    return all_slides_data


def find_slide_by_ai(api_key, file_bytes: bytes, file_type: str, slide_type_prompt: str, deck_name: str):
    """
    Uses a multimodal AI (gpt-4o) to intelligently find the best matching slide/page
    based on combined text and (placeholder) visual content.
    """
    if not slide_type_prompt: return {"slide": None, "index": -1, "justification": "No keyword provided."}
    
    if not api_key:
        return {"slide": None, "index": -1, "justification": "OpenAI API Key is missing."}

    client = openai.OpenAI(api_key=api_key)
    
    slides_data = get_all_slide_data(file_bytes, file_type) # Get text AND placeholder image data

    system_prompt = f"""
    You are an expert presentation analyst. Your task is to find the best slide/page in a document that matches a user's description.
    The user is looking for a slide/page representing: '{slide_type_prompt}'.
    
    Analyze both the provided **text content** and the **visual structure (from the image)** for each slide/page to infer its purpose.
    
    **For 'Timeline' slides/pages:** Look for strong textual indicators of sequential progression (dates, years, quarters, phased language like "Phase 1", "roadmap", "milestones"). **Crucially, also use visual patterns from the image, such as horizontal or vertical arrangements of distinct elements, flow arrows, or clear segmentation over time.** Prioritize slides/pages that combine strong textual cues with implied or explicit visual timeline structures.

    **For 'Objectives' slides/pages:** These will typically contain goal-oriented language, targets, key results, and strategic aims in both text and potentially visually organized lists or impact statements.

    You must prioritize actual content slides/pages over simple divider or table of contents pages.
    Return a JSON object with 'best_match_index' (integer, or -1) and 'justification' (brief, one-sentence).
    """

    user_parts = [
        {"type": "text", "text": f"Find the best slide/page for '{slide_type_prompt}' in the '{deck_name}' with the following pages/slides:"}
    ]
    for slide_info in slides_data:
        user_parts.append({"type": "text", "text": f"\n--- Page/Slide {slide_info['slide_index'] + 1} (Text): {slide_info['text']}"})
        # Include the placeholder image for the AI to "see"
        user_parts.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{slide_info['image_data']}"
            }
        })
    
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_parts}
    ]

    try:
        response = openai.OpenAI(api_key=api_key).chat.completions.create( # Use client directly
            model="gpt-4o", # Changed model to gpt-4o for multimodal
            messages=messages,
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        best_index = result.get("best_match_index", -1)
        justification = result.get("justification", "No justification provided.")
        
        return {"slide": slides_data[best_index] if best_index != -1 else None, 
                "index": best_index, 
                "justification": justification}
    except openai.APIError as e:
        return {"slide": None, "index": -1, "justification": f"OpenAI API Error: {e}"}
    except json.JSONDecodeError as e:
        return {"slide": None, "index": -1, "justification": f"AI response was not valid JSON: {e}"}
    except Exception as e:
        return {"slide": None, "index": -1, "justification": f"An unexpected error occurred during AI analysis: {e}"}


def analyze_and_map_content(api_key, gtm_slide_content_data, template_slides_data, user_keyword):
    """
    Uses a multimodal AI (gpt-4o) to analyze GTM content (text + visual), find the best
    template layout (text + visual), and process the GTM content by inserting regional placeholders.
    """
    if not api_key:
        return {"best_template_index": -1, "justification": "OpenAI API Key is missing.", "processed_content": gtm_slide_content_data}

    client = openai.OpenAI(api_key=api_key)

    system_prompt = f"""
    You are an expert presentation content mapper. Your primary task is to help a user
    integrate content from a Global (GTM) slide/page into the most appropriate regional template.

    Given the `gtm_slide_content` (with its text and image) and a list of `template_slides_data`
    (each with an index, text content, and image data), you must perform two critical tasks:

    1.  **Select the BEST Template:**
        * **Crucially, you must review *each and every* template slide/page text summary AND its associated visual cue provided.**
        * Semantically and **visually** evaluate which template slide's structure and implied purpose would *best* accommodate the `gtm_slide_content`.
        * **Perform a comparative analysis:** Do not just pick the first decent match. Compare all options to find the single most suitable template based on a combined understanding of text and visuals.
        * Consider factors like:
            * Does the template's textual layout (e.g., presence of sections, bullet points, titles) **and its visual layout (e.g., number of content blocks, placement of image placeholders, overall design)** match the theme/type of the GTM content.
            * Is there sufficient space or logical sections in the template for the GTM content based on its textual and visual structure?
            * Is the template visually appropriate for the content's nature (e.g., if GTM content is a timeline, does the template's visual suggest a timeline-like structure with distinct steps)?

    2.  **Process GTM Content for Regionalization:**
        * Analyze the `gtm_slide_content` (title and body text).
        * Identify any parts of the text that are highly likely to be *regional-specific* (e.g., local market data, specific regional initiatives, detailed local performance figures, regional names, or examples relevant only to one region).
        * For these regional-specific parts, replace them with a concise, generic placeholder like `[REGIONAL DATA HERE]`, `[LOCAL EXAMPLE]`, `[Qx REGIONAL METRICS]`, `[REGIONAL IMPACT]`, `[LOCAL TEAM]`, etc. Be intelligent about the placeholder text.
        * The goal is to provide a global baseline with clear, actionable markers for regional teams to fill in.
        * Maintain the original overall structure, headings, and flow of the text where possible.

    You MUST return a JSON object with the following keys:
    -   `best_template_index`: An integer representing the index of the best template slide/page from the `template_slides_data` list.
    -   `justification`: A brief, one-sentence justification for choosing that template, explicitly mentioning why it's better than other contenders if applicable.
    -   `processed_content`: An object with 'title' and 'body' keys, containing the
        GTM content with regional placeholders inserted.
    """}]
    
    user_parts = [
        {"type": "text", "text": f"User's original keyword for this content: '{user_keyword}'"},
        {"type": "text", "text": "GTM Slide/Page Content to Process (Text):"},
        {"type": "text", "text": json.dumps(gtm_slide_content_data.get('text', {}), indent=2)},
        # Include GTM slide's placeholder image
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{gtm_slide_content_data.get('image_data', Base64_PLACEHOLDER_IMAGE)}"}}
    ]

    user_parts.append({"type": "text", "text": "\nAvailable Template Slides/Pages Summary and Visuals:"})
    for slide_info in template_slides_data:
        user_parts.append({"type": "text", "text": f"\n--- Template Slide/Page {slide_info['slide_index'] + 1} (Text): {slide_info['text']}"})
        # Include each template slide's placeholder image
        user_parts.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{slide_info['image_data']}"
            }
        })
    
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_parts}
    ]

    try:
        response = client.chat.completions.create(
            model="gpt-4o", # Changed model to gpt-4o for multimodal
            messages=messages,
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        
        if "best_template_index" not in result or "justification" not in result or "processed_content" not in result:
            raise ValueError("AI response missing required keys.")

        best_index = result["best_template_index"]
        justification = result["justification"]
        processed_content = result["processed_content"]

        if "title" not in processed_content: processed_content["title"] = gtm_slide_content_data.get("title", "")
        if "body" not in processed_content: processed_content["body"] = gtm_slide_content_data.get("body", "")

        return {
            "best_template_index": best_index,
            "justification": justification,
            "processed_content": processed_content
        }

    except openai.APIError as e:
        print(f"OpenAI API Error in analyze_and_map_content: {e}")
        return {"best_template_index": -1, "justification": f"OpenAI API Error: {e}", "processed_content": gtm_slide_content_data}
    except json.JSONDecodeError as e:
        print(f"JSON Decode Error in analyze_and_map_content: {e}")
        return {"best_template_index": -1, "justification": f"AI response was not valid JSON: {e}", "processed_content": gtm_slide_content_data}
    except Exception as e:
        print(f"An unexpected error occurred in analyze_and_map_content: {e}")
        return {"best_template_index": -1, "justification": f"An error occurred during content mapping: {e}", "processed_content": gtm_slide_content_data}


def get_slide_content(slide):
    """
    Extracts title and body text from a PPTX slide object.
    This is used for the *destination* PPTX when populating.
    """
    if not slide: return {"title": "", "body": ""}
    
    text_shapes = sorted([s for s in slide.shapes if s.has_text_frame and s.text.strip()], key=lambda s: s.top)
    
    title = ""
    body = ""
    
    if text_shapes:
        title = text_shapes[0].text.strip()
        body = "\n".join(s.text.strip() for s in text_shapes[1:])
        
    return {"title": title, "body": body}

def populate_slide(slide, content):
    """
    Populates a PPTX slide's placeholders or main text boxes with new content.
    This content is expected to already contain any necessary regional placeholders.
    """
    title_populated, body_populated = False, False
    
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        
        is_title_placeholder = (
            hasattr(shape, 'is_placeholder') and shape.is_placeholder and 
            shape.placeholder_format.type in (1, 2, 8)
        )
        is_top_text_box = (shape.top < Pt(150))

        if not title_populated and (is_title_placeholder or is_top_text_box):
            tf = shape.text_frame
            tf.clear()
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = content.get("title", "")
            title_populated = True
            
        is_body_placeholder = (
            hasattr(shape, 'is_placeholder') and shape.is_placeholder and 
            shape.placeholder_format.type in (3, 4, 8, 14)
        )
        is_lorem_ipsum = "lorem ipsum" in shape.text.lower()
        is_empty_text_box = not shape.text.strip() and shape.height > Pt(100)

        if not body_populated and (is_body_placeholder or is_lorem_ipsum or is_empty_text_box):
            tf = shape.text_frame
            tf.clear()
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = content.get("body", "")
            body_populated = True

        if title_populated and body_populated:
            break

# --- Streamlit App ---
st.set_page_config(page_title="Dynamic AI Presentation Assembler", layout="wide")
st.title("📊 Dynamic AI Presentation Assembler")

with st.sidebar:
    st.header("1. API Key & Decks")
    api_key = st.text_input("OpenAI API Key", type="password")
    st.markdown("---")
    st.header("2. Upload Decks")
    # Template deck is PPTX for output structure
    template_files = st.file_uploader("Upload Template Deck(s) (PPTX)", type=["pptx"], accept_multiple_files=True)
    # GTM Global deck is PDF for AI analysis (text + visual placeholder here)
    gtm_file = st.file_uploader("Upload GTM Global Deck (PDF)", type=["pdf"])
    st.markdown("---")
    st.header("3. Define Presentation Structure")
    
    if 'structure' not in st.session_state: 
        st.session_state.structure = []
    
    if st.button("Add New Step", use_container_width=True):
        st.session_state.structure.append({"id": str(uuid.uuid4()), "keyword": "", "action": "Copy from GTM (as is)"})

    for i, step in enumerate(st.session_state.structure):
        with st.container(border=True):
            cols = st.columns([3, 3, 1])
            step["keyword"] = cols[0].text_input("Slide Type", step["keyword"], key=f"keyword_{step['id']}")
            step["action"] = cols[1].selectbox(
                "Action", 
                ["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"], 
                index=["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"].index(step["action"]), 
                key=f"action_{step['id']}"
            )
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
                st.write("Step 1/3: Loading decks...")
                new_prs = Presentation(io.BytesIO(template_files[0].getvalue()))
                gtm_file_bytes = gtm_file.getvalue() # Get PDF bytes for GTM deck

                process_log = []
                st.write("Step 2/3: Building new presentation based on your structure...")
                
                num_template_slides = len(new_prs.slides)
                num_structure_steps = len(st.session_state.structure)

                if num_structure_steps < num_template_slides:
                    for i in range(num_template_slides - 1, num_structure_steps - 1, -1):
                        rId = new_prs.slides._sldIdLst[i].rId
                        new_prs.part.drop_rel(rId)
                        del new_prs.slides._sldIdLst[i]
                    st.info(f"Removed {num_template_slides - num_structure_steps} unused slides from the template.")
                elif num_structure_steps > num_template_slides:
                     st.warning(f"Warning: Your defined structure has more steps ({num_structure_steps}) than the template has slides ({num_template_slides}). Extra steps will be ignored.")

                for i, step in enumerate(st.session_state.structure):
                    if i >= len(new_prs.slides): 
                        break

                    current_dest_slide_index = i
                    dest_slide = new_prs.slides[current_dest_slide_index] 
                    
                    keyword, action = step["keyword"], step["action"]
                    log_entry = {"step": i + 1, "keyword": keyword, "action": action, "log": []}
                    
                    if action == "Copy from GTM (as is)":
                        log_entry["log"].append(f"**Warning:** 'Copy from GTM (as is)' is selected but GTM deck is a PDF. This action cannot directly copy PPTX shapes from a PDF. Proceeding with 'Merge' logic for content extraction based on text and placeholder visuals.")
                        action = "Merge: Template Layout + GTM Content" 
                    
                    if action == "Merge: Template Layout + GTM Content":
                        # For PDF GTM, get content and placeholder visual for AI analysis
                        gtm_ai_selection_result = find_slide_by_ai(api_key, gtm_file_bytes, 'pdf', keyword, "GTM Deck (Content Source)")
                        log_entry["log"].append(f"**GTM Content Source Justification:** {gtm_ai_selection_result['justification']}")
                        
                        raw_gtm_content = {"title": "", "body": "", "image_data": Base64_PLACEHOLDER_IMAGE}
                        if gtm_ai_selection_result["slide"]:
                            full_text = gtm_ai_selection_result["slide"].get("text", "")
                            lines = full_text.split('\n')
                            raw_gtm_content["title"] = lines[0] if lines else ""
                            raw_gtm_content["body"] = "\n".join(lines[1:]) if len(lines) > 1 else ""
                            raw_gtm_content["image_data"] = gtm_ai_selection_result["slide"].get("image_data", Base64_PLACEHOLDER_IMAGE)

                        template_file_bytes = template_files[0].getvalue() 
                        template_slides_data = get_all_slide_data(template_file_bytes, 'pptx') # Get template text + placeholder visual data

                        ai_mapping_result = analyze_and_map_content(
                            api_key, 
                            raw_gtm_content, # Pass text + placeholder visual content
                            template_slides_data, # Pass text + placeholder visual template data
                            keyword
                        )
                        log_entry["log"].append(f"**AI Template Mapping Justification:** {ai_mapping_result['justification']}")

                        selected_template_index = ai_mapping_result["best_template_index"]
                        processed_content = ai_mapping_result["processed_content"]

                        if selected_template_index != -1 and selected_template_index < len(new_prs.slides):
                            populate_slide(dest_slide, processed_content)
                            log_entry["log"].append(f"**Action:** Merged processed content from GTM (PDF) page {gtm_ai_selection_result['index'] + 1} into Template slide {current_dest_slide_index + 1}, with regional placeholders. AI suggested template type at index {selected_template_index + 1}.")
                        else:
                            log_entry["log"].append("**Action:** AI could not determine a suitable template layout or process content. Template slide was left as is.")
                    
                    process_log.append(log_entry)
 
                st.success("Successfully built the new presentation structure.")
                
                st.write("Step 3/3: Finalizing...")
                st.subheader("📋 Process Log")
                for entry in process_log:
                    with st.expander(f"Step {entry['step']}: '{entry['keyword']}' ({entry['action']})"):
                        for line in entry['log']: 
                            st.markdown(f"- {line}")
                
                output_buffer = io.BytesIO()
                new_prs.save(output_buffer)
                output_buffer.seek(0)

                st.success("✨ Your new regional presentation has been assembled!")
                st.download_button(
                    "Download Assembled PowerPoint", 
                    data=output_buffer, 
                    file_name="Dynamic_AI_Assembled_Deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"A critical error occurred: {e}")
                st.exception(e)
else:
    st.info("Please provide an API Key, upload at least one Template Deck (PPTX) and a GTM Global Deck (PDF), and define the structure in the sidebar to begin.")

