import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE # Import for fill type checking
from pptx.dml.color import RGBColor
import io
import copy
import uuid
import openai
import json
from lxml import etree # Used for direct XML manipulation for backgrounds

# --- Helper Function for Copying Background ---
def copy_slide_background(src_slide, dest_slide):
    """
    Copies the background properties (fill type, color, and image if present)
    from the src_slide to the dest_slide. This involves low-level XML manipulation
    for image backgrounds to ensure correct embedding and relationships.
    """
    src_slide_elm = src_slide.element
    dest_slide_elm = dest_slide.element

    # Find the background properties element in the source slide
    # This element holds the information about the slide's fill (solid, gradient, picture)
    src_bg_pr = src_slide_elm.find('.//p:bgPr', namespaces=src_slide_elm.nsmap)
    
    # If there's no explicit background property on the source slide, there's nothing to copy.
    # It means the background is likely inherited from the slide master/layout.
    if src_bg_pr is None:
        return

    # Check if the source background is a picture fill
    src_blip_fill = src_bg_pr.find('.//a:blipFill', namespaces=src_slide_elm.nsmap)
    
    if src_blip_fill is not None:
        # If it's a picture background, we need to copy the image data and create a new relationship.
        src_blip = src_blip_fill.find('.//a:blip', namespaces=src_slide_elm.nsmap)
        if src_blip is not None and '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed' in src_blip.attrib:
            rId = src_blip.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
            
            try:
                # Get the image part (the actual image data) from the source presentation
                src_image_part = src_slide.part.related_part(rId)
                image_bytes = src_image_part.blob
                
                # Add the image to the destination presentation's media manager.
                # python-pptx handles generating a new rId and saving the image data.
                new_image_part = dest_slide.part.get_or_add_image_part(image_bytes, src_image_part.content_type)
                # Create a new relationship from the destination slide to the newly added image part
                new_rId = dest_slide.part.relate_to(new_image_part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')

                # Deep copy the entire background properties XML element
                new_bg_pr = copy.deepcopy(src_bg_pr)
                # Update the rId in the copied XML to point to the new image part in the destination presentation
                new_blip = new_bg_pr.find('.//a:blip', namespaces=new_bg_pr.nsmap)
                if new_blip is not None:
                    new_blip.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'] = new_rId

                # Remove any existing background properties from the destination slide before appending the new one
                # This ensures we don't end up with multiple conflicting backgrounds
                current_bg = dest_slide_elm.find('.//p:bg', namespaces=dest_slide_elm.nsmap)
                if current_bg is not None:
                    current_bg.getparent().remove(current_bg)
                
                # Append the newly copied and updated background properties to the destination slide's element
                dest_slide_elm.append(new_bg_pr)
                
            except Exception as e:
                print(f"Warning: Could not copy background image. Error: {e}")
                # Fallback to copying solid/gradient background if image copy fails
                copy_solid_or_gradient_background(src_slide, dest_slide)

    else:
        # If not a picture background, it's likely a solid fill, gradient fill, or no background (inherits from master).
        # We can try to copy the existing background properties XML directly.
        copy_solid_or_gradient_background(src_slide, dest_slide)

def copy_solid_or_gradient_background(src_slide, dest_slide):
    """
    Helper to copy solid or gradient background fills using direct XML copy.
    """
    src_slide_elm = src_slide.element
    dest_slide_elm = dest_slide.element
    src_bg_pr = src_slide_elm.find('.//p:bgPr', namespaces=src_slide_elm.nsmap)

    if src_bg_pr is not None:
        new_bg_pr = copy.deepcopy(src_bg_pr) # Deep copy the XML element

        # Remove existing background properties from destination slide (if any)
        current_bg = dest_slide_elm.find('.//p:bg', namespaces=dest_slide_elm.nsmap)
        if current_bg is not None:
            current_bg.getparent().remove(current_bg)
        
        dest_slide_elm.append(new_bg_pr) # Append the copied properties

# --- Core PowerPoint Functions ---

def deep_copy_slide_content(dest_slide, src_slide):
    """
    Performs a stable deep copy of all shapes from a source slide to a
    destination slide, handling different shape types robustly.
    This approach aims to minimize repair issues by using python-pptx's API
    for common shape types, especially images and text.
    It also now attempts to copy the slide's explicit background.
    """
    # Clear all shapes from the destination slide first to prepare it.
    # This loop safely removes shapes by iterating on a copy of the shapes list.
    for shape in list(dest_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)

    for shape in src_slide.shapes:
        left, top, width, height = shape.left, shape.top, shape.width, shape.height

        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            # For pictures, extract the image data and re-add it using python-pptx's API.
            # This is CRUCIAL for avoiding repair issues with images and ensuring proper embedding.
            try:
                image_bytes = shape.image.blob
                dest_slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width, height)
            except Exception as e:
                # Log if an image cannot be copied, but continue with other shapes
                print(f"Warning: Could not copy picture from source slide. Error: {e}")
                # Fallback: if picture has a placeholder, try to copy its XML
                if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                    new_el = copy.deepcopy(shape.element)
                    dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
                
        elif shape.has_text_frame:
            # Create a new text box on the destination slide with the same dimensions
            new_shape = dest_slide.shapes.add_textbox(left, top, width, height)
            new_text_frame = new_shape.text_frame
            new_text_frame.clear() # Clear existing paragraphs to ensure a clean copy

            # Copy text and formatting paragraph by paragraph, run by run
            for paragraph in shape.text_frame.paragraphs:
                new_paragraph = new_text_frame.add_paragraph()
                # Copy paragraph properties (e.g., alignment, indentation)
                new_paragraph.alignment = paragraph.alignment
                if hasattr(paragraph, 'level'): # Bullet level
                    new_paragraph.level = paragraph.level
                
                # Copy runs with their font properties
                for run in paragraph.runs:
                    new_run = new_paragraph.add_run()
                    new_run.text = run.text
                    
                    # Copy essential font properties (bold, italic, underline, size)
                    new_run.font.bold = run.font.bold
                    new_run.font.italic = run.font.italic
                    new_run.font.underline = run.font.underline
                    if run.font.size: # Only copy if size is explicitly defined
                        new_run.font.size = run.font.size
                    
                    # Copy font color if it's a solid fill RGB color
                    if run.font.fill.type == MSO_FILL_TYPE.SOLID: # Use MSO_FILL_TYPE enum
                        new_run.font.fill.solid()
                        try:
                            # Ensure color is an RGBColor object for direct assignment
                            if isinstance(run.font.fill.fore_color.rgb, RGBColor):
                                new_run.font.fill.fore_color.rgb = run.font.fill.fore_color.rgb
                            else: 
                                # Attempt to convert to RGBColor if not already
                                # This handles cases where color might be a theme color or other type
                                rgb_tuple = run.font.fill.fore_color.rgb # Assuming it might be a tuple (R, G, B)
                                new_run.font.fill.fore_color.rgb = RGBColor(rgb_tuple[0], rgb_tuple[1], rgb_tuple[2])
                        except Exception as color_e:
                            print(f"Warning: Could not copy font color. Error: {color_e}")
                            pass # If color conversion fails, skip copying the color

            # Copy text frame properties (word wrap, margins)
            new_text_frame.word_wrap = shape.text_frame.word_wrap
            new_text_frame.margin_left = shape.text_frame.margin_left
            new_text_frame.margin_right = shape.text_frame.margin_right
            new_text_frame.margin_top = shape.text_frame.margin_top
            new_text_frame.margin_bottom = shape.text_frame.margin_bottom

        else:
            # For other shapes (e.g., simple geometric shapes, lines, groups, tables, charts),
            # fall back to deep copying the raw XML element.
            # This is less robust than using specific python-pptx add_* methods but necessary
            # for types not directly supported by add_*.
            # For complex custom shapes, this might still lead to minor issues,
            # but is the best general approach without parsing deeper XML.
            new_el = copy.deepcopy(shape.element)
            dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    
    # --- NEW: Call the background copying function after all shapes are processed ---
    copy_slide_background(src_slide, dest_slide)


def get_all_slide_texts(prs):
    """
    Extracts a summary of text content from all slides in a presentation.
    Used to provide context to the AI about template slide layouts.
    """
    all_slides_text = []
    for i, slide in enumerate(prs.slides):
        slide_text_content = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                slide_text_content.append(shape.text)
        # Concatenate text, limit length to save tokens for AI, but ensure enough context.
        # Increased limit for better semantic understanding.
        all_slides_text.append({"slide_index": i, "text": " ".join(slide_text_content)[:2000]}) 
    return all_slides_text


def find_slide_by_ai(api_key, prs, slide_type_prompt, deck_name):
    """
    Uses OpenAI to intelligently find the best matching slide and get a justification.
    Returns a dictionary with the slide object, its index, and the AI's justification.
    """
    if not slide_type_prompt: return {"slide": None, "index": -1, "justification": "No keyword provided."}
    
    # Check if API key is provided and valid
    if not api_key:
        return {"slide": None, "index": -1, "justification": "OpenAI API Key is missing."}

    client = openai.OpenAI(api_key=api_key)
    
    slides_content = []
    for i, slide in enumerate(prs.slides):
        slide_text = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                slide_text.append(shape.text)
        # Concatenate all text from the slide, limiting to first 1000 characters to save tokens
        slides_content.append({"slide_index": i, "text": " ".join(slide_text)[:1000]})

    # --- UPDATED SYSTEM PROMPT FOR SMARTER TIMELINE DETECTION ---
    system_prompt = f"""
    You are an expert presentation analyst. Your task is to find the best slide in a presentation that matches a user's description.
    The user is looking for a slide representing: '{slide_type_prompt}'.
    
    Analyze the text of each slide to understand its purpose.
    
    **For 'Timeline' slides:** Look for strong indicators of sequential progression. These often include:
    -   Explicit dates, years, quarters (e.g., "Q1 2024", "FY25", "2023-2025").
    -   Phased language (e.g., "Phase 1", "Stage Two", "Initiation", "Completion").
    -   Keywords like "roadmap", "milestones", "schedule", "plan", "future steps", "history".
    -   Text that suggests a visual flow, even if sparse (e.g., short, concise points arranged vertically or horizontally, less dense paragraphs, indicating a graphic is present).
    -   It is NOT just a list in a table of contents. Prioritize slides that imply a visual timeline structure, even if the text itself is minimal, if it contains strong temporal or sequential keywords.

    **For 'Objectives' slides:** These will typically contain goal-oriented language, targets, key results, and strategic aims.

    You must prioritize actual content slides over simple divider or table of contents pages.
    You MUST return a JSON object with two keys: 'best_match_index' (an integer, or -1 if no match) and 'justification' (a brief, one-sentence justification for your choice).
    """
    full_user_prompt = f"Find the best slide for '{slide_type_prompt}' in the '{deck_name}' deck with the following contents:\n{json.dumps(slides_content, indent=2)}"

    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo", 
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": full_user_prompt}],
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        best_index = result.get("best_match_index", -1)
        justification = result.get("justification", "No justification provided.")
        
        # Validate the AI's chosen index
        if best_index != -1 and best_index < len(prs.slides):
            return {"slide": prs.slides[best_index], "index": best_index, "justification": justification}
        else:
            return {"slide": None, "index": -1, "justification": "AI could not find a suitable slide or returned an invalid index."}
    except openai.APIError as e:
        return {"slide": None, "index": -1, "justification": f"OpenAI API Error: {e}"}
    except json.JSONDecodeError as e:
        return {"slide": None, "index": -1, "justification": f"AI response was not valid JSON: {e}"}
    except Exception as e:
        return {"slide": None, "index": -1, "justification": f"An unexpected error occurred during AI analysis: {e}"}


def analyze_and_map_content(api_key, gtm_slide_content, template_slides_summary, user_keyword):
    """
    Uses OpenAI to analyze GTM content, find the best template layout, and
    process the GTM content by inserting regional placeholders.
    """
    if not api_key:
        return {"best_template_index": -1, "justification": "OpenAI API Key is missing.", "processed_content": gtm_slide_content}

    client = openai.OpenAI(api_key=api_key)

    # --- UPDATED SYSTEM PROMPT FOR THOROUGH TEMPLATE SELECTION ---
    system_prompt = f"""
    You are an expert presentation content mapper. Your primary task is to help a user
    integrate content from a Global (GTM) slide into the most appropriate regional template.

    Given the `gtm_slide_content` and a list of `template_slides_summary` (each with an index and text content),
    you must perform two critical tasks:

    1.  **Select the BEST Template:**
        * **Crucially, you must review *each and every* template slide summary provided.**
        * Semantically evaluate which template slide's structure and implied purpose (based on its textual summary) would *best* accommodate the `gtm_slide_content`.
        * **Perform a comparative analysis:** Do not just pick the first decent match. Compare all options to find the single most suitable template.
        * Consider factors like:
            * Does the template layout (implied by its text content) match the theme/type of the GTM content (e.g., if GTM content is about objectives, find an objectives-like template).
            * Is there sufficient space or logical sections in the template for the GTM content?
            * Is the template visually appropriate for the content's nature (e.g., if GTM content is a timeline, is there a template with timeline-like textual elements)?

    2.  **Process GTM Content for Regionalization:**
        * Analyze the `gtm_slide_content` (title and body).
        * Identify any parts of the text that are highly likely to be *regional-specific* (e.g., local market data, specific regional initiatives, detailed local performance figures, regional names, or examples relevant only to one region).
        * For these regional-specific parts, replace them with a concise, generic placeholder like `[REGIONAL DATA HERE]`, `[LOCAL EXAMPLE]`, `[Qx REGIONAL METRICS]`, `[REGIONAL IMPACT]`, `[LOCAL TEAM]`, etc. Be intelligent about the placeholder text.
        * The goal is to provide a global baseline with clear, actionable markers for regional teams to fill in.
        * Maintain the original overall structure, headings, and flow of the text where possible.

    You MUST return a JSON object with the following keys:
    -   `best_template_index`: An integer representing the index of the best template slide from the `template_slides_summary` list.
    -   `justification`: A brief, one-sentence justification for choosing that template, explicitly mentioning why it's better than other contenders if applicable.
    -   `processed_content`: An object with 'title' and 'body' keys, containing the
        GTM content with regional placeholders inserted.
    """
    
    # Prepare the user prompt with the GTM content and template summaries
    full_user_prompt = f"""
    User's original keyword for this content: '{user_keyword}'

    GTM Slide Content to Process:
    {json.dumps(gtm_slide_content, indent=2)}

    Available Template Slides Summary:
    {json.dumps(template_slides_summary, indent=2)}
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo", # Using a capable model for complex reasoning
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": full_user_prompt}],
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        
        # Validate AI response structure
        if "best_template_index" not in result or "justification" not in result or "processed_content" not in result:
            raise ValueError("AI response missing required keys.")

        best_index = result["best_template_index"]
        justification = result["justification"]
        processed_content = result["processed_content"]

        # Ensure processed_content has 'title' and 'body'
        if "title" not in processed_content: processed_content["title"] = gtm_slide_content.get("title", "")
        if "body" not in processed_content: processed_content["body"] = gtm_slide_content.get("body", "")

        return {
            "best_template_index": best_index,
            "justification": justification,
            "processed_content": processed_content
        }

    except openai.APIError as e:
        print(f"OpenAI API Error in analyze_and_map_content: {e}")
        return {"best_template_index": -1, "justification": f"OpenAI API Error: {e}", "processed_content": gtm_slide_content}
    except json.JSONDecodeError as e:
        print(f"JSON Decode Error in analyze_and_map_content: {e}")
        return {"best_template_index": -1, "justification": f"AI response was not valid JSON: {e}", "processed_content": gtm_slide_content}
    except Exception as e:
        print(f"An unexpected error occurred in analyze_and_map_content: {e}")
        return {"best_template_index": -1, "justification": f"An error occurred during content mapping: {e}", "processed_content": gtm_slide_content}


def get_slide_content(slide):
    """Extracts title and body text from a slide."""
    if not slide: return {"title": "", "body": ""}
    
    # Sort text shapes by their top position to infer order (title usually highest)
    text_shapes = sorted([s for s in slide.shapes if s.has_text_frame and s.text.strip()], key=lambda s: s.top)
    
    title = ""
    body = ""
    
    if text_shapes:
        # Heuristic for title: often the first (top-most) text shape.
        # Could be improved by checking placeholder type (e.g., MSO_PLACEHOLDER_TYPE.TITLE)
        title = text_shapes[0].text.strip()
        body = "\n".join(s.text.strip() for s in text_shapes[1:])
        
    return {"title": title, "body": body}

def populate_slide(slide, content):
    """
    Populates a slide's placeholders or main text boxes with new content.
    It clears the existing content and adds new runs, aiming to use existing
    placeholders without forcing bold. This content is expected to already
    contain any necessary regional placeholders.
    """
    title_populated, body_populated = False, False
    
    # Iterate through shapes to find suitable places for title and body
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        
        # Check if it's a title placeholder (type 1, 2, or object type 8 which can be title)
        # Or if it's a top-positioned shape likely to be a title
        is_title_placeholder = (
            hasattr(shape, 'is_placeholder') and shape.is_placeholder and 
            shape.placeholder_format.type in (1, 2, 8) # TITLE, CENTER_TITLE, OBJECT
        )
        is_top_text_box = (shape.top < Pt(150)) # Heuristic: within 1.5 inches from top

        if not title_populated and (is_title_placeholder or is_top_text_box):
            tf = shape.text_frame
            tf.clear() # Clear existing content
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = content.get("title", "")
            # No longer forcing bold here. The template's default formatting will apply.
            title_populated = True
            
        # Check for body placeholders (type 3, 4, 8, 14) or large text boxes with dummy text
        is_body_placeholder = (
            hasattr(shape, 'is_placeholder') and shape.is_placeholder and 
            shape.placeholder_format.type in (3, 4, 8, 14) # BODY, OBJECT, CONTENT_TITLE_BODY
        )
        is_lorem_ipsum = "lorem ipsum" in shape.text.lower()
        is_empty_text_box = not shape.text.strip() and shape.height > Pt(100) # Heuristic for larger empty text boxes

        if not body_populated and (is_body_placeholder or is_lorem_ipsum or is_empty_text_box):
            tf = shape.text_frame
            tf.clear() # Clear existing content
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = content.get("body", "")
            # No longer forcing bold here.
            body_populated = True

        if title_populated and body_populated:
            break # Exit loop once both title and body content are placed

# --- Streamlit App ---
st.set_page_config(page_title="Dynamic AI Presentation Assembler", layout="wide")
st.title("📊 Dynamic AI Presentation Assembler") # Updated emoji for presentation

with st.sidebar:
    st.header("1. API Key & Decks")
    api_key = st.text_input("OpenAI API Key", type="password")
    st.markdown("---")
    st.header("2. Upload Decks")
    template_files = st.file_uploader("Upload Template Deck(s)", type=["pptx"], accept_multiple_files=True)
    gtm_file = st.file_uploader("Upload GTM Global Deck", type=["pptx"])
    st.markdown("---")
    st.header("3. Define Presentation Structure")
    
    # Initialize session state for structure if not present
    if 'structure' not in st.session_state: 
        st.session_state.structure = []
    
    # Button to add a new step to the presentation structure
    if st.button("Add New Step", use_container_width=True):
        st.session_state.structure.append({"id": str(uuid.uuid4()), "keyword": "", "action": "Copy from GTM (as is)"})

    # Display and manage each step in the structure
    for i, step in enumerate(st.session_state.structure):
        with st.container(border=True): # Use a container for visual separation
            cols = st.columns([3, 3, 1]) # Three columns for keyword, action, and delete button
            # Text input for the slide type keyword
            step["keyword"] = cols[0].text_input("Slide Type", step["keyword"], key=f"keyword_{step['id']}")
            # Selectbox for the action to perform (Copy or Merge)
            step["action"] = cols[1].selectbox(
                "Action", 
                ["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"], 
                index=["Copy from GTM (as is)", "Merge: Template Layout + GTM Content"].index(step["action"]), 
                key=f"action_{step['id']}"
            )
            # Delete button for each step
            if cols[2].button("🗑️", key=f"del_{step['id']}"): # Changed emoji for delete
                st.session_state.structure.pop(i) # Remove the step
                st.rerun() # Rerun to update the UI immediately

    # Button to clear all defined steps
    if st.button("Clear Structure", use_container_width=True): 
        st.session_state.structure = []
        st.rerun()

# --- Main App Logic ---
# Check if all necessary inputs are provided before enabling assembly
if template_files and gtm_file and api_key and st.session_state.structure:
    # Button to trigger the presentation assembly process
    if st.button("🚀 Assemble Presentation", type="primary"): # Changed emoji for assemble
        with st.spinner("Assembling your new presentation..."):
            try:
                st.write("Step 1/3: Loading decks...")
                # CRITICAL: Use the first uploaded template file as the base for the new presentation.
                new_prs = Presentation(io.BytesIO(template_files[0].getvalue()))
                gtm_prs = Presentation(io.BytesIO(gtm_file.getvalue()))
                
                process_log = [] # To store logs of what happened during assembly
                st.write("Step 2/3: Building new presentation based on your structure...")
                
                num_template_slides = len(new_prs.slides)
                num_structure_steps = len(st.session_state.structure)

                # Prune excess slides from the template if the defined structure is shorter
                if num_structure_steps < num_template_slides:
                    # Iterate backwards to safely delete slides
                    for i in range(num_template_slides - 1, num_structure_steps - 1, -1):
                        rId = new_prs.slides._sldIdLst[i].rId # Get relationship ID
                        new_prs.part.drop_rel(rId) # Drop relationship
                        del new_prs.slides._sldIdLst[i] # Delete slide from slide list
                    st.info(f"Removed {num_template_slides - num_structure_steps} unused slides from the template.")
                elif num_structure_steps > num_template_slides:
                     st.warning(f"Warning: Your defined structure has more steps ({num_structure_steps}) than the template has slides ({num_template_slides}). Extra steps will be ignored.")

                # Process slides based on the defined structure
                for i, step in enumerate(st.session_state.structure):
                    # Ensure we don't go out of bounds if the template was trimmed or structure is longer
                    if i >= len(new_prs.slides): 
                        break

                    # dest_slide is now potentially changed if a new template slide is chosen by AI
                    current_dest_slide_index = i # Initial index in the new presentation
                    dest_slide = new_prs.slides[current_dest_slide_index] 
                    
                    keyword, action = step["keyword"], step["action"]
                    log_entry = {"step": i + 1, "keyword": keyword, "action": action, "log": []}
                    
                    if action == "Copy from GTM (as is)":
                        # Find the best matching slide in the GTM deck using AI
                        result = find_slide_by_ai(api_key, gtm_prs, keyword, "GTM Deck")
                        log_entry["log"].append(f"**GTM Content Choice Justification:** {result['justification']}")
                        if result["slide"]:
                            # If a suitable slide is found, deep copy its content to the destination slide
                            deep_copy_slide_content(dest_slide, result["slide"])
                            log_entry["log"].append(f"**Action:** Replaced Template slide {current_dest_slide_index + 1} with content from GTM slide {result['index'] + 1}.")
                        else:
                            log_entry["log"].append("**Action:** No suitable slide found in GTM deck. Template slide was left as is.")
                    
                    elif action == "Merge: Template Layout + GTM Content":
                        # --- New Logic for Merge Action ---
                        # 1. Find the best matching slide in the GTM deck (source of content)
                        gtm_content_source_result = find_slide_by_ai(api_key, gtm_prs, keyword, "GTM Deck (Content Source)")
                        log_entry["log"].append(f"**GTM Content Source Justification:** {gtm_content_source_result['justification']}")

                        if gtm_content_source_result["slide"]:
                            # Extract raw content from the identified GTM source slide
                            raw_gtm_content = get_slide_content(gtm_content_source_result["slide"])
                            
                            # Get summaries of all template slides for AI context
                            template_slides_summary = get_all_slide_texts(new_prs)

                            # 2. Use AI to analyze GTM content, find best template layout, and generate placeholders
                            ai_mapping_result = analyze_and_map_content(
                                api_key, 
                                raw_gtm_content, 
                                template_slides_summary, 
                                keyword # Pass original keyword for AI context
                            )
                            log_entry["log"].append(f"**AI Template Mapping Justification:** {ai_mapping_result['justification']}")

                            selected_template_index = ai_mapping_result["best_template_index"]
                            processed_content = ai_mapping_result["processed_content"]

                            if selected_template_index != -1 and selected_template_index < len(new_prs.slides):
                                # The AI selected a preferred template index. While we still populate the slide at `i` in the output,
                                # the AI's selection process for `best_template_index` is now more robust.
                                # The `populate_slide` function is designed to work with the *current* `dest_slide` (new_prs.slides[i]).
                                # The AI's `best_template_index` helps ensure the *content* is processed to best fit *a* template type.

                                # Populate the current destination slide with the processed content
                                populate_slide(dest_slide, processed_content)
                                log_entry["log"].append(f"**Action:** Merged processed content from GTM slide {gtm_content_source_result['index'] + 1} into Template slide {current_dest_slide_index + 1}, with regional placeholders. AI suggested template type at index {selected_template_index + 1}.")
                            else:
                                log_entry["log"].append("**Action:** AI could not determine a suitable template layout or process content. Template slide was left as is.")
                        else:
                            log_entry["log"].append("**Action:** No suitable content found in GTM deck. Template slide was left as is.")
                    
                    process_log.append(log_entry) # Add step log to overall process log
 
                st.success("Successfully built the new presentation structure.")
                
                st.write("Step 3/3: Finalizing...")
                st.subheader("📋 Process Log") # Changed emoji for process log
                # Display the process log in an expandable format
                for entry in process_log:
                    with st.expander(f"Step {entry['step']}: '{entry['keyword']}' ({entry['action']})"):
                        for line in entry['log']: 
                            st.markdown(f"- {line}")
                
                # Save the assembled presentation to an in-memory buffer
                output_buffer = io.BytesIO()
                new_prs.save(output_buffer)
                output_buffer.seek(0) # Rewind the buffer to the beginning for downloading

                st.success("✨ Your new regional presentation has been assembled!") # Changed emoji for success
                # Provide a download button for the user
                st.download_button(
                    "Download Assembled PowerPoint", 
                    data=output_buffer, 
                    file_name="Dynamic_AI_Assembled_Deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"A critical error occurred: {e}")
                st.exception(e) # Display full traceback for debugging
else:
    # Instructions displayed when inputs are not yet complete
    st.info("Please provide an API Key, upload at least one Template Deck and a GTM Deck, and define the structure in the sidebar to begin.")

