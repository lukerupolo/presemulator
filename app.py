import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.dml.color import RGBColor
import io
import copy
import uuid
import openai
import json
import requests
import os
import subprocess # For running git commands
import tempfile   # For temporary directories
import shutil     # For cleaning up directories
import mimetypes  # For determining file types

# --- Configuration for the Conversion Service ---
CONVERSION_SERVICE_URL = os.getenv("CONVERSION_SERVICE_URL", "http://localhost:8000/convert_document")
# Get GitHub Token from environment variable. Set this securely in your deployment.
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

# --- Helper Function for Cloning and Downloading Files from GitHub ---
@st.cache_data(show_spinner="Cloning GitHub repository and finding files...")
def _download_files_from_github(repo_url: str, branch: str, file_paths_or_patterns: list[str]) -> list[tuple[bytes, str, str]]:
    """
    Clones a GitHub repository (using PAT if provided) and extracts file bytes.
    Supports basic file paths and directory patterns.
    Returns a list of (file_bytes, mime_type, file_name).
    """
    downloaded_files_data = []
    
    # Construct authenticated URL if token is available
    if GITHUB_TOKEN:
        # Example: https://github.com/org/repo.git -> https://oauth2:YOUR_PAT@github.com/org/repo.git
        # Or using https://user:PAT@github.com/org/repo.git style for older git versions
        # Modern git handles the token better as a credential helper, but for subprocess this is safer
        parsed_url = repo_url.replace("https://github.com/", f"https://oauth2:{GITHUB_TOKEN}@github.com/")
    else:
        parsed_url = repo_url

    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        repo_name = repo_url.split('/')[-1].replace(".git", "") # Extract repo name from URL
        repo_path = os.path.join(temp_dir, repo_name)

        # Clone the repository
        st.info(f"Cloning '{repo_url}' (branch: {branch or 'default'})...")
        # --depth 1 for a shallow clone (only latest commit) to save time/space
        clone_command = ["git", "clone", "--depth", "1"] 
        if branch:
            clone_command.extend(["--branch", branch])
        clone_command.append(parsed_url)
        clone_command.append(repo_path) # Clone into the named directory inside temp_dir

        # Run git clone command
        result = subprocess.run(clone_command, capture_output=True, text=True, check=False)
        if result.returncode != 0:
            st.error(f"Failed to clone repository: {result.stderr}")
            raise Exception(f"Git clone failed: {result.stderr}")
        
        st.info("Repository cloned. Searching for specified files/folders...")

        # Find and read files based on provided paths/patterns
        for path_or_pattern in file_paths_or_patterns:
            full_path_in_repo = os.path.join(repo_path, path_or_pattern)
            
            if os.path.isdir(full_path_in_repo):
                # If it's a directory, walk through it to find all files
                for root, _, files in os.walk(full_path_in_repo):
                    for file_name in files:
                        file_abs_path = os.path.join(root, file_name)
                        try:
                            with open(file_abs_path, "rb") as f:
                                file_bytes = f.read()
                            # Guess MIME type, with fallbacks for common types
                            mime_type, _ = mimetypes.guess_type(file_name)
                            if not mime_type: 
                                if file_name.lower().endswith(".pptx"): mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                                elif file_name.lower().endswith(".pdf"): mime_type = "application/pdf"
                                else: mime_type = "application/octet-stream" # Default binary type
                            
                            downloaded_files_data.append((file_bytes, mime_type, file_name))
                            st.success(f"Found and loaded: {file_name} (Type: {mime_type})")
                        except Exception as e:
                            st.warning(f"Could not read file {file_name} in '{root}': {e}")
            elif os.path.isfile(full_path_in_repo):
                # If it's a single file
                try:
                    with open(full_path_in_repo, "rb") as f:
                        file_bytes = f.read()
                    mime_type, _ = mimetypes.guess_type(full_path_in_repo)
                    if not mime_type: 
                        if full_path_in_repo.lower().endswith(".pptx"): mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        elif full_path_in_repo.lower().endswith(".pdf"): mime_type = "application/pdf"
                        else: mime_type = "application/octet-stream"

                    downloaded_files_data.append((file_bytes, mime_type, os.path.basename(full_path_in_repo)))
                    st.success(f"Found and loaded: {os.path.basename(full_path_in_repo)} (Type: {mime_type})")
                except Exception as e:
                    st.warning(f"Could not read single file {full_path_in_repo}: {e}")
            else:
                st.warning(f"Path not found or not a recognized file/directory: '{path_or_pattern}' in repo.")

        if not downloaded_files_data:
            st.error(f"No files found matching patterns {file_paths_or_patterns} in '{repo_url}'. Please check the paths and verify they exist in the repository's '{branch}' branch.")

    except Exception as e:
        st.error(f"An error occurred during GitHub operation: {e}")
        return []
    finally:
        # Clean up the temporary directory where the repo was cloned
        if temp_dir and os.path.exists(temp_dir):
            st.info(f"Cleaning up temporary directory: {temp_dir}")
            shutil.rmtree(temp_dir) 

    return downloaded_files_data


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
        src_blip = src_blip_fill.find('.//a:blip', namespaces=src_blip_fill.nsmap)
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


# MODIFIED: This function now calls the external conversion service
def get_all_slide_data(file_bytes: bytes, file_type: str) -> list[dict]:
    """
    Sends the document file to the external conversion service to get
    extracted text and Base64-encoded image data for each slide/page.
    """
    files = {'file': (f"document.{file_type.split('/')[-1]}", file_bytes, file_type)}
    
    try:
        response = requests.post(CONVERSION_SERVICE_URL, files=files, timeout=300) # Added timeout
        response.raise_for_status() # Raise an exception for HTTP errors (4xx or 5xx)
        return response.json()['slides']
    except requests.exceptions.RequestException as e:
        st.error(f"Error connecting to conversion service or during conversion: {e}. Please ensure the conversion service is running at {CONVERSION_SERVICE_URL} and has required system dependencies (LibreOffice, PyMuPDF).")
        st.stop()
    except KeyError:
        st.error("Conversion service returned an unexpected response format.")
        st.stop()


def find_slide_by_ai(api_key, file_bytes: bytes, file_type: str, slide_type_prompt: str, deck_name: str):
    """
    Uses a multimodal AI (gpt-4o) to intelligently find the best matching slide/page
    based on combined text and actual visual content provided by the conversion service.
    """
    if not slide_type_prompt: return {"slide": None, "index": -1, "justification": "No keyword provided."}
    
    if not api_key:
        return {"slide": None, "index": -1, "justification": "OpenAI API Key is missing."}

    client = openai.OpenAI(api_key=api_key)
    
    slides_data = get_all_slide_data(file_bytes, file_type) # Get actual text and image data from conversion service

    system_prompt = f"""
    You are an expert presentation analyst. Your task is to find the best slide/page in a document that matches a user's description.
    The user is looking for a slide/page representing: '{slide_type_prompt}'.
    
    Analyze both the provided **text content** and the **visual structure (from the image)** for each slide/page to infer its purpose.
    
    **For 'Timeline' slides/pages:** Look for strong textual indicators of sequential progression (dates, years, quarters, phased language like "Phase 1", "roadmap", "milestones"). **Crucially, also use visual patterns from the image, such as horizontal or vertical arrangements of distinct elements, flow arrows, or clear segmentation over time. Do NOT rely solely on explicit textual labels (e.g., 'Timeline slide'). Focus on patterns that *imply* a visual timeline.** Prioritize slides/pages that combine strong textual cues with implied or explicit visual timeline structures.

    **For 'Objectives' slides/pages:** These will typically contain goal-oriented language, targets, key results, and strategic aims in both text and visually organized lists or impact statements.

    You must prioritize actual content slides/pages over simple divider or table of contents pages.
    Return a JSON object with 'best_match_index' (integer, or -1) and 'justification' (brief, one-sentence).
    """

    user_parts = [
        {"type": "text", "text": f"Find the best slide/page for '{slide_type_prompt}' in the '{deck_name}' with the following pages/slides:"}
    ]
    for slide_info in slides_data:
        user_parts.append({"type": "text", "text": f"\n--- Page/Slide {slide_info['slide_index'] + 1} (Text): {slide_info['text']}"})
        # Include the actual image data from the conversion service
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
            model="gpt-4o",
            messages=messages,
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        best_index = result.get("best_match_index", -1)
        justification = result.get("justification", "No justification provided.")
        
        # When finding a slide, we now have real data in slides_data.
        # Ensure we return the full data for the selected slide.
        selected_slide_data = slides_data[best_index] if best_index != -1 and best_index < len(slides_data) else None
        return {"slide": selected_slide_data, 
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
    Uses a multimodal AI (gpt-4o) to analyze GTM content (text + actual visual), find the best
    template layout (text + actual visual), and process the GTM content by inserting regional placeholders.
    """
    if not api_key:
        return {"best_template_index": -1, "justification": "OpenAI API Key is missing.", "processed_content": gtm_slide_content_data}

    client = openai.OpenAI(api_key=api_key)

    system_prompt = f"""
    You are an expert presentation content mapper. Your primary task is to help a user
    integrate content from a Global (GTM) slide/page into the most appropriate regional template.

    Given the `gtm_slide_content` (with its text and image) and a list of `template_slides_data`
    (each with an index and text content, and image data), you must perform two critical tasks:

    1.  **Select the BEST Template:**
        * **Crucially, you must review *each and every* template slide/page text summary AND its associated visual content.**
        * Semantically and **visually** evaluate which template slide's structure and implied purpose would *best* accommodate the `gtm_slide_content`.
        * **Perform a comparative analysis:** Do not just pick the first decent match. Compare all options to find the single most suitable template based on a combined understanding of text and visuals. **Prioritize templates where the text *imlies* a strong visual match, rather than just explicitly stating a type.** For instance, a template with short, sequential bullet points and dates might be a better visual timeline fit than one that simply has "Timeline" in its title but dense paragraphs.
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
    """
    
    user_parts = [
        {"type": "text", "text": f"User's original keyword for this content: '{user_keyword}'"},
        {"type": "text", "text": "GTM Slide/Page Content to Process (Text):"},
        {"type": "text", "text": json.dumps(gtm_slide_content_data.get('text', {}), indent=2)},
        # Include GTM slide's actual image data from the conversion service
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{gtm_slide_content_data['image_data']}"}} 
    ]

    user_parts.append({"type": "text", "text": "\nAvailable Template Slides/Pages Summary and Visuals:"})
    for slide_info in template_slides_data:
        user_parts.append({"type": "text", "text": f"\n--- Template Slide/Page {slide_info['slide_index'] + 1} (Text): {slide_info['text']}"})
        # Include each template slide's actual image data from the conversion service
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
            model="gpt-4o",
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
    st.header("2. Input Documents from GitHub")
    st.info("Files will be pulled from the specified GitHub repository. Ensure the GITHUB_TOKEN environment variable is set for private repos.")
    github_repo_url = st.text_input("GitHub Repository URL (e.g., https://github.com/org/repo-name.git)", value="https://github.com/lukerupolo/presemulator.git")
    github_branch = st.text_input("GitHub Branch (optional, default: main)", value="main")
    
    st.subheader("Template Documents Paths")
    template_paths_raw = st.text_area(
        "Comma-separated paths/folders for Template Deck(s) within repo (e.g., templates/part1.pptx, templates/part2/, templates/cover.pdf)",
        value="templates/slide_template.pptx" # Example path, replace with your actual
    )
    
    st.subheader("GTM Global Document Path")
    gtm_paths_raw = st.text_area(
        "Comma-separated paths/folders for GTM Global Deck(s) within repo (e.g., gtm/global.pptx, gtm/report.pdf)",
        value="gtm_decks/example_gtm.pdf" # Example path, replace with your actual
    )

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
if github_repo_url and template_paths_raw and gtm_paths_raw and api_key and st.session_state.structure:
    template_paths = [p.strip() for p in template_paths_raw.split(',') if p.strip()]
    gtm_paths = [p.strip() for p in gtm_paths_raw.split(',') if p.strip()]

    if not template_paths:
        st.error("Please provide at least one path for Template Deck(s).")
        st.stop()
    if not gtm_paths:
        st.error("Please provide at least one path for GTM Global Deck(s).")
        st.stop()

    if st.button("🚀 Assemble Presentation", type="primary"):
        with st.spinner("Assembling your new presentation..."):
            try:
                st.write("Step 1/3: Pulling documents from GitHub...")
                
                all_template_downloaded_files = _download_files_from_github(github_repo_url, github_branch, template_paths)
                all_gtm_downloaded_files = _download_files_from_github(github_repo_url, github_branch, gtm_paths)

                if not all_template_downloaded_files:
                    st.error("No template files found or downloaded from GitHub. Please check paths and repository.")
                    st.stop()
                if not all_gtm_downloaded_files:
                    st.error("No GTM files found or downloaded from GitHub. Please check paths and repository.")
                    st.stop()
                
                # For GTM, we process only the first found file for now
                gtm_file_to_process_bytes, gtm_file_to_process_type, gtm_file_to_process_name = all_gtm_downloaded_files[0]
                if len(all_gtm_downloaded_files) > 1:
                    st.warning(f"Multiple GTM files found. Only '{gtm_file_to_process_name}' will be processed.")

                st.write("Step 2/3: Loading and processing documents...")
                
                base_pptx_template_found = False
                new_prs = None 
                all_template_slides_for_ai = [] 

                for file_bytes, file_type, file_name in all_template_downloaded_files:
                    if file_type == 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
                        if not base_pptx_template_found:
                            new_prs = Presentation(io.BytesIO(file_bytes))
                            st.info(f"Using '{file_name}' as the primary base PPTX template.")
                            base_pptx_template_found = True
                        else:
                            current_prs_to_merge = Presentation(io.BytesIO(file_bytes))
                            st.info(f"Merging slides from '{file_name}' into the base template.")
                            for slide_to_merge in current_prs_to_merge.slides:
                                # Add a new blank slide to new_prs (with a generic layout)
                                new_slide = new_prs.slides.add_slide(new_prs.slide_layouts[0]) 
                                deep_copy_slide_content(new_slide, slide_to_merge) 
                    
                    all_template_slides_for_ai.extend(get_all_slide_data(file_bytes, file_type))

                if new_prs is None:
                    st.error("Error: At least one PPTX file must be found in the 'Template Documents Paths' to serve as the base for the assembled presentation.")
                    st.stop() 

                process_log = []
                st.write("Step 3/3: Building new presentation based on your structure...")
                
                num_template_slides = len(new_prs.slides) 
                num_structure_steps = len(st.session_state.structure)

                if num_structure_steps < num_template_slides:
                    for i in range(num_template_slides - 1, num_structure_steps - 1, -1):
                        rId = new_prs.slides._sldIdLst[i].rId
                        new_prs.part.drop_rel(rId)
                        del new_prs.slides._sldIdLst[i]
                    st.info(f"Removed {num_template_slides - num_structure_steps} unused slides from the merged template.")
                elif num_structure_steps > num_template_slides:
                     st.warning(f"Warning: Your defined structure has more steps ({num_structure_steps}) than the merged template has slides ({num_template_slides}). Extra steps will be ignored.")

                for i, step in enumerate(st.session_state.structure):
                    if i >= len(new_prs.slides): 
                        break

                    current_dest_slide_index = i
                    dest_slide = new_prs.slides[current_dest_slide_index] 
                    
                    keyword = step["keyword"]
                    action = step["action"]
                    
                    log_entry = {"step": i + 1, "keyword": keyword, "action": action, "log": []}
                    
                    if action == "Copy from GTM (as is)":
                        if gtm_file_to_process_type == 'application/vnd.openxmlformats-officedocument.presentationml.presentation': 
                            gtm_prs = Presentation(io.BytesIO(gtm_file_to_process_bytes))
                            result = find_slide_by_ai(api_key, gtm_file_to_process_bytes, gtm_file_to_process_type, keyword, "GTM Deck")
                            log_entry["log"].append(f"**GTM Content Choice Justification (PPTX Copy):** {result['justification']}")
                            if result["slide"]:
                                src_slide_object = gtm_prs.slides[result["index"]] 
                                deep_copy_slide_content(dest_slide, src_slide_object)
                                log_entry["log"].append(f"**Action:** Replaced Template slide {current_dest_slide_index + 1} with content from GTM PPTX slide {result['index'] + 1}.")
                            else:
                                log_entry["log"].append("**Action:** No suitable slide found in GTM PPTX deck. Template slide was left as is.")
                        else: 
                            log_entry["log"].append(f"**Warning:** 'Copy from GTM (as is)' is selected but GTM deck is a PDF. This action cannot directly copy PPTX shapes from a PDF. Proceeding with 'Merge' logic for content extraction based on text and assumed visuals.")
                            
                            gtm_ai_selection_result = find_slide_by_ai(api_key, gtm_file_to_process_bytes, gtm_file_to_process_type, keyword, "GTM Deck (Content Source)")
                            log_entry["log"].append(f"**GTM Content Source Justification (PDF Fallback Merge):** {gtm_ai_selection_result['justification']}")
                            
                            raw_gtm_content = {"title": "", "body": ""}
                            if gtm_ai_selection_result["slide"]:
                                full_text = gtm_ai_selection_result["slide"].get("text", "")
                                lines = full_text.split('\n')
                                raw_gtm_content["title"] = lines[0] if lines else ""
                                raw_gtm_content["body"] = "\n".join(lines[1:]) if len(lines) > 1 else ""

                            ai_mapping_result = analyze_and_map_content(
                                api_key, 
                                raw_gtm_content,
                                all_template_slides_for_ai, 
                                keyword
                            )
                            log_entry["log"].append(f"**AI Template Mapping Justification (PDF Fallback Merge):** {ai_mapping_result['justification']}")

                            selected_template_index = ai_mapping_result["best_template_index"]
                            processed_content = ai_mapping_result["processed_content"]

                            if selected_template_index != -1 and selected_template_index < len(new_prs.slides):
                                populate_slide(dest_slide, processed_content)
                                log_entry["log"].append(f"**Action:** Merged processed content from GTM (PDF) page {gtm_ai_selection_result['index'] + 1} into Template slide {current_dest_slide_index + 1}, with regional placeholders. AI suggested template type from template index {selected_template_index + 1}.")
                            else:
                                log_entry["log"].append("**Action:** AI could not determine a suitable template layout or process content for PDF. Template slide was left as is.")

                    elif action == "Merge: Template Layout + GTM Content":
                        gtm_ai_selection_result = find_slide_by_ai(api_key, gtm_file_to_process_bytes, gtm_file_to_process_type, keyword, "GTM Deck (Content Source)")
                        log_entry["log"].append(f"**GTM Content Source Justification:** {gtm_ai_selection_result['justification']}")
                        
                        raw_gtm_content = {"title": "", "body": ""}
                        if gtm_ai_selection_result["slide"]:
                            full_text = gtm_ai_selection_result["slide"].get("text", "")
                            lines = full_text.split('\n')
                            raw_gtm_content["title"] = lines[0] if lines else ""
                            raw_gtm_content["body"] = "\n".join(lines[1:]) if len(lines) > 1 else ""

                        ai_mapping_result = analyze_and_map_content(
                            api_key, 
                            raw_gtm_content,
                            all_template_slides_for_ai, 
                            keyword
                        )
                        log_entry["log"].append(f"**AI Template Mapping Justification:** {ai_mapping_result['justification']}")

                        selected_template_index = ai_mapping_result["best_template_index"]
                        processed_content = ai_mapping_result["processed_content"]

                        if selected_template_index != -1 and selected_template_index < len(new_prs.slides):
                            populate_slide(dest_slide, processed_content)
                            log_entry["log"].append(f"**Action:** Merged processed content from GTM ({gtm_file_to_process_name}) page/slide {gtm_ai_selection_result['index'] + 1} into Template slide {current_dest_slide_index + 1}, with regional placeholders. AI suggested template type from template index {selected_template_index + 1}.")
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
    st.info("Please provide an API Key, GitHub Repository URL, paths to Template/GTM documents, and define the structure in the sidebar to begin.")

