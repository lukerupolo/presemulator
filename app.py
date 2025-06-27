import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.dml import MSO_THEME_COLOR
import openai
import io
import os
import json

# --- Core Functions ---

def find_title_and_body_shapes(slide):
    """Finds the title and body shapes on a slide using heuristics."""
    title_shape, body_shape = None, None
    for shape in slide.placeholders:
        if 'TITLE' in shape.placeholder_format.type.name:
            title_shape = shape
        elif 'BODY' in shape.placeholder_format.type.name or 'OBJECT' in shape.placeholder_format.type.name:
            body_shape = shape
    if not title_shape and not body_shape: # Fallback for non-standard slides
        text_boxes = sorted([s for s in slide.shapes if s.has_text_frame and s.text.strip()], key=lambda s: len(s.text), reverse=True)
        if len(text_boxes) >= 1: body_shape = text_boxes[0]
        if len(text_boxes) >= 2: title_shape = text_boxes[1]
    return title_shape, body_shape

def extract_text_from_slide(slide):
    """Extracts all text from a single slide for classification."""
    return " ".join(shape.text for shape in slide.shapes if shape.has_text_frame)

def classify_slides(api_key, slides_text):
    """Uses AI to classify each slide."""
    try:
        client = openai.OpenAI(api_key=api_key)
    except Exception as e:
        raise ValueError(f"Failed to initialize OpenAI client: {e}")

    system_prompt = """
    You are a presentation analyst. You will be given a JSON object with slide numbers and their text content.
    Your task is to classify each slide's primary purpose. The valid classifications are: "Objectives", "Timeline", or "Other".
    You MUST return a JSON object with a single key 'classifications', containing a list of strings (e.g., ["Objectives", "Timeline", "Other", ...]).
    This list must have the exact same number of elements as the input.
    """
    full_user_prompt = f"Classify the following slide contents:\n{json.dumps(slides_text, indent=2)}"
    
    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo",
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": full_user_prompt}],
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)['classifications']
    except Exception as e:
        st.error(f"AI classification failed: {e}")
        return None

def delete_slides(prs, indices_to_delete):
    """Deletes slides from a presentation object by their indices."""
    slides = prs.slides
    # Sort indices in descending order to avoid re-indexing issues during deletion
    for idx in sorted(indices_to_delete, reverse=True):
        rId = slides._sldIdLst[idx].rId
        slides.part.drop_rel(rId)
        del slides._sldIdLst[idx]

def get_ai_modified_timeline_content(api_key, slide_content):
    """Sends timeline content to the AI with a specific prompt."""
    try:
        client = openai.OpenAI(api_key=api_key)
    except Exception as e:
        raise ValueError(f"Failed to initialize OpenAI client: {e}")

    system_prompt = """
    You are a marketing manager working on a regional response deck. You will be given the content of a timeline slide.
    Your task is to rewrite the content, filling in the necessary parts of the timeline section that will be adjusted for relevant regional responses. 
    Insert a clear placeholder like '[INSERT REGIONAL DETAIL HERE]' where specific local information needs to be added.
    Return a JSON object with 'title' and 'body' keys containing the rewritten text.
    """
    full_user_prompt = f"Adapt the following timeline slide content:\n{json.dumps(slide_content, indent=2)}"

    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo",
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": full_user_prompt}],
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        st.error(f"AI content modification for timeline failed: {e}")
        return slide_content # Return original on failure

def preserve_and_set_text(text_frame, new_text):
    """Replaces text in a text_frame while preserving formatting."""
    if not text_frame.paragraphs:
        p = text_frame.add_paragraph()
    else:
        p = text_frame.paragraphs[0]
        
    font_name, font_size, font_bold, font_italic, font_color = None, Pt(18), None, None, None
    if p.runs:
        original_font = p.runs[0].font
        font_name, font_size, font_bold, font_italic, font_color = original_font.name, original_font.size, original_font.bold, original_font.italic, original_font.color
    
    for para in list(text_frame.paragraphs):
        p_element = para._p
        p_element.getparent().remove(p_element)

    p = text_frame.add_paragraph()
    run = p.add_run()
    run.text = new_text
    
    font = run.font
    if font_name: font.name = font_name
    if font_size: font.size = font_size
    if font_bold is not None: font.bold = font_bold
    if font_italic is not None: font.italic = font_italic
    if font_color:
        if font_color.type == MSO_THEME_COLOR:
            font.color.theme_color, font.color.brightness = font_color.theme_color, font_color.brightness
        elif hasattr(font_color, 'rgb'):
            font.color.rgb = font_color.rgb

# --- UI Functions ---
def display_summary(summary_data):
    st.subheader("📝 Summary of Modifications")
    for item in summary_data:
        st.markdown(f"---")
        st.markdown(f"### Slide {item['slide_number']} (Kept as: **{item['classification']}**)")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Before**")
            st.text_area("Title", value=item['original']['title'], height=70, disabled=True, key=f"orig_title_{item['slide_number']}")
            st.text_area("Body", value=item['original']['body'], height=200, disabled=True, key=f"orig_body_{item['slide_number']}")
        with col2:
            st.markdown("**After**")
            st.text_area("Title", value=item['modified']['title'], height=70, disabled=True, key=f"mod_title_{item['slide_number']}")
            st.text_area("Body", value=item['modified']['body'], height=200, disabled=True, key=f"mod_body_{item['slide_number']}")

# --- Streamlit App ---
st.set_page_config(page_title="AI Presentation Pruner & Editor", layout="wide")
st.title("🤖 AI Presentation Pruner & Editor")
st.write("This tool automatically identifies and keeps only 'Objectives' and 'Timeline' slides, then intelligently adapts the timeline content for regional responses.")

with st.sidebar:
    st.header("Controls")
    api_key = st.text_input("Enter your OpenAI API Key", type="password")
    st.markdown("---")
    uploaded_file = st.file_uploader("Upload a PowerPoint (.pptx)", type=["pptx"])

if uploaded_file is not None:
    if st.button("✨ Process Presentation", type="primary"):
        if not api_key:
            st.error("Please enter your OpenAI API key.")
        else:
            with st.spinner("Processing... This is a multi-step process and may take some time."):
                try:
                    file_content = uploaded_file.getvalue()
                    prs = Presentation(io.BytesIO(file_content))
                    
                    # 1. Classify all slides
                    st.write("Step 1/5: Analyzing and classifying all slides...")
                    slides_text_for_classification = [{"slide_number": i+1, "text": extract_text_from_slide(s)} for i, s in enumerate(prs.slides)]
                    classifications = classify_slides(api_key, slides_text_for_classification)
                    
                    if not classifications: raise ValueError("Could not classify slides.")

                    # 2. Identify slides to delete and prepare for summary
                    indices_to_delete = []
                    kept_slides_info = []
                    for i, classification in enumerate(classifications):
                        if classification not in ["Objectives", "Timeline"]:
                            indices_to_delete.append(i)
                        else:
                            title_shape, body_shape = find_title_and_body_shapes(prs.slides[i])
                            kept_slides_info.append({
                                "original_index": i,
                                "classification": classification,
                                "original_content": {"title": title_shape.text if title_shape else "", "body": body_shape.text if body_shape else ""}
                            })
                    
                    st.write(f"Step 2/5: Pruning presentation... Deleting {len(indices_to_delete)} slide(s).")
                    delete_slides(prs, indices_to_delete)

                    # 3. Process remaining slides
                    st.write("Step 3/5: Adapting content for kept slides...")
                    summary_data = []
                    final_slides = prs.slides
                    for i in range(len(final_slides)):
                        slide_info = kept_slides_info[i]
                        current_content = slide_info["original_content"]
                        modified_content = {}

                        if slide_info["classification"] == "Objectives":
                            modified_content = current_content # Perfect copy
                        
                        elif slide_info["classification"] == "Timeline":
                            modified_content = get_ai_modified_timeline_content(api_key, current_content)

                        # Update the presentation object
                        title_shape, body_shape = find_title_and_body_shapes(final_slides[i])
                        if title_shape: preserve_and_set_text(title_shape.text_frame, modified_content.get('title', ''))
                        if body_shape: preserve_and_set_text(body_shape.text_frame, modified_content.get('body', ''))

                        summary_data.append({
                            "slide_number": i + 1,
                            "classification": slide_info["classification"],
                            "original": current_content,
                            "modified": modified_content
                        })

                    st.write("Step 4/5: Preparing your download...")
                    output_buffer = io.BytesIO()
                    prs.save(output_buffer)
                    output_buffer.seek(0)
                    
                    st.success("🎉 Your presentation has been successfully pruned and modified!")
                    
                    st.write("Step 5/5: Generating summary...")
                    display_summary(summary_data)
                    
                    new_filename = f"{os.path.splitext(uploaded_file.name)[0]}_regional_deck.pptx"
                    st.download_button(
                        label="Download Modified Presentation",
                        data=output_buffer,
                        file_name=new_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
                except Exception as e:
                    st.error(f"A critical error occurred: {e}")
                    st.exception(e)
else:
    st.info("Upload a PowerPoint and provide an API key to begin.")
