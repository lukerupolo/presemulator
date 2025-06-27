import streamlit as st
from pptx import Presentation
import openai
import io
import os
import json

# --- Core Functions ---

def extract_text_from_pptx(prs):
    """Extracts text from each slide, returning a list of strings."""
    slide_texts = []
    for slide in prs.slides:
        text_on_slide = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text_on_slide.append(run.text)
        slide_texts.append(" ".join(text_on_slide))
    return slide_texts

def get_ai_modified_content(api_key, original_texts, user_prompt):
    """
    Sends the extracted text and a user prompt to OpenAI for modification.

    Args:
        api_key (str): The user's OpenAI API key.
        original_texts (list): A list of strings, where each string is the text of a slide.
        user_prompt (str): The user's instruction for the AI.

    Returns:
        list: A list of modified slide texts from the AI.
    """
    try:
        client = openai.OpenAI(api_key=api_key)
    except Exception as e:
        raise ValueError(f"Failed to initialize OpenAI client. Check your API key. Error: {e}")

    # The system prompt guides the AI to return data in a specific JSON format.
    system_prompt = """
    You are an expert presentation editor. You will be given the text content of a series of slides as a JSON array of strings.
    You will also be given an instruction. Your task is to rewrite the text for each slide according to the instruction.
    You MUST return a JSON object with a single key 'modified_slides'. This key should contain an array of strings.
    This array must have the exact same number of elements as the input array.
    Each string in the output array should be the complete, rewritten text for the corresponding slide.
    Do not add any extra commentary. Only return the JSON object.
    """

    # We combine the user's prompt with the extracted text for the AI.
    full_user_prompt = f"""
    Instruction: "{user_prompt}"

    Original slide texts (JSON array):
    {json.dumps(original_texts, indent=2)}
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4-turbo",  # Using a powerful model for better results
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_user_prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.7,
        )
        response_content = response.choices[0].message.content
        modified_data = json.loads(response_content)
        
        if 'modified_slides' not in modified_data or not isinstance(modified_data['modified_slides'], list):
            raise ValueError("AI response did not contain 'modified_slides' list.")
            
        return modified_data['modified_slides']

    except Exception as e:
        st.error(f"An error occurred while communicating with OpenAI: {e}")
        st.error(f"OpenAI Response (if any): {response_content if 'response_content' in locals() else 'N/A'}")
        return None


def update_presentation_with_new_text(prs, new_texts):
    """
    Replaces the text in the presentation with the modified text from the AI.
    It targets the main content placeholder on each slide.
    """
    if len(prs.slides) != len(new_texts):
        st.warning("Warning: The number of modified slides from the AI does not match the original presentation.")
        return prs

    for i, slide in enumerate(prs.slides):
        new_text = new_texts[i]
        
        # Heuristic: Find the largest text placeholder (body) to replace content
        body_shape = None
        for shape in slide.placeholders:
            if shape.placeholder_format.type in ('BODY', 'OBJECT'):
                 body_shape = shape
                 break
        
        # Fallback: find the shape with the most text if no body placeholder
        if not body_shape:
            max_text_len = 0
            for shape in slide.shapes:
                if shape.has_text_frame and len(shape.text) > max_text_len:
                    max_text_len = len(shape.text)
                    body_shape = shape
        
        if body_shape:
            text_frame = body_shape.text_frame
            text_frame.clear()  # Clear existing content
            p = text_frame.paragraphs[0]
            p.text = new_text
            p.font.size = Pt(18) # You can set a default font size
        else:
            # If still no suitable shape, you might add a new text box
            # For now, we'll just log a warning for that slide
            st.warning(f"Could not find a suitable text box on slide {i+1} to replace content.")

    return prs

# --- Streamlit UI ---

st.set_page_config(page_title="AI PowerPoint Editor", layout="wide")
st.title("🤖 AI-Powered PowerPoint Content Editor")
st.write("This application uses AI to analyze and rewrite the content of your PowerPoint presentation based on your instructions.")

# --- Sidebar for Inputs ---
with st.sidebar:
    st.header("Controls")
    api_key = st.text_input("Enter your OpenAI API Key", type="password")
    
    st.markdown("---")
    
    uploaded_file = st.file_uploader("1. Upload a PowerPoint (.pptx)", type=["pptx"])
    
    st.markdown("---")

    user_prompt = st.text_area(
        "2. Enter your editing instruction",
        height=150,
        placeholder="e.g., 'Rewrite the objectives on these slides to reflect an Australian market perspective.' or 'Summarize the key points on each slide into three bullet points.'"
    )

# --- Main App Logic ---
if uploaded_file is not None:
    original_filename = uploaded_file.name
    
    if st.button("✨ Process Presentation with AI", type="primary"):
        if not api_key:
            st.error("Please enter your OpenAI API key in the sidebar.")
        elif not user_prompt:
            st.error("Please enter an instruction for the AI in the sidebar.")
        else:
            with st.spinner("Processing your presentation... This may take a moment."):
                try:
                    # 1. Load presentation and extract text
                    st.write("Step 1/4: Reading your presentation...")
                    file_content = uploaded_file.getvalue()
                    prs = Presentation(io.BytesIO(file_content))
                    original_texts = extract_text_from_pptx(prs)

                    # 2. Send to AI for modification
                    st.write("Step 2/4: Asking the AI to rewrite the content...")
                    modified_texts = get_ai_modified_content(api_key, original_texts, user_prompt)

                    if modified_texts:
                        # 3. Update the presentation with new text
                        st.write("Step 3/4: Updating the slides with the new content...")
                        # We need to reload the presentation object to have a fresh one to edit
                        prs_to_edit = Presentation(io.BytesIO(file_content))
                        updated_prs = update_presentation_with_new_text(prs_to_edit, modified_texts)

                        # 4. Save and provide download link
                        st.write("Step 4/4: Preparing your download...")
                        output_buffer = io.BytesIO()
                        updated_prs.save(output_buffer)
                        output_buffer.seek(0)
                        
                        base, ext = os.path.splitext(original_filename)
                        new_filename = f"{base}_ai_modified.pptx"

                        st.success("🎉 Your presentation has been successfully modified!")
                        
                        st.download_button(
                            label="Download Modified PowerPoint",
                            data=output_buffer,
                            file_name=new_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
                except Exception as e:
                    st.error(f"A critical error occurred: {e}")
else:
    st.info("Upload a PowerPoint file and provide your API key and instructions in the sidebar to begin.")

