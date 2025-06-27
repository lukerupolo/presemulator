import streamlit as st
from pptx import Presentation
from pptx.util import Pt  # <--- FIX: Added this import statement
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
            # Check for body or object placeholders which typically hold the main content
            if shape.placeholder_format.type in ('BODY', 'OBJECT'):
                 body_shape = shape
                 break
        
        # Fallback: find the shape with the most text if no standard body placeholder is found
        if not body_shape:
            max_text_len = -1 # Use -1 to ensure any shape with text is chosen
            candidate_shape = None
            for shape in slide.shapes:
                if shape.has_text_frame and len(shape.text) > max_text_len:
                    max_text_len = len(shape.text)
                    candidate_shape = shape
            body_shape = candidate_shape
        
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
    
