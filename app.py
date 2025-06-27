import streamlit as st
from pptx import Presentation
import io
import os

def replace_text_in_pptx(file_content, find_text, replace_text):
    """
    Finds and replaces text in a PowerPoint presentation.

    Args:
        file_content (bytes): The content of the .pptx file.
        find_text (str): The text to search for (case-insensitive).
        replace_text (str): The text to replace it with.

    Returns:
        io.BytesIO: A bytes buffer containing the modified presentation.
    """
    # Load the presentation from the in-memory file content
    prs = Presentation(io.BytesIO(file_content))

    # Iterate through all slides in the presentation
    for slide in prs.slides:
        # Iterate through all shapes in each slide
        for shape in slide.shapes:
            # Check if the shape has a text frame
            if not shape.has_text_frame:
                continue
            
            # Iterate through all paragraphs in the text frame
            for paragraph in shape.text_frame.paragraphs:
                # Iterate through all runs in each paragraph
                # A run is a contiguous stretch of text with the same formatting
                for run in paragraph.runs:
                    # Perform a case-insensitive replacement
                    if find_text.lower() in run.text.lower():
                        # Simple text replacement
                        # Note: This replaces the entire run's text.
                        # For more complex scenarios where a run might contain the word multiple times
                        # or mixed with other words, a more sophisticated replacement logic
                        # would be needed to preserve formatting.
                        run.text = run.text.lower().replace(find_text.lower(), replace_text)


    # Save the modified presentation to an in-memory bytes buffer
    output_buffer = io.BytesIO()
    prs.save(output_buffer)
    # Rewind the buffer to the beginning
    output_buffer.seek(0)
    return output_buffer


# --- Streamlit UI ---
st.set_page_config(page_title="PowerPoint Text Replacer", layout="centered")
st.title("PowerPoint Text Replacer")
st.write("This app finds every instance of the word 'FRANCHISE' in your PowerPoint file and replaces it with 'BUNNIES'.")

# File uploader for .pptx files
uploaded_file = st.file_uploader("Upload a PowerPoint (.pptx) file", type=["pptx"])

if uploaded_file is not None:
    # Get the file content as bytes
    file_content = uploaded_file.getvalue()
    
    # Get the original filename
    original_filename = uploaded_file.name
    
    # Create a new filename for the output
    base, ext = os.path.splitext(original_filename)
    new_filename = f"{base}_modified_with_bunnies.pptx"

    st.success(f"File '{original_filename}' uploaded successfully!")

    # Add a button to trigger the replacement process
    if st.button("Replace 'FRANCHISE' with 'BUNNIES'"):
        with st.spinner("Finding and replacing text... Please wait."):
            try:
                # Call the function to perform the text replacement
                modified_pptx_buffer = replace_text_in_pptx(file_content, "FRANCHISE", "BUNNIES")

                st.success("Text replacement complete!")

                # Provide a download button for the modified file
                st.download_button(
                    label="Download Modified PowerPoint",
                    data=modified_pptx_buffer,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            except Exception as e:
                st.error(f"An error occurred during processing: {e}")

else:
    st.info("Please upload a .pptx file to get started.")
