# new_conversion_service.py
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
import os
import tempfile
import base64
import io
from pptx import Presentation # For text extraction
import win32com.client # For controlling PowerPoint
import pythoncom # For COM initialization

app = FastAPI()

# --- New Helper Function using PowerPoint Automation ---

def _convert_pptx_to_images_and_text_windows(pptx_bytes: bytes) -> list[dict]:
    """
    Converts a PPTX file to images and extracts text by automating the
    PowerPoint application on Windows.
    Requires PowerPoint to be installed.
    """
    results = []
    
    # Use a temporary directory to save the file and images
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_pptx_path = os.path.join(temp_dir, "input.pptx")
        
        # Write the uploaded bytes to a temporary PPTX file
        with open(temp_pptx_path, "wb") as f:
            f.write(pptx_bytes)

        powerpoint = None
        presentation = None
        try:
            # Initialize the COM library for this thread
            pythoncom.CoInitialize()
            
            # Start the PowerPoint application
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1 # Make it visible for debugging, can be set to 0
            
            # Open the presentation
            presentation = powerpoint.Presentations.Open(temp_pptx_path, WithWindow=False)

            # Use python-pptx to get the text content in parallel
            prs_for_text = Presentation(io.BytesIO(pptx_bytes))

            # Iterate through each slide
            for i, slide in enumerate(prs_for_text.slides):
                # 1. Export the slide as a PNG image using PowerPoint
                image_path = os.path.join(temp_dir, f"slide_{i+1}.png")
                presentation.Slides[i].Export(image_path, "PNG")

                # 2. Read the exported image bytes and encode to Base64
                with open(image_path, "rb") as img_file:
                    image_data = base64.b64encode(img_file.read()).decode('utf-8')

                # 3. Extract text using python-pptx
                slide_text_content = []
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        slide_text_content.append(shape.text)
                text = " ".join(slide_text_content)[:2000]

                results.append({
                    "slide_index": i,
                    "text": text,
                    "image_data": image_data
                })

        except Exception as e:
            # If anything goes wrong, raise an error
            raise HTTPException(status_code=500, detail=f"PowerPoint automation failed: {e}. Ensure PowerPoint is installed and not blocked by security settings.")
        finally:
            # 4. VERY IMPORTANT: Clean up the COM objects
            if presentation:
                presentation.Close()
            if powerpoint:
                powerpoint.Quit()
            # Uninitialize the COM library
            pythoncom.CoUninitialize()

    return results


@app.post("/convert_document")
async def convert_document_endpoint(file: UploadFile = File(...)):
    """
    Endpoint to convert an uploaded PPTX document.
    NOTE: PDF conversion is removed as this strategy is Windows-only for PPTX.
    """
    file_bytes = await file.read()
    file_type = file.content_type

    if file_type == 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
        # Use the new Windows-specific conversion function
        slides_data = _convert_pptx_to_images_and_text_windows(file_bytes)
    else:
        raise HTTPException(status_code=400, detail=f"Unsupported file type: {file_type}. This version only supports PPTX files.")
    
    return JSONResponse(content={"slides": slides_data})
