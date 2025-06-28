# conversion_service.py
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
import subprocess
import io
import base64
import fitz # PyMuPDF
import os
import shutil
import tempfile
import json
import traceback

app = FastAPI()

# --- THE FIX: PROVIDE THE FULL PATH TO UNOCONV ---
# We specify the full path to the unoconv executable to avoid any PATH issues.
UNOCONV_PATH = r"C:\Users\lukin\AppData\Local\Programs\Python\Python313\Scripts\unoconv"

def _convert_pptx_to_images_and_text(pptx_bytes: bytes) -> list[dict]:
    """
    Converts a PPTX file to a list of Base64 encoded PNG images, one per slide,
    and extracts text from each slide. Requires LibreOffice and unoconv.
    """
    results = []
    
    with tempfile.TemporaryDirectory() as temp_dir:
        pptx_path = os.path.join(temp_dir, "input.pptx")
        output_dir = os.path.join(temp_dir, "output_images")
        os.makedirs(output_dir)

        with open(pptx_path, "wb") as f:
            f.write(pptx_bytes)

        try:
            # --- MODIFIED COMMAND: Use the full path from UNOCONV_PATH ---
            command_render = [
                UNOCONV_PATH, "-f", "png", 
                "--output", os.path.join(output_dir, "slide-.png"),
                pptx_path
            ]
            # Use shell=True on Windows if direct execution fails, as it helps resolve command paths.
            result = subprocess.run(command_render, capture_output=True, text=True, check=False, shell=True)
            if result.returncode != 0:
                raise subprocess.CalledProcessError(result.returncode, command_render, output=result.stdout, stderr=result.stderr)
            
            from pptx import Presentation
            prs = Presentation(io.BytesIO(pptx_bytes))

            for i, slide in enumerate(prs.slides):
                slide_text_content = []
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        slide_text_content.append(shape.text)
                text = " ".join(slide_text_content)[:2000]

                image_path = os.path.join(output_dir, f"slide-{i+1}.png") 
                if not os.path.exists(image_path):
                    image_path = os.path.join(output_dir, f"slide-{i+1:02d}.png")
                if not os.path.exists(image_path):
                    print(f"Warning: Image file not found for slide {i+1} at {image_path}")
                    image_data = ""
                else:
                    with open(image_path, "rb") as img_file:
                        image_data = base64.b64encode(img_file.read()).decode('utf-8')

                results.append({
                    "slide_index": i,
                    "text": text,
                    "image_data": image_data
                })
        except subprocess.CalledProcessError as e:
            detail_message = f"PPTX conversion failed. unoconv exited with code {e.returncode}.\n"
            if e.stdout:
                detail_message += f"stdout:\n{e.stdout}\n"
            if e.stderr:
                detail_message += f"stderr:\n{e.stderr}\n"
            if not e.stdout and not e.stderr:
                detail_message += "No output captured from unoconv (check if LibreOffice is fully configured for headless mode)."
            raise HTTPException(status_code=500, detail=detail_message)
        except FileNotFoundError:
            # This error is now much more specific if it happens
            raise HTTPException(status_code=500, detail=f"The command '{UNOCONV_PATH}' was not found. Please verify the path is correct.")
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"An unexpected error occurred during PPTX processing: {e}")
            
    return results

def _convert_pdf_to_images_and_text(pdf_bytes: bytes) -> list[dict]:
    results = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for i, page in enumerate(doc):
            pix = page.get_pixmap()
            image_bytes = pix.tobytes("png")
            image_data = base64.b64encode(image_bytes).decode('utf-8')
            text = page.get_text()[:2000]
            results.append({
                "slide_index": i,
                "text": text,
                "image_data": image_data
            })
        doc.close()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF conversion failed: {e}")
    return results

@app.post("/convert_document")
async def convert_document_endpoint(file: UploadFile = File(...)):
    try:
        file_bytes = await file.read()
        file_type = file.content_type

        if file_type == 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
            slides_data = _convert_pptx_to_images_and_text(file_bytes)
        elif file_type == 'application/pdf':
            slides_data = _convert_pdf_to_images_and_text(file_bytes)
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported file type: {file_type}. Only PPTX and PDF are supported.")
        
        return JSONResponse(content={"slides": slides_data})

    except Exception as e:
        print("--- AN UNHANDLED ERROR OCCURRED ---")
        traceback.print_exc()
        print("-----------------------------------")
        raise HTTPException(status_code=500, detail=f"A critical error occurred in the conversion service: {e}")
