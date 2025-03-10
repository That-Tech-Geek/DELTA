import os
import re
import json
import uuid
import requests
from io import BytesIO
import matplotlib.pyplot as plt
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import PyPDF2  # For PDF parsing
from bs4 import BeautifulSoup  # For HTML parsing

# Additional imports for Google Slides API
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- Configuration ---
# Cohere configuration for text generation
COHERE_API_KEY = st.secrets["COHERE_API_KEY"]
COHERE_TEXT_ENDPOINT = st.secrets["COHERE_TEXT_EP"]  # e.g. "https://api.cohere.ai/generate"

# Gemini configuration for image generation
GEMINI_API_KEY = st.secrets["API-KEY"]
GEMINI_IMAGE_ENDPOINT = st.secrets["EP"]  # This endpoint should return image data

# Google service account credentials are stored under st.secrets["google_service_account"]

# --- Helper: Robust JSON Extraction ---
def extract_json(text):
    """
    Attempt to extract a valid JSON substring from a text response.
    Tries both object ({...}) and array ([...]) patterns.
    """
    obj_match = re.search(r'({.*})', text, re.DOTALL)
    if obj_match:
        try:
            return json.loads(obj_match.group(1))
        except Exception:
            pass
    arr_match = re.search(r'(\[.*\])', text, re.DOTALL)
    if arr_match:
        try:
            return json.loads(arr_match.group(1))
        except Exception:
            pass
    raise ValueError("No valid JSON could be extracted.")

# --- Cohere Text Generation Function ---
def cohere_text_generate(prompt, max_tokens=150, temperature=0.6):
    headers = {
        "Authorization": f"Bearer {COHERE_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "command-xlarge-nightly",  # Adjust model as needed
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature
    }
    try:
        response = requests.post(COHERE_TEXT_ENDPOINT, headers=headers, json=payload)
        response.raise_for_status()
    except Exception as e:
        st.error("Request to Cohere API failed.")
        st.error(str(e))
        raise

    raw_text = response.text.strip()
    if not raw_text:
        st.error("API Problem: Cohere API returned an empty response.")
        raise ValueError("Cohere API returned an empty response.")

    content_type = response.headers.get("Content-Type", "").lower()
    if "application/json" in content_type:
        try:
            data = response.json()
        except json.JSONDecodeError as e:
            st.error("Failed to parse JSON from Cohere API. Raw response:")
            st.text(raw_text)
            raise e
        try:
            generated_text = data["generations"][0]["text"].strip()
        except (KeyError, IndexError) as e:
            st.error("Cohere API did not return any generated text.")
            raise ValueError("Cohere API did not return any generated text.") from e
        return generated_text
    elif "text/html" in content_type or raw_text.lower().startswith("<!doctype html"):
        soup = BeautifulSoup(raw_text, "html.parser")
        parsed_text = soup.get_text(separator="\n", strip=True)
        if not parsed_text:
            st.error("Parsed HTML is empty.")
            raise ValueError("Parsed HTML is empty.")
        return parsed_text
    else:
        return raw_text

# --- Gemini Image Generation Function ---
def gemini_image_generate(prompt, width=512, height=512):
    headers = {
        "Authorization": f"Bearer {GEMINI_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "prompt": prompt,
        "width": width,
        "height": height
    }
    try:
        response = requests.post(GEMINI_IMAGE_ENDPOINT, headers=headers, json=payload)
        response.raise_for_status()
    except Exception as e:
        st.error("Request to Gemini Image API failed.")
        st.error(str(e))
        raise

    raw_data = response.content
    if raw_data.strip().lower().startswith(b"<!doctype html>"):
        html_text = BeautifulSoup(raw_data, "html.parser").get_text(separator="\n", strip=True)
        st.error("Expected image data but received HTML. Parsed HTML content:")
        st.text(html_text)
        raise ValueError("Non-image response received from Gemini Image API.")
    return raw_data

# --- Chart Generation ---
def generate_chart(chart_info):
    """
    Generates a chart using matplotlib based on chart_info.
    Expected format:
      {
        "type": "bar",  # or "line"
        "title": "Chart Title",
        "labels": ["Label1", "Label2", ...],
        "values": [val1, val2, ...]
      }
    """
    chart_type = chart_info.get("type", "bar")
    title = chart_info.get("title", "")
    labels = chart_info.get("labels", [])
    values = chart_info.get("values", [])
    
    plt.style.use('dark_background')
    plt.figure(figsize=(4, 3))
    
    if chart_type == "bar":
        plt.bar(labels, values, color='#4B0082')
    elif chart_type == "line":
        plt.plot(labels, values, marker='o', linestyle='-', color='#4B0082')
    
    plt.title(title, color='white')
    plt.tight_layout()
    
    img_stream = BytesIO()
    plt.savefig(img_stream, format='PNG', facecolor='black')
    plt.close()
    img_stream.seek(0)
    return img_stream

# --- Deep Research Generation ---
def generate_deep_research_content(slide_title, slide_content):
    prompt = (
        "You are a consultant at a top consulting firm. Provide a deep research summary for a client presentation slide with the title "
        f"'{slide_title}' and content: '{slide_content}'. Include key insights, critical analysis, and relevant references as bullet points. "
        "Output only the bullet points."
    )
    research_text = cohere_text_generate(prompt, max_tokens=150, temperature=0.5)
    return research_text

# --- Outline Generation ---
def generate_slide_outline(analysis_text):
    prompt = (
        "You are a consultant at a top consulting firm. Based on the following analysis, design a complete slide deck outline "
        "with natural flow. For each slide, provide a 'title' and 'content'. "
        "If an image would enhance the slide, include an 'image_prompt' key with a brief description. "
        "If a chart is needed, include a 'chart' key with an object specifying 'type' (bar or line), 'title', 'labels', and 'values'. "
        "Output a valid JSON array with no extra commentary.\n\n"
        "Analysis:\n" + analysis_text + "\n\nOutput the JSON array only."
    )
    outline_text = cohere_text_generate(prompt, max_tokens=400)
    if not outline_text:
        st.error("API Problem: The API returned an empty output for the slide outline.")
        raise ValueError("Empty output from API.")
    try:
        slides = json.loads(outline_text)
    except json.JSONDecodeError as e:
        st.error("Parsing Problem: The API returned non-empty output that could not be parsed as JSON. Raw output:")
        st.text(outline_text)
        try:
            slides = extract_json(outline_text)
            st.warning("JSON was extracted using regex fallback.")
        except Exception as e2:
            st.error("Failed to extract valid JSON from the response.")
            slides = None  # signal that parsing failed
    return slides, outline_text

# --- Convert Outline to Markdown ---
def convert_outline_to_md(slides):
    md = ""
    for idx, slide in enumerate(slides, start=1):
        title = slide.get("title", "Untitled Slide")
        content = slide.get("content", "")
        md += f"# Slide {idx}: {title}\n\n"
        md += f"**Content:**\n\n{content}\n\n"
        if "image_prompt" in slide:
            image_prompt = slide.get("image_prompt")
            md += f"**Image Prompt:** {image_prompt}\n\n"
        if "chart" in slide:
            chart = slide.get("chart")
            md += f"**Chart Details:**\n"
            md += f"- Type: {chart.get('type', '')}\n"
            md += f"- Title: {chart.get('title', '')}\n"
            labels = chart.get("labels", [])
            values = chart.get("values", [])
            if labels and values:
                md += f"- Labels: {', '.join(labels)}\n"
                md += f"- Values: {', '.join(map(str, values))}\n"
        md += "\n---\n\n"
    return md

# --- Google Slides Creation ---
def create_google_slides(slides_outline, presentation_title):
    # Build credentials from service account info stored in st.secrets["google_service_account"]
    creds = service_account.Credentials.from_service_account_info(st.secrets["google_service_account"])
    slides_service = build("slides", "v1", credentials=creds)
    # Create a new presentation
    presentation = slides_service.presentations().create(body={"title": presentation_title}).execute()
    presentation_id = presentation.get("presentationId")
    # For each slide in the outline, add a new slide with title and content text boxes
    for slide in slides_outline:
        slide_id = "slide_" + uuid.uuid4().hex
        title_box_id = "title_" + uuid.uuid4().hex
        content_box_id = "content_" + uuid.uuid4().hex
        requests_list = [
            {
                "createSlide": {
                    "objectId": slide_id,
                    "insertionIndex": "1",
                    "slideLayoutReference": {"predefinedLayout": "BLANK"}
                }
            },
            {
                "createShape": {
                    "objectId": title_box_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "height": {"magnitude": 50, "unit": "PT"},
                            "width": {"magnitude": 400, "unit": "PT"}
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 50,
                            "translateY": 50,
                            "unit": "PT"
                        }
                    }
                }
            },
            {
                "insertText": {
                    "objectId": title_box_id,
                    "insertionIndex": 0,
                    "text": slide.get("title", "Untitled Slide")
                }
            },
            {
                "createShape": {
                    "objectId": content_box_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "height": {"magnitude": 200, "unit": "PT"},
                            "width": {"magnitude": 400, "unit": "PT"}
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 50,
                            "translateY": 150,
                            "unit": "PT"
                        }
                    }
                }
            },
            {
                "insertText": {
                    "objectId": content_box_id,
                    "insertionIndex": 0,
                    "text": slide.get("content", "")
                }
            }
        ]
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": requests_list}
        ).execute()
    return presentation_id

# --- Streamlit App ---
def main():
    st.title("AI-Driven Google Slides Generator")
    st.write("Paste your compiled analysis below (including research, data, and insights).")

    # Option to either upload a PDF or paste text
    uploaded_file = st.file_uploader("Upload your analysis document (PDF file)", type="pdf")
    if uploaded_file is not None:
        try:
            pdf_reader = PyPDF2.PdfReader(uploaded_file)
            analysis_text = ""
            for page in pdf_reader.pages:
                text = page.extract_text()
                if text:
                    analysis_text += text + "\n"
        except Exception as e:
            st.error(f"Error parsing PDF: {e}")
            return
    else:
        analysis_text = st.text_area("Or paste your analysis text here", height=300)

    if analysis_text:
        if st.button("Generate Slide Outline, Markdown, and Google Slides"):
            with st.spinner("Generating slide outline..."):
                try:
                    slides_outline, raw_outline = generate_slide_outline(analysis_text)
                except Exception as e:
                    st.error(f"Failed to generate slide outline: {e}")
                    return
            if slides_outline:
                st.success("Slide outline generated and parsed as JSON successfully!")
                md_output = convert_outline_to_md(slides_outline)
                st.markdown("### Slide Outline in Markdown")
                st.markdown(md_output)
                with st.spinner("Creating Google Slides presentation..."):
                    try:
                        presentation_title = "Generated Presentation"
                        presentation_id = create_google_slides(slides_outline, presentation_title)
                        presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
                    except Exception as e:
                        st.error(f"Failed to create Google Slides presentation: {e}")
                        return
                st.success("Google Slides presentation created successfully!")
                st.markdown(f"**Presentation URL:** [Click Here]({presentation_url})")
            else:
                st.warning("Using raw Cohere output as Markdown since JSON parsing failed:")
                st.markdown("### Raw Outline Markdown")
                st.markdown(raw_outline)

if __name__ == "__main__":
    main()
