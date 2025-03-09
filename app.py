import os
import json
import requests
from io import BytesIO
import matplotlib.pyplot as plt
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import PyPDF2  # For PDF parsing
from bs4 import BeautifulSoup  # For HTML parsing

# --- Configuration ---
GEMINI_API_KEY = st.secrets["API-KEY"]
# Ensure these endpoints point to the correct Gemini API endpoints.
GEMINI_TEXT_ENDPOINT = st.secrets["EP"]
GEMINI_IMAGE_ENDPOINT = st.secrets["EP"]

# --- Gemini API Functions ---

def gemini_text_generate(prompt, max_tokens=150, temperature=0.6):
    headers = {
        "Authorization": f"Bearer {GEMINI_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature
    }
    try:
        response = requests.post(GEMINI_TEXT_ENDPOINT, headers=headers, json=payload)
        response.raise_for_status()
    except Exception as e:
        st.error("Request to Gemini API failed.")
        st.error(str(e))
        raise

    raw_text = response.text.strip()
    if not raw_text:
        st.error("Gemini API returned an empty response.")
        raise ValueError("Gemini API returned an empty response.")

    # Check the Content-Type header
    content_type = response.headers.get("Content-Type", "").lower()
    if "application/json" in content_type:
        try:
            data = response.json()
        except json.JSONDecodeError as e:
            st.error("Failed to parse JSON from Gemini API. Raw response:")
            st.text(raw_text)
            raise e
        generated_text = data.get("generated_text", "").strip()
        if not generated_text:
            st.error("Gemini API did not return any generated text.")
            raise ValueError("Gemini API did not return any generated text.")
        return generated_text
    elif "text/html" in content_type or raw_text.lower().startswith("<!doctype html"):
        # Parse the HTML and extract text
        soup = BeautifulSoup(raw_text, "html.parser")
        parsed_text = soup.get_text(separator="\n", strip=True)
        if not parsed_text:
            st.error("Parsed HTML is empty.")
            raise ValueError("Parsed HTML is empty.")
        return parsed_text
    else:
        # Fallback: return raw_text if we don't recognize the content type.
        return raw_text

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

    # For image responses, if HTML is returned, extract text to notify the error.
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
    research_text = gemini_text_generate(prompt, max_tokens=150, temperature=0.5)
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
    outline_text = gemini_text_generate(prompt, max_tokens=400)
    try:
        slides = json.loads(outline_text)
    except json.JSONDecodeError as e:
        st.error("Error parsing JSON from Gemini response:")
        st.text(outline_text)
        raise e
    return slides

# --- PowerPoint Generation ---

def create_ppt_from_outline(slides, filename="generated_presentation.pptx"):
    prs = Presentation()
    for slide in slides:
        title_text = slide.get("title", "Untitled Slide")
        content_text = slide.get("content", "")
        image_prompt = slide.get("image_prompt")
        chart_info = slide.get("chart")

        slide_layout = prs.slide_layouts[5]
        ppt_slide = prs.slides.add_slide(slide_layout)
        ppt_slide.background.fill.solid()
        ppt_slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        # Title textbox
        title_box = ppt_slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        title_tf = title_box.text_frame
        title_tf.text = title_text
        for paragraph in title_tf.paragraphs:
            for run in paragraph.runs:
                run.font.name = "Lexend"
                run.font.bold = True
                run.font.size = Pt(44)
                run.font.color.rgb = RGBColor(75, 0, 130)

        # Content textbox
        content_box = ppt_slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(1.5))
        content_tf = content_box.text_frame
        content_tf.text = content_text
        for paragraph in content_tf.paragraphs:
            for run in paragraph.runs:
                run.font.name = "Lexend"
                run.font.size = Pt(24)
                run.font.color.rgb = RGBColor(255, 255, 255)

        # Research textbox
        research_summary = generate_deep_research_content(title_text, content_text)
        research_box = ppt_slide.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(9), Inches(1))
        research_tf = research_box.text_frame
        research_tf.text = research_summary
        for paragraph in research_tf.paragraphs:
            for run in paragraph.runs:
                run.font.name = "Lexend"
                run.font.size = Pt(18)
                run.font.color.rgb = RGBColor(211, 211, 211)

        # Optional image
        if image_prompt:
            st.info(f"Generating image for slide: {title_text}")
            image_data = gemini_image_generate(image_prompt)
            image_stream = BytesIO(image_data)
            ppt_slide.shapes.add_picture(image_stream, Inches(6), Inches(3.5), width=Inches(3))

        # Optional chart
        if chart_info:
            st.info(f"Generating chart for slide: {title_text}")
            chart_stream = generate_chart(chart_info)
            ppt_slide.shapes.add_picture(chart_stream, Inches(0.5), Inches(3.5), width=Inches(4))

    prs.save(filename)
    return filename

# --- Streamlit App ---

def main():
    st.title("AI-Driven Presentation Generator")
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
        if st.button("Generate Presentation"):
            with st.spinner("Generating slide outline..."):
                try:
                    slides_outline = generate_slide_outline(analysis_text)
                except Exception as e:
                    st.error(f"Failed to generate slide outline: {e}")
                    return
            st.success("Slide outline generated successfully!")
            st.json(slides_outline)

            with st.spinner("Creating PowerPoint presentation..."):
                try:
                    ppt_filename = create_ppt_from_outline(slides_outline)
                except Exception as e:
                    st.error(f"Failed to create presentation: {e}")
                    return
            st.success("Presentation created successfully!")
            with open(ppt_filename, "rb") as f:
                st.download_button(
                    label="Download Presentation",
                    data=f,
                    file_name=ppt_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

if __name__ == "__main__":
    main()
