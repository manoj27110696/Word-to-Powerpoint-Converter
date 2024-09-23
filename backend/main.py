from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from docx import Document
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
import os

app = FastAPI()

# Allow CORS for all origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["POST"],  # Restrict to the methods you need
    allow_headers=["*"],  # You can also restrict headers if needed
)

@app.post("/convert")
async def convert(docx_file: UploadFile = File(...)):
    output_pptx = 'output.pptx'

    doc = Document(docx_file.file)
    prs = Presentation()

    # Split the text into chunks of 4 lines, keeping headers
    lines = []
    chunks = []
    current_chunk = []
    current_title = ""

    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Heading'):
            if current_chunk:
                chunks.append((current_title, current_chunk))
                current_chunk = []
            current_title = paragraph.text
        else:
            if paragraph.text.strip() != "":
                current_chunk.append(paragraph.text)
                if len(current_chunk) == 4:
                    chunks.append((current_title, current_chunk))
                    current_chunk = []

    if current_chunk:
        chunks.append((current_title, current_chunk))

    # Add each chunk as a new slide
    for title, chunk in chunks:
        add_slide_with_text(prs, title, chunk)

    # Save the updated presentation
    prs.save(output_pptx)

    return FileResponse(output_pptx, filename=output_pptx, media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation')

def set_white_text_formatting(text_frame, font_size):
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(255, 255, 255)  # White text
            run.font.size = Pt(font_size)  # Set font size
        paragraph.alignment = PP_ALIGN.CENTER  # Center the text

def remove_bullets(text_frame):
    for paragraph in text_frame.paragraphs:
        paragraph.level = 0  # Ensure no bullet levels are set
        pPr = paragraph._element.get_or_add_pPr()  # Get or add paragraph properties
        if pPr is not None:
            buNone = pPr.find('.//a:buNone', namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
            if buNone is None:
                buNone = parse_xml('<a:buNone %s/>' % nsdecls('a'))
                pPr.append(buNone)

def set_slide_background(slide, color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_slide_with_text(prs, title_text, lines):
    slide_layout = prs.slide_layouts[1]  # Title and Content layout
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]

    title.text = title_text
    content.text = "\n".join(lines)

    # Set font sizes
    for paragraph in title.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(60)  # Set header font size to 60

    # Remove bullets and center text
    remove_bullets(content.text_frame)
    set_white_text_formatting(content.text_frame, 54)  # Set text font size to 54

    # Set the background color to dark
    set_slide_background(slide, RGBColor(0, 0, 0))  # Dark background