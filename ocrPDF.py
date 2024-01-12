# Script to create searchable PDF from scan PDF or images using Azure Form Recognizer
# Required packages
# pip install azure-ai-formrecognizer pypdf2 reportlab pillow pdf2image
import argparse
import io
import math
import os
import shutil
import sys
import tempfile

from PIL import Image, ImageSequence
from PyPDF2 import PdfWriter, PdfReader
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from dotenv import load_dotenv
from pdf2image import convert_from_path
from reportlab.lib import pagesizes
from reportlab.pdfgen import canvas

load_dotenv()
endpoint = os.getenv("ENDPOINT")
key = os.getenv("API_KEY")
print(endpoint)


def dist(p1, p2):
    return math.sqrt((p1.x - p2.x) * (p1.x - p2.x) + (p1.y - p2.y) * (p1.y - p2.y))


def split_pdf_into_pages(pdf_path):
    # Check if the provided path is a valid file
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"No file found at {pdf_path}")

    # Create a temporary directory
    temp_dir = tempfile.TemporaryDirectory()
    page_paths = []

    # Read the PDF
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)

        # Iterate through each page and save it as a new PDF
        for page_num in range(len(reader.pages)):
            writer = PdfWriter()
            writer.add_page(reader.pages[page_num])

            # Define the path for the individual page
            # output_filename = os.path.join(temp_dir.name, f'page_{page_num + 1}.pdf')
            output_filename = os.path.join(os.getcwd() + f"page_{page_num + 1}.pdf")
            with open(output_filename, 'wb') as outputpdf_file:
                writer.write(outputpdf_file)

            page_paths.append(output_filename)

    # Return the list of paths
    return page_paths


def combine_pdfs(pdf_paths):
    # Create a PdfWriter object
    pdf_writer = PdfWriter()

    # Iterate over each PDF path and add them to the writer object
    for path in pdf_paths:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"No file found at {path}")

        with open(path, 'rb') as file:
            pdf_reader = PdfReader(file)
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)

    # Create a temporary file for the combined PDF
    temp_dir = tempfile.TemporaryDirectory()
    # combined_pdf_path = os.path.join(temp_dir.name, 'combined.pdf')
    combined_pdf_path = os.path.join(os.getcwd(), 'combined.pdf')

    # Write the combined PDF to the file
    with open(combined_pdf_path, 'wb') as outputStream:
        pdf_writer.write(outputStream)

    # Return the path of the combined PDF
    return combined_pdf_path


def ocr_page(pdf_page):
    if args.output:
        output_file = args.output
    else:
        output_file = pdf_page + ".ocr.pdf"
    # Loading input file
    print(f"Loading input file {pdf_page}")
    if pdf_page.lower().endswith('.pdf'):
        # read existing PDF as images
        image_pages = convert_from_path(pdf_page)
    elif pdf_page.lower().endswith(('.tif', '.tiff', '.jpg', '.jpeg', '.png', '.bmp')):
        # read input image (potential multi page Tiff)
        image_pages = ImageSequence.Iterator(Image.open(pdf_page))
    else:
        sys.exit(
            f"Error: Unsupported input file extension {pdf_page}. Supported extensions: PDF, TIF, TIFF, JPG, JPEG, "
            f"PNG, BMP.")

    # Running OCR using Azure Form Recognizer Read API
    print(f"Starting Azure Form Recognizer OCR process...")
    document_analysis_client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(key),
                                                      headers={"x-ms-useragent": "searchable-pdf-blog/1.0.0"})

    with open(pdf_page, "rb") as f:
        poller = document_analysis_client.begin_analyze_document("prebuilt-read", document=f)

    ocr_results = poller.result()
    print(f"Azure Form Recognizer finished OCR text for {len(ocr_results.pages)} pages.")

    # Generate OCR overlay layer
    print(f"Generating searchable PDF...")
    output = PdfWriter()
    default_font = "Times-Roman"
    for page_id, pdf_page in enumerate(ocr_results.pages):
        ocr_overlay = io.BytesIO()

        # Calculate overlay PDF page size
        if image_pages[page_id].height > image_pages[page_id].width:
            page_scale = float(image_pages[page_id].height) / pagesizes.letter[1]
        else:
            page_scale = float(image_pages[page_id].width) / pagesizes.letter[1]

        page_width = float(image_pages[page_id].width) / page_scale
        page_height = float(image_pages[page_id].height) / page_scale

        scale = (page_width / pdf_page.width + page_height / pdf_page.height) / 2.0
        pdf_canvas = canvas.Canvas(ocr_overlay, pagesize=(page_width, page_height))

        # Add image into PDF page
        pdf_canvas.drawInlineImage(image_pages[page_id], 0, 0, width=page_width, height=page_height,
                                   preserveAspectRatio=True)

        text = pdf_canvas.beginText()
        # Set text rendering mode to invisible
        text.setTextRenderMode(3)
        for word in pdf_page.words:
            # Calculate optimal font size
            desired_text_width = max(dist(word.polygon[0], word.polygon[1]),
                                     dist(word.polygon[3], word.polygon[2])) * scale
            desired_text_height = max(dist(word.polygon[1], word.polygon[2]),
                                      dist(word.polygon[0], word.polygon[3])) * scale
            font_size = desired_text_height
            actual_text_width = pdf_canvas.stringWidth(word.content, default_font, font_size)

            # Calculate text rotation angle
            text_angle = math.atan2(
                (word.polygon[1].y - word.polygon[0].y + word.polygon[2].y - word.polygon[3].y) / 2.0,
                (word.polygon[1].x - word.polygon[0].x + word.polygon[2].x - word.polygon[3].x) / 2.0)
            text.setFont(default_font, font_size)
            text.setTextTransform(math.cos(text_angle), -math.sin(text_angle), math.sin(text_angle),
                                  math.cos(text_angle), word.polygon[3].x * scale,
                                  page_height - word.polygon[3].y * scale)
            text.setHorizScale(desired_text_width / actual_text_width * 100)
            text.textOut(word.content + " ")

        pdf_canvas.drawText(text)
        pdf_canvas.save()

        # Move to the beginning of the buffer
        ocr_overlay.seek(0)

        # Create a new PDF page
        new_pdf_page = PdfReader(ocr_overlay)
        output.add_page(new_pdf_page.pages[0])

    # Save output searchable PDF file
    with open(output_file, "wb") as outputStream:
        output.write(outputStream)

    print(f"Searchable PDF is created: {output_file}")
    return output_file


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('input_file', type=str, help="Input PDF or image (jpg, jpeg, tif, tiff, bmp, png) file name")
    parser.add_argument('-o', '--output', type=str, required=False, default="",
                        help="Output PDF file name. Default: input_file + .ocr.pdf")
    args = parser.parse_args()

    input_file = args.input_file

    final_path = os.path.join(os.path.dirname(input_file), os.path.splitext(os.path.basename(input_file))[0] + ".ocr.pdf")

    pages = split_pdf_into_pages(input_file)
    output_pages = []

    for page in pages:
        output_pages.append(ocr_page(page))
        os.remove(page)

    final_file = combine_pdfs(output_pages)

    for page in output_pages:
        os.remove(page)

    shutil.move(final_file, final_path)

    print("OCR completed!")
