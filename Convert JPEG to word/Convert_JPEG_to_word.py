import pytesseract
from PIL import Image
from docx import Document
import cv2

# Specify the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def jpeg_to_word():
    # Ask for the JPEG file path
    jpeg_path = input("Enter the path to the JPEG file: ")
    # Ask for the Word document path
    word_path = input("Enter the path to save the Word document: ")
    # Ensure the path ends with .docx
    if not word_path.endswith('.docx'):
        word_path += '.docx'

    # Open the image file
    img = cv2.imread(jpeg_path)

    # Convert the image to gray scale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Apply threshold to get image with only black and white
    _, img_bin = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    img_bin = 255 - img_bin

    # Use Tesseract to do OCR on the image
    text = pytesseract.image_to_string(img_bin)

    # Create a new Word document
    doc = Document()
    # Add the text to the Word document
    doc.add_paragraph(text)
    # Save the Word document
    doc.save(word_path)

# Example usage:
jpeg_to_word()
