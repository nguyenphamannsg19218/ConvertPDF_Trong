import pdfplumber
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import re
import os
import latex2mathml.converter
from lxml import etree
from PIL import Image
import pytesseract
import io
import tempfile

def extract_latex_formulas(text):
    """
    Extract LaTeX formulas from text using regex patterns.
    
    Args:
        text (str): Input text containing LaTeX formulas.
    
    Returns:
        list: List of tuples (is_formula, content) where is_formula indicates if content is LaTeX.
    """
    inline_pattern = r'\$(.*?)\$'
    display_pattern = r'\$\$(.*?)\$\$|\[(.*?)\]'
    
    segments = []
    last_pos = 0
    
    for match in re.finditer(f'({inline_pattern})|({display_pattern})', text, re.DOTALL):
        start, end = match.span()
        if start > last_pos:
            segments.append((False, text[last_pos:start]))
        
        formula = match.group(1) or match.group(3) or match.group(4)
        if formula:
            formula = formula.strip('$[]')
            segments.append((True, formula))
        
        last_pos = end
    
    if last_pos < len(text):
        segments.append((False, text[last_pos:]))
    
    return segments

def latex_to_mathml(latex):
    """
    Convert LaTeX formula to MathML.
    
    Args:
        latex (str): LaTeX formula string.
    
    Returns:
        str: MathML string or None if conversion fails.
    """
    try:
        mathml = latex2mathml.converter.convert(latex)
        return mathml
    except Exception as e:
        print(f"Error converting LaTeX to MathML: {str(e)}")
        return None

def add_mathml_to_doc(doc, mathml):
    """
    Add MathML content to a Word document.
    
    Args:
        doc: python-docx Document object.
        mathml (str): MathML string to add.
    """
    try:
        mathml_tree = etree.fromstring(mathml)
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        omath = etree.Element('{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath')
        omath.append(mathml_tree)
        run._element.append(omath)
    except Exception as e:
        print(f"Error adding MathML to document: {str(e)}")

def extract_latex_from_image(image):
    """
    Extract LaTeX code from an image using Tesseract OCR.
    
    Args:
        image: PIL Image object.
    
    Returns:
        str: Extracted LaTeX code or None if extraction fails.
    """
    try:
        # Preprocess image for better OCR results
        image = image.convert('L')  # Convert to grayscale
        image = image.point(lambda x: 0 if x < 128 else 255, '1')  # Binarize
        
        # Perform OCR
        latex = pytesseract.image_to_string(image, lang='eng', config='--psm 6')
        
        # Clean up extracted text to identify LaTeX
        latex = latex.strip()
        if latex:
            return latex
        return None
    except Exception as e:
        print(f"Error extracting LaTeX from image: {str(e)}")
        return None

def convert_pdf_to_word(pdf_path, word_path):
    """
    Convert a PDF file to a Word document, preserving LaTeX formulas (text and image-based).
    
    Args:
        pdf_path (str): Path to the input PDF file.
        word_path (str): Path to save the output Word document.
    """
    try:
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found at: {pdf_path}")

        doc = Document()
        
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Extract text
                text = page.extract_text()
                if text:
                    segments = extract_latex_formulas(text)
                    for is_formula, content in segments:
                        if is_formula:
                            mathml = latex_to_mathml(content)
                            if mathml:
                                add_mathml_to_doc(doc, mathml)
                        else:
                            paragraph = doc.add_paragraph()
                            run = paragraph.add_run(content)
                            run.font.name = 'Times New Roman'
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                            run.font.size = Pt(12)
                
                # Extract images and process for potential LaTeX formulas
                if hasattr(page, 'images') and page.images:
                    for img in page.images:
                        # Extract image from PDF
                        x0, y0, x1, y1 = img['x0'], img['top'], img['x1'], img['bottom']
                        img_crop = page.crop((x0, y0, x1, y1)).to_image(resolution=300)
                        img_pil = img_crop.original
                        
                        # Try to extract LaTeX from image
                        latex = extract_latex_from_image(img_pil)
                        if latex:
                            mathml = latex_to_mathml(latex)
                            if mathml:
                                add_mathml_to_doc(doc, mathml)
        
        # Save the Word document
        doc.save(word_path)
        print(f"Conversion successful! Word document saved at: {word_path}")
        
    except Exception as e:
        print(f"Error during conversion: {str(e)}")

def main():
    pdf_path = input("Enter the path to the PDF file: ")
    word_path = input("Enter the path to save the Word file (e.g., output.docx): ")
    
    output_dir = os.path.dirname(word_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    convert_pdf_to_word(pdf_path, word_path)

if __name__ == "__main__":
    main()
