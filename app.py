# streamlit_app.py
import io
import os
import time
import tempfile
import streamlit as st
from pdf2image import convert_from_bytes
from pdf2docx import Converter
from docx import Document
from docx.shared import Inches
from PIL import Image
import pytesseract

st.set_page_config(page_title="PDF → Word (Visual/Text/Hybrid)", layout="centered")

def compress_image(image, quality=60):
    img_io = io.BytesIO()
    image.save(img_io, format="JPEG", optimize=True, quality=quality)
    img_io.seek(0)
    return Image.open(img_io)

def pdf_to_word_visual(pdf_bytes, dpi=150, quality=60, max_pages=None):
    images = convert_from_bytes(pdf_bytes, dpi=dpi)
    if max_pages:
        images = images[:max_pages]

    doc = Document()
    for i, img in enumerate(images):
        img = compress_image(img, quality)
        img_byte = io.BytesIO()
        img.save(img_byte, format="JPEG")
        img_byte.seek(0)
        if i > 0:
            doc.add_page_break()
        doc.add_picture(img_byte, width=Inches(6.8))
    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream

def pdf_to_word_text(pdf_bytes, max_pages=None):
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, "input.pdf")
        out_path = os.path.join(tmpdir, "output.docx")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)
        cv = Converter(pdf_path)
        if max_pages:
            cv.convert(out_path, start=0, end=max_pages-1)
        else:
            cv.convert(out_path)
        cv.close()
        with open(out_path, "rb") as f:
            return io.BytesIO(f.read())

def pdf_to_word_hybrid(pdf_bytes, dpi=150, quality=60, lang="eng", max_pages=None):
    images = convert_from_bytes(pdf_bytes, dpi=dpi)
    if max_pages:
        images = images[:max_pages]

    doc = Document()
    for i, img in enumerate(images):
        img = compress_image(img, quality)
        img_byte = io.BytesIO()
        img.save(img_byte, format="JPEG")
        img_byte.seek(0)

        if i > 0:
            doc.add_page_break()
        doc.add_picture(img_byte, width=Inches(6.8))

        # OCR text từ ảnh
        text = pytesseract.image_to_string(img, lang=lang).strip()
        if text:
            doc.add_paragraph("\n" + text)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream

st.title("PDF → Word Converter")
st.write("Chọn chế độ: hình ảnh (Visual), trích text (Text), hoặc kết hợp OCR (Hybrid).")

with st.sidebar:
    mode = st.selectbox("Chế độ", ["visual", "text", "hybrid"])
    dpi = st.number_input("DPI (ảnh)", min_value=72, max_value=300, value=150, step=10)
    quality = st.slider("Chất lượng JPEG (%)", min_value=30, max_value=95, value=60, step=5)
    max_pages = st.number_input("Giới hạn số trang (0 = tất cả)", min_value=0, value=0, step=1)
    lang = st.text_input("Ngôn ngữ OCR (ví dụ: eng, vie, eng+vie)", value="eng")

file = st.file_uploader("Tải lên file PDF", type=["pdf"])

if file:
    pdf_bytes = file.read()
    if st.button("Convert"):
        start = time.time()
        try:
            mp = int(max_pages) if max_pages else 0
            mp = mp if mp > 0 else None

            if mode == "visual":
                out_stream = pdf_to_word_visual(pdf_bytes, dpi=dpi, quality=quality, max_pages=mp)
                out_name = file.name.replace(".pdf", "_visual.docx")
            elif mode == "text":
                out_stream = pdf_to_word_text(pdf_bytes, max_pages=mp)
                out_name = file.name.replace(".pdf", "_text.docx")
            else:  # hybrid
                out_stream = pdf_to_word_hybrid(pdf_bytes, dpi=dpi, quality=quality, lang=lang, max_pages=mp)
                out_name = file.name.replace(".pdf", "_hybrid.docx")

            st.success(f"Xong trong {time.time()-start:.1f}s")
            st.download_button("Tải file DOCX", data=out_stream, file_name=out_name,
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Lỗi: {e}")
