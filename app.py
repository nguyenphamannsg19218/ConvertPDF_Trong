from flask import Flask, render_template, request, send_file
from pdf2image import convert_from_bytes
from pdf2docx import Converter
from docx import Document
from docx.shared import Inches
from PIL import Image
import pytesseract
import tempfile
import os
import io
import time
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def compress_image(image, quality=60):
    """Nén ảnh để giảm dung lượng"""
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

        # Chèn ảnh trang PDF
        doc.add_picture(img_byte, width=Inches(6.8))

        # OCR text từ ảnh
        text = pytesseract.image_to_string(img, lang=lang).strip()
        if text:
            doc.add_paragraph("\n" + text)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        mode = request.form.get("mode")
        dpi = int(request.form.get("dpi", 150))
        quality = int(request.form.get("quality", 60))
        max_pages = int(request.form.get("max_pages", 0)) or None
        lang = request.form.get("lang", "eng")

        file = request.files.get("pdf_file")
        if not file or not file.filename.lower().endswith(".pdf"):
            return "Vui lòng chọn file PDF hợp lệ."

        filename = secure_filename(file.filename)
        pdf_bytes = file.read()

        start = time.time()
        if mode == "visual":
            output_stream = pdf_to_word_visual(pdf_bytes, dpi, quality, max_pages)
            out_name = filename.replace(".pdf", "_visual.docx")
        elif mode == "text":
            output_stream = pdf_to_word_text(pdf_bytes, max_pages)
            out_name = filename.replace(".pdf", "_text.docx")
        elif mode == "hybrid":
            output_stream = pdf_to_word_hybrid(pdf_bytes, dpi, quality, lang, max_pages)
            out_name = filename.replace(".pdf", "_hybrid.docx")
        else:
            return "Chế độ không hợp lệ."

        print(f"Convert xong trong {time.time()-start:.1f}s")
        return send_file(output_stream,
                         as_attachment=True,
                         download_name=out_name,
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
