import streamlit as st
import tempfile
import os
import subprocess
from marker.convert import convert_single_pdf
from PIL import Image
import pytesseract
import pdfplumber
import io

# Set path cho Tesseract
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

st.title("Chuyển PDF Sang Word - Hỗ Trợ Công Thức Toán Dạng Ảnh Chụp")

# Tải lên file PDF
uploaded_file = st.file_uploader("Tải lên file PDF", type="pdf")

if uploaded_file is not None:
    try:
        with st.spinner("Đang xử lý PDF..."):
            # Tạo thư mục tạm
            with tempfile.TemporaryDirectory() as temp_dir:
                pdf_path = os.path.join(temp_dir, "input.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(uploaded_file.getvalue())

                # Chuyển PDF sang markdown với marker (OCR toàn bộ, hỗ trợ math)
                full_text, images, out_meta = convert_single_pdf(pdf_path, ocr_all_pages=True)

                # Fallback OCR bằng pytesseract nếu marker không trích xuất được
                if not full_text.strip():
                    fallback_text = ""
                    with pdfplumber.open(pdf_path) as pdf:
                        for page in pdf.pages:
                            if not page.extract_text():
                                page_image = page.to_image(resolution=150).original
                                ocr_text = pytesseract.image_to_string(page_image, lang='eng+vie')
                                fallback_text += ocr_text + "\n"
                    full_text = fallback_text

                # Lưu markdown tạm
                md_path = os.path.join(temp_dir, "output.md")
                with open(md_path, "w", encoding="utf-8") as f:
                    f.write(full_text)

                # Chuyển markdown sang docx bằng pandoc
                docx_path = os.path.join(temp_dir, "output.docx")
                subprocess.run(
                    ["pandoc", md_path, "-o", docx_path, "--from=markdown+tex_math_dollars", "--to=docx"],
                    check=True
                )

                # Tải xuống file Word
                with open(docx_path, "rb") as f:
                    st.download_button(
                        label="Tải File Word (.docx)",
                        data=f,
                        file_name="converted.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

        st.success("Chuyển đổi hoàn tất! Công thức toán dạng ảnh đã được OCR và chuyển sang equation trong Word.")

    except Exception as e:
        st.error(f"Lỗi khi xử lý: {e}")
        st.info("Kiểm tra logs trên Streamlit Cloud hoặc đảm bảo requirements và packages đã cập nhật.")
