import streamlit as st
import os
import uuid
from pdf2docx import Converter

# --- CẤU HÌNH TRANG VÀ TIÊU ĐỀ ---
st.set_page_config(
    page_title="PDF sang Word",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="auto"
)

# --- CSS TÙY CHỈNH (TÙY CHỌN) ĐỂ GIAO DIỆN ĐẸP HƠN ---
st.markdown("""
<style>
    /* Chỉnh sửa kiểu nút */
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 15px 32px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 12px;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    /* Chỉnh sửa kiểu nút tải xuống */
    .stDownloadButton>button {
        background-color: #008CBA;
        color: white;
        border: none;
        padding: 15px 32px;
        text-align: center;
        font-size: 16px;
        border-radius: 12px;
        width: 100%;
    }
    .stDownloadButton>button:hover {
        background-color: #007B9E;
    }
</style>
""", unsafe_allow_html=True)


# --- TIÊU ĐỀ VÀ MÔ TẢ ỨNG DỤNG ---
st.title("Chuyển đổi PDF sang Word (DOCX)")
st.write("Tải lên tệp PDF của bạn để chuyển đổi nó thành một tài liệu Word có thể chỉnh sửa.")
st.markdown("---") # Đường kẻ ngang phân cách

# --- KHU VỰC TẢI FILE LÊN ---
uploaded_file = st.file_uploader(
    "1. Kéo và thả hoặc nhấn để chọn tệp PDF",
    type=["pdf"],
    help="Chỉ chấp nhận các tệp có định dạng .pdf"
)

# Kiểm tra xem người dùng đã tải file lên chưa
if uploaded_file is not None:
    # Lấy tên file gốc
    original_filename = uploaded_file.name
    st.info(f"📁 Tệp đã chọn: **{original_filename}**")

    # --- NÚT BẮT ĐẦU CHUYỂN ĐỔI ---
    if st.button("🚀 Bắt đầu chuyển đổi"):
        # Tạo thư mục tạm thời để lưu file, tránh xung đột
        temp_dir = "temp_files"
        os.makedirs(temp_dir, exist_ok=True)
        
        # Tạo một đường dẫn duy nhất cho file PDF tải lên
        pdf_path = os.path.join(temp_dir, f"{uuid.uuid4()}_{original_filename}")
        
        # Lưu file PDF tải lên vào máy chủ tạm thời
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Hiển thị thanh tiến trình
        with st.spinner('🧙‍♂️ Đang thực hiện phép thuật... Vui lòng chờ trong giây lát!'):
            try:
                # Tạo tên file DOCX đầu ra
                docx_filename = f"{os.path.splitext(original_filename)[0]}.docx"
                docx_path = os.path.join(temp_dir, docx_filename)

                # --- LÕI CHUYỂN ĐỔI ---
                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()

                # Đọc file DOCX đã được tạo vào bộ nhớ
                with open(docx_path, 'rb') as docx_file:
                    docx_bytes = docx_file.read()

                # Hiển thị thông báo thành công
                st.success("🎉 Chuyển đổi thành công!")
                st.balloons()

                # --- NÚT TẢI FILE XUỐNG ---
                st.download_button(
                    label="📥 Tải xuống tệp Word (.docx)",
                    data=docx_bytes,
                    file_name=docx_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                # Dọn dẹp file tạm sau khi hoàn tất
                os.remove(pdf_path)
                os.remove(docx_path)

            except Exception as e:
                st.error(f"❌ Đã xảy ra lỗi trong quá trình chuyển đổi:")
                st.error(e)

# --- Chân trang ---
st.markdown("---")
st.markdown("Được tạo bằng ❤️ với [Streamlit](https://streamlit.io).")
