import os
import uuid
import zipfile
from flask import (Flask, request, render_template, send_from_directory,
                   flash, redirect, url_for)
from werkzeug.utils import secure_filename
from pdf2docx import Converter
from pdf2image import convert_from_path

# --- CẤU HÌNH ---
# THAY ĐỔI ĐƯỜNG DẪN NÀY tới thư mục bin của Poppler bạn đã tải về.
# Ví dụ trên Windows: r'C:\poppler-23.11.0\Library\bin'
# Trên macOS/Linux, bạn có thể để là None nếu đã cài qua brew/apt
POPPLER_PATH = r'C:\path\to\your\poppler\bin' # <<< THAY ĐỔI DÒNG NÀY

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__)
app.config['SECRET_KEY'] = 'mot-chuoi-bi-mat-sieu-dai-sieu-kho-doan' # Cần cho flash message
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Tạo các thư mục cần thiết nếu chúng chưa tồn tại
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --- CÁC HÀM HỖ TRỢ ---

def allowed_file(filename):
    """Kiểm tra xem file có phần mở rộng là .pdf hay không"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup_files(folder_path):
    """Hàm dọn dẹp file (chưa được sử dụng trong ví dụ này nhưng hữu ích)"""
    # Bạn có thể triển khai một cơ chế dọn dẹp các file cũ ở đây
    # Ví dụ: xóa các thư mục đã tạo quá 1 giờ
    pass

# --- CÁC ROUTE (ĐIỀU HƯỚNG) CỦA ỨNG DỤNG ---

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 1. Kiểm tra file có được gửi lên không
        if 'file' not in request.files:
            flash('Không có phần tệp nào trong yêu cầu.', 'danger')
            return redirect(request.url)
        
        file = request.files['file']

        # 2. Kiểm tra người dùng có chọn file không
        if file.filename == '':
            flash('Chưa có tệp nào được chọn.', 'warning')
            return redirect(request.url)

        # 3. Kiểm tra file hợp lệ và lưu lại
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            
            # Tạo một thư mục duy nhất cho mỗi lần xử lý để tránh xung đột file
            job_id = str(uuid.uuid4())
            input_dir = os.path.join(app.config['UPLOAD_FOLDER'], job_id)
            output_dir = os.path.join(app.config['OUTPUT_FOLDER'], job_id)
            os.makedirs(input_dir)
            os.makedirs(output_dir)

            pdf_path = os.path.join(input_dir, filename)
            file.save(pdf_path)

            conversion_type = request.form.get('conversion_type')

            try:
                # --- XỬ LÝ CHUYỂN ĐỔI ---
                if conversion_type == 'to_docx':
                    output_filename = f"{os.path.splitext(filename)[0]}.docx"
                    output_path = os.path.join(output_dir, output_filename)
                    
                    # Chuyển đổi PDF sang DOCX
                    cv = Converter(pdf_path)
                    cv.convert(output_path, start=0, end=None)
                    cv.close()
                    
                    return redirect(url_for('download_file', job_id=job_id, filename=output_filename))

                elif conversion_type == 'to_jpg':
                    # Chuyển đổi PDF sang ảnh (mỗi trang một ảnh)
                    images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
                    
                    if not images:
                        flash('Không thể trích xuất ảnh từ PDF.', 'danger')
                        return redirect(request.url)

                    # Nếu chỉ có 1 trang, trả về ảnh trực tiếp
                    if len(images) == 1:
                        output_filename = 'page_1.jpg'
                        output_path = os.path.join(output_dir, output_filename)
                        images[0].save(output_path, 'JPEG')
                        return redirect(url_for('download_file', job_id=job_id, filename=output_filename))
                    
                    # Nếu có nhiều trang, nén chúng thành file ZIP
                    else:
                        zip_filename = f"{os.path.splitext(filename)[0]}.zip"
                        zip_path = os.path.join(output_dir, zip_filename)
                        
                        with zipfile.ZipFile(zip_path, 'w') as zipf:
                            for i, image in enumerate(images):
                                image_filename = f'page_{i+1}.jpg'
                                image_path = os.path.join(output_dir, image_filename)
                                image.save(image_path, 'JPEG')
                                zipf.write(image_path, arcname=image_filename)

                        return redirect(url_for('download_file', job_id=job_id, filename=zip_filename))

            except Exception as e:
                # Ghi lại lỗi để debug
                print(f"Đã xảy ra lỗi: {e}")
                # Thông báo lỗi cho người dùng
                flash(f'Đã có lỗi xảy ra trong quá trình chuyển đổi. Vui lòng thử lại. Lỗi: {e}', 'danger')
                return redirect(request.url)
        else:
            flash('Loại tệp không được phép. Vui lòng chỉ tải lên file .pdf', 'danger')
            return redirect(request.url)

    return render_template('index.html')


@app.route('/download/<job_id>/<filename>')
def download_file(job_id, filename):
    """Cung cấp file cho người dùng tải về"""
    directory = os.path.join(app.config['OUTPUT_FOLDER'], job_id)
    try:
        return send_from_directory(directory, filename, as_attachment=True)
    except FileNotFoundError:
        flash('Không tìm thấy tệp để tải xuống.', 'danger')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
