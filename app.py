from flask import Flask, request, render_template, flash, redirect, url_for, send_file
import io
import base64
import zipfile
import os

# Import hàm điều phối duy nhất từ logic_handler
from logic_handler import process_unified_file

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'a_very_strong_and_unified_secret_key')

def get_chxd_list():
    """Đọc danh sách CHXD từ file text để hiển thị trên giao diện."""
    try:
        with open("DS_CHXD.txt", "r", encoding="utf-8") as f:
            return [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        flash("Lỗi nghiêm trọng: Không tìm thấy file DS_CHXD.txt!", "danger")
        return []
    except Exception as e:
        flash(f"Lỗi khi đọc file DS_CHXD.txt: {e}", "danger")
        return []

@app.route('/', methods=['GET'])
def index():
    """Hiển thị trang upload chính."""
    chxd_list = get_chxd_list()
    # Truyền vào một dict rỗng để template không bị lỗi ở lần tải đầu tiên
    return render_template('index.html', chxd_list=chxd_list, form_data={})

@app.route('/process', methods=['POST'])
def process():
    """Xử lý file tải lên bằng logic hợp nhất."""
    # Lấy danh sách CHXD để có thể render lại trang nếu có lỗi
    chxd_list = get_chxd_list()
    try:
        form_data = {
            "selected_chxd": request.form.get('chxd'),
            "price_periods": request.form.get('price_periods'),
            "invoice_number": request.form.get('invoice_number', '').strip(),
            "confirmed_date": request.form.get('confirmed_date'),
            "encoded_file": request.form.get('file_content_b64')
        }

        if not form_data["selected_chxd"]:
            flash('Vui lòng chọn CHXD.', 'warning')
            # Khi redirect, không cần truyền lại chxd_list vì route '/' sẽ tự lấy
            return redirect(url_for('index'))

        file_content = None
        if form_data["encoded_file"]:
            file_content = base64.b64decode(form_data["encoded_file"])
        elif 'file' in request.files and request.files['file'].filename != '':
            file_content = request.files['file'].read()
        else:
            flash('Vui lòng tải lên file Bảng kê.', 'warning')
            return redirect(url_for('index'))

        # Gọi hàm điều phối duy nhất
        result = process_unified_file(
            file_content=file_content, 
            selected_chxd=form_data["selected_chxd"],
            price_periods=form_data["price_periods"],
            new_price_invoice_number=form_data["invoice_number"],
            confirmed_date_str=form_data["confirmed_date"]
        )
        
        # Xử lý kết quả trả về
        if isinstance(result, dict) and result.get('choice_needed'):
            # SỬA LỖI: Lưu lại file để dùng cho lần submit sau
            form_data["encoded_file"] = base64.b64encode(file_content).decode('utf-8')
            # Truyền lại tất cả dữ liệu đã có của form để giữ lựa chọn
            return render_template('index.html', chxd_list=chxd_list, date_ambiguous=True, date_options=result['options'], form_data=form_data)
        
        elif isinstance(result, dict) and 'old' in result:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                if result.get('old'):
                    result['old'].seek(0)
                    zipf.writestr('UpSSE_gia_cu.xlsx', result['old'].read())
                if result.get('new'):
                    result['new'].seek(0)
                    zipf.writestr('UpSSE_gia_moi.xlsx', result['new'].read())
            zip_buffer.seek(0)
            return send_file(zip_buffer, as_attachment=True, download_name='UpSSE_2_giai_doan.zip', mimetype='application/zip')

        elif isinstance(result, io.BytesIO):
            result.seek(0)
            return send_file(result, as_attachment=True, download_name='UpSSE.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        else:
            raise ValueError("Hàm xử lý không trả về kết quả hợp lệ.")

    except (ValueError, NotImplementedError) as ve:
        flash(str(ve), 'danger')
        # SỬA LỖI: Render lại trang với các lựa chọn đã có thay vì redirect
        return render_template('index.html', chxd_list=chxd_list, form_data=request.form)
    except Exception as e:
        flash(f"Đã xảy ra lỗi không mong muốn: {e}", 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data=request.form)

# --- DÒNG QUAN TRỌNG ĐỂ KHỞI ĐỘNG SERVER ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
