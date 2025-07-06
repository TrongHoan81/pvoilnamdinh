from flask import Flask, request, render_template, flash, redirect, url_for, send_file
import io
import base64
import zipfile
import os
# Thêm thư viện openpyxl để đọc file Excel
from openpyxl import load_workbook

# Import hàm điều phối duy nhất từ logic_handler
from logic_handler import process_unified_file

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'a_very_strong_and_unified_secret_key')

def get_chxd_list():
    """
    CẢI TIẾN: Đọc danh sách CHXD trực tiếp từ cột D của file Data_HDDT.xlsx
    thay vì từ file DS_CHXD.txt.
    """
    chxd_list = []
    try:
        # Mở file cấu hình chính
        wb = load_workbook("Data_HDDT.xlsx", data_only=True)
        ws = wb.active
        # Duyệt qua cột D (cột thứ 4) từ dòng 3 để lấy tên các CHXD
        for row in ws.iter_rows(min_row=3, min_col=4, max_col=4, values_only=True):
            chxd_name = row[0]
            if chxd_name and isinstance(chxd_name, str) and chxd_name.strip():
                chxd_list.append(chxd_name.strip())
        
        # Sắp xếp lại danh sách theo thứ tự alphabet để dễ nhìn
        chxd_list.sort()
        return chxd_list
        
    except FileNotFoundError:
        flash("Lỗi nghiêm trọng: Không tìm thấy file cấu hình Data_HDDT.xlsx!", "danger")
        return [] # Trả về danh sách rỗng nếu file không tồn tại
    except Exception as e:
        flash(f"Lỗi khi đọc file Data_HDDT.xlsx: {e}", "danger")
        return []

@app.route('/', methods=['GET'])
def index():
    """Hiển thị trang upload chính."""
    chxd_list = get_chxd_list()
    return render_template('index.html', chxd_list=chxd_list, form_data={})

@app.route('/process', methods=['POST'])
def process():
    """Xử lý file tải lên bằng logic hợp nhất."""
    chxd_list = get_chxd_list()
    form_data = {
        "selected_chxd": request.form.get('chxd'),
        "price_periods": request.form.get('price_periods', '1'),
        "invoice_number": request.form.get('invoice_number', '').strip(),
        "confirmed_date": request.form.get('confirmed_date'),
        "encoded_file": request.form.get('file_content_b64')
    }
    
    try:
        if not form_data["selected_chxd"]:
            flash('Vui lòng chọn CHXD.', 'warning')
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
            form_data["encoded_file"] = base64.b64encode(file_content).decode('utf-8')
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

    except ValueError as ve:
        flash(str(ve).replace('\n', '<br>'), 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data=form_data)
    except Exception as e:
        flash(f"Đã xảy ra lỗi không mong muốn: {e}", 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data=form_data)

# --- DÒNG QUAN TRỌNG ĐỂ KHỞI ĐỘNG SERVER ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
