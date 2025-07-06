import base64
import io
import os
import zipfile
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from openpyxl import load_workbook

# --- CÁC IMPORT MỚI ---
from detector import detect_report_type
from hddt_handler import process_hddt_report
from pos_handler import process_pos_report

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'a_very_strong_and_unified_secret_key')

def get_chxd_list():
    """Đọc danh sách CHXD trực tiếp từ cột D của file Data_HDDT.xlsx."""
    chxd_list = []
    try:
        wb = load_workbook("Data_HDDT.xlsx", data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=3, min_col=4, max_col=4, values_only=True):
            chxd_name = row[0]
            if chxd_name and isinstance(chxd_name, str) and chxd_name.strip():
                chxd_list.append(chxd_name.strip())
        chxd_list.sort()
        return chxd_list
    except FileNotFoundError:
        flash("Lỗi nghiêm trọng: Không tìm thấy file cấu hình Data_HDDT.xlsx!", "danger")
        return []
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
    """Xử lý file tải lên bằng cách gọi handler tương ứng."""
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

        # --- LOGIC ĐIỀU PHỐI MỚI ---
        report_type = detect_report_type(file_content)
        result = None

        if report_type == 'POS':
            result = process_pos_report(
                file_content_bytes=file_content,
                selected_chxd=form_data["selected_chxd"],
                price_periods=form_data["price_periods"],
                new_price_invoice_number=form_data["invoice_number"]
            )
        elif report_type == 'HDDT':
            result = process_hddt_report(
                file_content_bytes=file_content,
                selected_chxd=form_data["selected_chxd"],
                price_periods=form_data["price_periods"],
                new_price_invoice_number=form_data["invoice_number"],
                confirmed_date_str=form_data["confirmed_date"]
            )
        else:
            raise ValueError("Không thể tự động nhận diện loại Bảng kê. Vui lòng kiểm tra lại file Excel bạn đã tải lên.")

        # --- Xử lý kết quả trả về (không đổi) ---
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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
