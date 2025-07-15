import base64
import io
import os
import zipfile
from flask import Flask, flash, redirect, render_template, request, send_file, url_for, get_flashed_messages, jsonify, session, after_this_request # Import after_this_request
from openpyxl import load_workbook
import re 
import pandas as pd 
from collections import defaultdict 
import uuid 

# Import the comparison logic from your handler file
from doisoatthue_handler import compare_invoices 

app = Flask(__name__)
# Configure a secret key for Flask sessions (needed for flash messages, etc., though not used directly in this example)
# IMPORTANT: Change this to a strong, random key in production!
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'a_very_strong_and_unified_secret_key_for_doisoatthue') 

# Directory to store temporary files (e.g., comparison results)
UPLOAD_FOLDER = 'temp_uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- CUSTOM JINJA2 FILTER ---
@app.template_filter('format_currency')
def format_currency_filter(value):
    """
    Định dạng số thành chuỗi tiền tệ có dấu phẩy phân cách hàng nghìn.
    Sử dụng cho hiển thị trong template Jinja2.
    """
    try:
        num = float(value)
        return f"{num:,.0f}" 
    except (ValueError, TypeError):
        return "0" 

# --- HÀM TIỆN ÍCH CHO VIỆC NẠP DỮ LIỆU CẤU HÌNH (Nếu có) ---
def load_all_static_config_data():
    """
    Placeholder for loading static configuration data if needed in the future.
    For this specific reconciliation app, no external static config files are used.
    """
    return {}, None 

# Load static data once when the app starts (if any)
_global_static_config_data, _static_config_error = load_all_static_config_data()
if _static_config_error:
    print(f"Error loading static configuration data: {_static_config_error}")

# Route for the home page (serves the index.html)
@app.route('/')
def index():
    """
    Renders the main HTML page for invoice comparison.
    """
    # Default values for initial load
    form_data = {"active_tab": "doisoat"} # Default to doisoat tab
    return render_template('index.html', 
                           form_data=form_data)

# Route to handle file uploads and comparison
@app.route('/compare_invoices', methods=['POST'])
def compare_invoices_route():
    """
    Handles the POST request for invoice comparison.
    Receives two Excel files, calls the comparison logic,
    and returns a JSON summary and a downloadable Excel file for mismatches.
    """
    tax_invoice_file_obj = request.files.get('tax_invoice_file')
    e_invoice_file_obj = request.files.get('e_invoice_file')

    tax_file_stream = None
    e_invoice_file_stream = None
    
    if tax_invoice_file_obj and e_invoice_file_obj and tax_invoice_file_obj.filename != '' and e_invoice_file_obj.filename != '':
        if not tax_invoice_file_obj.filename.lower().endswith(('.xlsx', '.xls')) or \
           not e_invoice_file_obj.filename.lower().endswith(('.xlsx', '.xls')):
            flash('Chỉ chấp nhận định dạng file Excel (.xlsx, .xls).', 'warning')
            return jsonify({'message': 'Chỉ chấp nhận định dạng file Excel (.xlsx, .xls).', 'redirect_url': url_for('index', active_tab='doisoat')}), 400
        
        tax_file_stream = io.BytesIO(tax_invoice_file_obj.read())
        e_invoice_file_stream = io.BytesIO(e_invoice_file_obj.read())

    else:
        flash('Vui lòng tải lên cả hai file bảng kê.', 'warning')
        return jsonify({'message': 'Vui lòng tải lên cả hai file bảng kê.', 'redirect_url': url_for('index', active_tab='doisoat')}), 400

    try:
        comparison_summary, output_excel_stream, _ = compare_invoices(
            tax_file_stream, e_invoice_file_stream
        )

        download_url = None
        if output_excel_stream:
            unique_filename = f"ket_qua_doi_soat_{uuid.uuid4().hex}.xlsx"
            file_path = os.path.join(UPLOAD_FOLDER, unique_filename)
            with open(file_path, 'wb') as f:
                f.write(output_excel_stream.getvalue())
            
            download_url = f"/download_results/{unique_filename}"

        comparison_summary['download_url'] = download_url
        flash('Đối soát thành công!', 'success')
        
        return jsonify(comparison_summary) # Return JSON for successful AJAX submission

    except ValueError as ve:
        flash(str(ve).replace('\n', '<br>'), 'danger')
        return jsonify({'message': str(ve), 'redirect_url': url_for('index', active_tab='doisoat')}), 400
    except Exception as e:
        app.logger.error(f"Error processing files: {e}", exc_info=True)
        flash(f'Có lỗi xảy ra trong quá trình xử lý: {e}', 'danger')
        return jsonify({'message': f'Có lỗi xảy ra trong quá trình xử lý: {e}', 'redirect_url': url_for('index', active_tab='doisoat')}), 500

# Route to serve the generated Excel file
@app.route('/download_results/<filename>')
def download_results(filename):
    """
    Serves the generated Excel file containing mismatch details.
    """
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.exists(file_path):
        @after_this_request
        def remove_file(response):
            try:
                os.remove(file_path)
                app.logger.info(f"Deleted temporary file: {file_path}")
            except Exception as e:
                app.logger.error(f"Error deleting temporary file {file_path}: {e}", exc_info=True)
            return response

        return send_file(file_path, as_attachment=True, download_name=f"ket_qua_doi_soat.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return jsonify({'message': 'File không tồn tại hoặc đã bị xóa.'}), 404

@app.route('/clear_flash_messages', methods=['GET'])
def clear_flash_messages():
    """Route này được gọi bởi JavaScript để xóa các thông báo flash trong session."""
    _ = get_flashed_messages()
    return '', 204

if __name__ == '__main__':
    # Ensure the upload folder exists when the app starts
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    # app.run(debug=True) # debug=True is for development, set to False for production - Dòng này đã được loại bỏ
