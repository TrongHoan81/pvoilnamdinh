import base64
import io
import os
import zipfile
from flask import Flask, flash, redirect, render_template, request, send_file, url_for, get_flashed_messages, jsonify, session, after_this_request
from openpyxl import load_workbook
import re 
import pandas as pd 
from collections import defaultdict 
import uuid 

# Import the comparison logic from your handler file
from doisoatthue_handler import compare_invoices 

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'a_very_strong_and_unified_secret_key_for_doisoatthue_v2') 

UPLOAD_FOLDER = 'temp_uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- CUSTOM JINJA2 FILTER ---
@app.template_filter('format_currency')
def format_currency_filter(value):
    """
    Định dạng số thành chuỗi tiền tệ có dấu phẩy phân cách hàng nghìn.
    """
    try:
        num = float(value)
        return f"{num:,.0f}" 
    except (ValueError, TypeError):
        return "0" 

# Route for the home page
@app.route('/')
def index():
    """
    Renders the main HTML page.
    """
    return render_template('index.html')

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

    if not (tax_invoice_file_obj and e_invoice_file_obj and tax_invoice_file_obj.filename and e_invoice_file_obj.filename):
        return jsonify({'message': 'Vui lòng tải lên cả hai file bảng kê.'}), 400

    if not tax_invoice_file_obj.filename.lower().endswith(('.xlsx', '.xls')) or \
       not e_invoice_file_obj.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'message': 'Chỉ chấp nhận định dạng file Excel (.xlsx, .xls).'}), 400
    
    tax_file_stream = io.BytesIO(tax_invoice_file_obj.read())
    e_invoice_file_stream = io.BytesIO(e_invoice_file_obj.read())

    try:
        # Nhận tất cả các kết quả từ hàm xử lý
        comparison_summary, output_excel_stream, overall_summary, item_summary_data = compare_invoices(
            tax_file_stream, e_invoice_file_stream
        )

        download_url = None
        if output_excel_stream:
            unique_filename = f"ket_qua_doi_soat_{uuid.uuid4().hex}.xlsx"
            file_path = os.path.join(UPLOAD_FOLDER, unique_filename)
            with open(file_path, 'wb') as f:
                f.write(output_excel_stream.getvalue())
            
            download_url = f"/download_results/{unique_filename}"

        # Gộp tất cả kết quả vào một đối tượng JSON để trả về
        full_results = {
            **comparison_summary, # Gồm matched_count, mismatched_invoices
            'download_url': download_url,
            'overall_summary': overall_summary,
            'item_summary': item_summary_data
        }
        
        return jsonify(full_results)

    except ValueError as ve:
        return jsonify({'message': str(ve)}), 400
    except Exception as e:
        app.logger.error(f"Error processing files: {e}", exc_info=True)
        return jsonify({'message': f'Có lỗi xảy ra trong quá trình xử lý: {e}'}), 500

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
            except Exception as e:
                app.logger.error(f"Error deleting temporary file {file_path}: {e}", exc_info=True)
            return response

        return send_file(file_path, as_attachment=True, download_name=f"ket_qua_doi_soat.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return jsonify({'message': 'File không tồn tại hoặc đã bị xóa.'}), 404

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
