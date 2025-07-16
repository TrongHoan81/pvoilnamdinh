import pandas as pd
import io
from datetime import datetime
import re
import numpy as np

# --- HÀM TIỆN ÍCH CHO XỬ LÝ NGÀY THÁNG ---
def _excel_date_to_datetime_robust(excel_date):
    """
    Chuyển đổi ngày tháng từ định dạng Excel sang đối tượng datetime.
    Hỗ trợ các định dạng số Excel, datetime object, và chuỗi.
    """
    if pd.isna(excel_date):
        return None
    
    if isinstance(excel_date, datetime):
        return excel_date
    
    if isinstance(excel_date, (int, float)):
        try:
            return pd.to_datetime(excel_date, unit='D', origin='1899-12-30').to_pydatetime()
        except Exception:
            pass
    
    if isinstance(excel_date, str):
        cleaned_str = excel_date.strip().replace('\xa0', ' ')
        formats_to_try = [
            '%d/%m/%Y %H:%M:%S', '%d/%m/%Y',
            '%m/%d/%Y %H:%M:%S', '%m/%d/%Y',
            '%Y-%m-%d',
            '%Y/%m/%d',
        ]
        for fmt in formats_to_try:
            try:
                return datetime.strptime(cleaned_str, fmt)
            except ValueError:
                continue
    return None

# --- HÀM TIỆN ÍCH ĐỂ LÀM SẠCH CHUỖI SỐ NGUYÊN ---
def _clean_numeric_string_for_int(s):
    """
    Làm sạch chuỗi số để chuyển đổi thành số nguyên chính xác.
    Xử lý định dạng '740,741.0' và '740741' thành '740741'.
    """
    if pd.isna(s):
        return None
    s_str = str(s).strip()
    s_str = s_str.replace(',', '')
    if s_str.endswith('.0'):
        s_str = s_str[:-2]
    cleaned_s = re.sub(r'[^\d-]', '', s_str)
    if not cleaned_s or cleaned_s == '-':
        return None
    try:
        return int(float(cleaned_s))
    except (ValueError, TypeError):
        print(f"WARNING: Could not convert '{s}' to integer after cleaning. Cleaned string: '{cleaned_s}'")
        return None

# --- HÀM TIỆN ÍCH TẠO CHỨNG MINH THƯ HÓA ĐƠN ---
def _create_invoice_identity(row, template_col, series_col, number_col):
    """
    Tạo 'Chứng minh thư' duy nhất cho hóa đơn bằng cách ghép Mẫu số, Ký hiệu, Số hóa đơn.
    """
    template = re.sub(r'\s+', '', str(row[template_col]).strip()).upper() if pd.notna(row[template_col]) else ''
    series = re.sub(r'\s+', '', str(row[series_col]).strip()).upper() if pd.notna(row[series_col]) else ''
    number = str(row[number_col]).strip() if pd.notna(row[number_col]) else ''
    try:
        number = str(int(float(number)))
    except (ValueError, TypeError):
        pass
    return f"{template}{series}{number}" if template and series and number else None


# --- HÀM CHÍNH THỰC HIỆN ĐỐI SOÁT HÓA ĐƠN ---
def compare_invoices(tax_invoice_file_stream, e_invoice_file_stream):
    """
    Thực hiện đối soát hóa đơn giữa dữ liệu của cơ quan thuế và dữ liệu hóa đơn điện tử.
    """
    try:
        # --- Đọc file Excel vào DataFrame ---
        df_tax = pd.read_excel(tax_invoice_file_stream, header=5, engine='openpyxl')
        df_e_invoice = pd.read_excel(e_invoice_file_stream, header=8, engine='openpyxl')

        if not df_e_invoice.empty:
            is_col_number_row = df_e_invoice.iloc[0].astype(str).str.fullmatch(r'\[\d+\]').any()
            if is_col_number_row:
                df_e_invoice = df_e_invoice.iloc[1:].reset_index(drop=True)

        # --- Chuẩn hóa tên cột ---
        tax_col_mapping = {
            'Ký hiệu mẫu số': 'invoice_template_symbol_tax', 'Ký hiệu hóa đơn': 'invoice_series_symbol_tax', 'Số hóa đơn': 'invoice_number_tax',
            'Tên người mua/Tên người nhận hàng': 'buyer_name_tax', 'MST người mua/MST người nhận hàng': 'buyer_tax_id_tax',
            'Tổng tiền chưa thuế': 'sub_total_amount_tax', 'Tổng tiền thuế': 'tax_amount_tax', 'Tổng tiền thanh toán': 'total_amount_tax',
            'Kết quả kiểm tra hóa đơn': 'invoice_check_result_tax', 'Trạng thái hóa đơn': 'invoice_status_tax',
        }
        df_tax = df_tax.rename(columns=tax_col_mapping)

        e_invoice_col_mapping = {
            'Mẫu số': 'invoice_template_symbol_e_inv', 'Ký hiệu': 'invoice_series_symbol_e_inv', 'Số hóa đơn': 'invoice_number_e_inv',
            'Tên khách hàng': 'customer_name_e_inv', 'MST khách hàng': 'buyer_tax_id_e_inv',
            'Thành tiền': 'sub_total_amount_e_inv', 'Tiền thuế': 'tax_amount_e_inv', 'Phải thu': 'total_amount_e_inv',
            'Fkey': 'fkey_e_inv', 'Trạng thái': 'status_e_inv', 'Mặt hàng': 'item_description_e_inv',
        }
        df_e_invoice = df_e_invoice.rename(columns=e_invoice_col_mapping)

        # --- Tính toán Tóm tắt Tổng thể ---
        if 'status_e_inv' in df_e_invoice.columns:
            df_e_invoice_published = df_e_invoice[df_e_invoice['status_e_inv'].astype(str).str.strip().str.lower() == 'đã phát hành'].copy()
        else:
            df_e_invoice_published = df_e_invoice.copy()
        total_e_invoices_published = len(df_e_invoice_published)

        total_tax_invoices_accepted = 0
        if 'invoice_check_result_tax' in df_tax.columns:
            total_tax_invoices_accepted = len(df_tax[df_tax['invoice_check_result_tax'].astype(str).str.strip() == 'Tổng cục thuế đã nhận không mã'])

        overall_summary = {
            'total_e_invoices_published': total_e_invoices_published,
            'total_tax_invoices_accepted': total_tax_invoices_accepted,
        }

        # --- Tạo cột 'Chứng minh thư' ---
        df_tax['invoice_identity'] = df_tax.apply(lambda row: _create_invoice_identity(row, 'invoice_template_symbol_tax', 'invoice_series_symbol_tax', 'invoice_number_tax'), axis=1)
        df_e_invoice_published['invoice_identity'] = df_e_invoice_published.apply(lambda row: _create_invoice_identity(row, 'invoice_template_symbol_e_inv', 'invoice_series_symbol_e_inv', 'invoice_number_e_inv'), axis=1)
        
        df_tax.dropna(subset=['invoice_identity'], inplace=True)
        df_e_invoice_published.dropna(subset=['invoice_identity'], inplace=True)

        # --- Hợp nhất DataFrames ---
        merged_df = pd.merge(df_e_invoice_published, df_tax, on='invoice_identity', how='outer', suffixes=('_e_inv', '_tax'), indicator=True)

        # --- Xử lý các cột số liệu để tính toán ---
        numeric_cols = ['sub_total_amount_e_inv', 'tax_amount_e_inv', 'total_amount_e_inv', 'sub_total_amount_tax', 'tax_amount_tax', 'total_amount_tax']
        for col in numeric_cols:
            if col in merged_df.columns:
                merged_df[f"{col}_num"] = merged_df[col].apply(_clean_numeric_string_for_int).fillna(0)
            else:
                merged_df[f"{col}_num"] = 0
        
        # --- Tính toán Bảng so sánh theo Mặt hàng ---
        item_summary_data = []
        if 'item_description_e_inv' in merged_df.columns:
            merged_df['item_group'] = merged_df['item_description_e_inv'].fillna('Không xác định (chỉ có trên bảng kê Thuế)')
            grouped_by_item = merged_df.groupby('item_group')
            item_summary_list = []
            for item_name, group in grouped_by_item:
                e_inv_count = len(group[group['_merge'] != 'right_only'])
                tax_count = len(group[group['_merge'] != 'left_only'])
                
                # Tính tổng
                e_inv_sub_total = group['sub_total_amount_e_inv_num'].sum()
                tax_sub_total = group['sub_total_amount_tax_num'].sum()
                e_inv_tax = group['tax_amount_e_inv_num'].sum()
                tax_tax = group['tax_amount_tax_num'].sum()
                e_inv_total = group['total_amount_e_inv_num'].sum()
                tax_total = group['total_amount_tax_num'].sum()

                # *** SỬA LỖI: Chuyển đổi tất cả các giá trị sang kiểu int của Python ***
                item_summary_list.append({
                    'item_name': item_name,
                    'metrics': [
                        {'name': 'Số lượng hóa đơn', 'e_inv_val': int(e_inv_count), 'tax_val': int(tax_count), 'diff': int(e_inv_count - tax_count)},
                        {'name': 'Tổng tiền chưa thuế', 'e_inv_val': int(e_inv_sub_total), 'tax_val': int(tax_sub_total), 'diff': int(e_inv_sub_total - tax_sub_total)},
                        {'name': 'Tổng tiền thuế', 'e_inv_val': int(e_inv_tax), 'tax_val': int(tax_tax), 'diff': int(e_inv_tax - tax_tax)},
                        {'name': 'Tổng tiền thanh toán', 'e_inv_val': int(e_inv_total), 'tax_val': int(tax_total), 'diff': int(e_inv_total - tax_total)},
                    ]
                })
            item_summary_data = sorted(item_summary_list, key=lambda x: x['item_name'])

        # --- Phát hiện sai lệch chi tiết ---
        merged_df['Lý do sai lệch chi tiết'] = ''
        merged_df.loc[merged_df['_merge'] == 'left_only', 'Lý do sai lệch chi tiết'] = 'Chưa được đẩy lên Cơ quan Thuế'
        merged_df.loc[merged_df['_merge'] == 'right_only', 'Lý do sai lệch chi tiết'] = 'Không có trong Bảng kê HĐĐT đã phát hành'

        both_present_df_indices = merged_df[merged_df['_merge'] == 'both'].index
        comparison_fields = [
            ('customer_name_e_inv', 'buyer_name_tax', 'Tên khách hàng'), ('buyer_tax_id_e_inv', 'buyer_tax_id_tax', 'MST khách hàng'),
            ('sub_total_amount_e_inv_num', 'sub_total_amount_tax_num', 'Tổng tiền chưa thuế'), ('tax_amount_e_inv_num', 'tax_amount_tax_num', 'Tiền thuế'),
            ('total_amount_e_inv_num', 'total_amount_tax_num', 'Tổng tiền thanh toán'),
        ]
        for e_inv_col, tax_col, display_name in comparison_fields:
            if tax_col in merged_df.columns and e_inv_col in merged_df.columns:
                tax_values = merged_df.loc[both_present_df_indices, tax_col]
                e_inv_values = merged_df.loc[both_present_df_indices, e_inv_col]
                if 'Tên' in display_name or 'MST' in display_name:
                    mismatch_mask = (tax_values.astype(str).str.strip().str.lower() != e_inv_values.astype(str).str.strip().str.lower())
                else:
                    mismatch_mask = (tax_values != e_inv_values)
                
                mismatch_indices = both_present_df_indices[mismatch_mask]
                merged_df.loc[mismatch_indices, 'Lý do sai lệch chi tiết'] = merged_df.loc[mismatch_indices, 'Lý do sai lệch chi tiết'].apply(
                    lambda x: (x + f'; Sai lệch {display_name}') if x else f'Sai lệch {display_name}'
                )

        mismatched_details_df = merged_df[merged_df['Lý do sai lệch chi tiết'].str.strip() != ''].copy()
        
        all_mismatched_invoice_numbers_for_summary = []
        if not mismatched_details_df.empty:
            for index, row in mismatched_details_df.iterrows():
                all_mismatched_invoice_numbers_for_summary.append(f"Số HĐ: {row.get('invoice_number_e_inv', 'N/A')} (CMT: {row['invoice_identity']}) - Lý do: {row['Lý do sai lệch chi tiết']}")
        
        # *** SỬA LỖI: Chuyển đổi sang kiểu int của Python ***
        matched_count = int(len(merged_df[(merged_df['_merge'] == 'both') & (merged_df['Lý do sai lệch chi tiết'] == '')]))
        
        comparison_summary = {
            'matched_count': matched_count,
            'mismatched_invoices': all_mismatched_invoice_numbers_for_summary
        }

        # --- Tạo file Excel kết quả chi tiết ---
        output_excel_stream = None
        if not mismatched_details_df.empty:
            output_excel_stream = io.BytesIO()
            full_display_col_mapping = {
                'invoice_identity': 'Số/ký hiệu hóa đơn bị sai lệch', 'customer_name_e_inv': 'Tên khách hàng (HĐĐT)', 'buyer_name_tax': 'Tên khách hàng (Thuế)',
                'buyer_tax_id_e_inv': 'Mã số thuế (HĐĐT)', 'buyer_tax_id_tax': 'Mã số thuế (Thuế)',
                'sub_total_amount_e_inv': 'Tổng tiền chưa thuế (HĐĐT)', 'sub_total_amount_tax': 'Tổng tiền chưa thuế (Thuế)',
                'tax_amount_e_inv': 'Tiền thuế (HĐĐT)', 'tax_amount_tax': 'Tiền thuế (Thuế)',
                'total_amount_e_inv': 'Tổng tiền thanh toán (HĐĐT)', 'total_amount_tax': 'Tổng tiền thanh toán (Thuế)',
                'fkey_e_inv': 'Mã FKEY', 'Lý do sai lệch chi tiết': 'Lý do sai lệch',
            }
            final_output_cols_order = list(full_display_col_mapping.keys())
            final_output_df = mismatched_details_df[[col for col in final_output_cols_order if col in mismatched_details_df.columns]].copy()
            final_output_df = final_output_df.rename(columns=full_display_col_mapping)
            final_output_df.to_excel(output_excel_stream, index=False, engine='openpyxl')
            output_excel_stream.seek(0)
        
        return comparison_summary, output_excel_stream, overall_summary, item_summary_data

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise ValueError(f"Lỗi nghiêm trọng trong quá trình đối soát: {e}")

