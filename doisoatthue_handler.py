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
            # Excel's epoch is 1899-12-30 for dates, 1900-01-00 for numbers (Windows vs Mac)
            # pandas uses 1899-12-30 as default origin for 'D' unit
            return pd.to_datetime(excel_date, unit='D', origin='1899-12-30').to_pydatetime()
        except Exception:
            pass
    
    if isinstance(excel_date, str):
        cleaned_str = excel_date.strip().replace('\xa0', ' ') # Loại bỏ khoảng trắng thừa và non-breaking space
        formats_to_try = [
            '%d/%m/%Y %H:%M:%S', '%d/%m/%Y', # DD/MM/YYYY
            '%m/%d/%Y %H:%M:%S', '%m/%d/%Y', # MM/DD/YYYY
            '%Y-%m-%d', # YYYY-MM-DD
            '%Y/%m/%d', # YYYY/MM/DD
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
    
    # Bước 1: Loại bỏ tất cả dấu phẩy (phân cách hàng nghìn)
    s_str = s_str.replace(',', '')
    
    # Bước 2: Loại bỏ ".0" ở cuối nếu có (vì các giá trị luôn là số nguyên)
    if s_str.endswith('.0'):
        s_str = s_str[:-2]
    
    # Bước 3: Loại bỏ tất cả các ký tự không phải chữ số (ngoại trừ dấu trừ ở đầu)
    # Đảm bảo chỉ giữ lại các chữ số và dấu trừ nếu có
    cleaned_s = re.sub(r'[^\d-]', '', s_str)
    
    # Kiểm tra chuỗi sau khi làm sạch có rỗng hoặc chỉ là dấu trừ không
    if not cleaned_s or cleaned_s == '-':
        return None
    
    # Cố gắng chuyển đổi sang int để xác thực, sau đó trả về dưới dạng chuỗi
    try:
        return str(int(cleaned_s))
    except ValueError:
        # Nếu không thể chuyển đổi thành số nguyên, nó không phải là số hợp lệ
        print(f"WARNING: Could not convert '{s}' to integer after cleaning. Cleaned string: '{cleaned_s}'")
        return None

# --- HÀM TIỆN ÍCH TẠO CHỨNG MINH THƯ HÓA ĐƠN ---
def _create_invoice_identity(row, template_col, series_col, number_col):
    """
    Tạo 'Chứng minh thư' duy nhất cho hóa đơn bằng cách ghép Mẫu số, Ký hiệu, Số hóa đơn.
    Số hóa đơn sẽ được chuẩn hóa (loại bỏ số 0 ở đầu).
    """
    # Sử dụng re.sub để loại bỏ tất cả khoảng trắng (bao gồm cả khoảng trắng, tab, xuống dòng)
    template = re.sub(r'\s+', '', str(row[template_col]).strip()).upper() if pd.notna(row[template_col]) else ''
    series = re.sub(r'\s+', '', str(row[series_col]).strip()).upper() if pd.notna(row[series_col]) else ''
    number = str(row[number_col]).strip() if pd.notna(row[number_col]) else ''

    # Loại bỏ số 0 ở đầu của Số hóa đơn
    try:
        number = str(int(number))
    except ValueError:
        pass # Giữ nguyên nếu không phải số

    return f"{template}{series}{number}" if template and series and number else None


# --- HÀM CHÍNH THỰC HIỆN ĐỐI SOÁT HÓA ĐƠN ---
def compare_invoices(tax_invoice_file_stream, e_invoice_file_stream):
    """
    Thực hiện đối soát hóa đơn giữa dữ liệu của cơ quan thuế và dữ liệu hóa đơn điện tử.

    Args:
        tax_invoice_file_stream: Đối tượng file (BytesIO) chứa dữ liệu bảng kê thuế.
        e_invoice_file_stream: Đối tượng file (BytesIO) chứa dữ liệu bảng kê hóa đơn điện tử.

    Returns:
        Một tuple chứa:
        - dict: Một từ điển tóm tắt kết quả đối soát (matched_count, mismatched_invoices).
        - io.BytesIO hoặc None: Đối tượng BytesIO chứa file Excel chi tiết các hóa đơn sai lệch,
                                hoặc None nếu không có sai lệch.
        - None: Luôn là None cho date_ambiguity_info (logic này đã được loại bỏ).
    """
    try:
        # --- Đọc file Excel vào DataFrame ---
        # Bảng kê Thuế: header ở hàng 6 (index 5)
        df_tax = pd.read_excel(tax_invoice_file_stream, header=5, engine='openpyxl')
        print(f"DEBUG: Columns read from Bảng kê Thuế (before renaming): {df_tax.columns.tolist()}")
        print(f"DEBUG: Shape of df_tax (after read, before rename): {df_tax.shape}")
        
        # Bảng kê HĐĐT: header ở hàng 9 (index 8)
        df_e_invoice = pd.read_excel(e_invoice_file_stream, header=8, engine='openpyxl')
        print(f"DEBUG: Columns read from Bảng kê HĐĐT (before renaming): {df_e_invoice.columns.tolist()}")
        print(f"DEBUG: Shape of df_e_invoice (after read, before rename): {df_e_invoice.shape}")

        # Loại bỏ hàng chứa số thứ tự cột ([1], [2], ...) trong Bảng kê HĐĐT nếu có
        if not df_e_invoice.empty:
            # Kiểm tra xem hàng đầu tiên của dữ liệu có chứa định dạng '[số]' không
            is_col_number_row = df_e_invoice.iloc[0].astype(str).str.fullmatch(r'\[\d+\]').any()
            
            if is_col_number_row:
                print("DEBUG: Detected column number row in Bảng kê HĐĐT (row 10), removing it.")
                df_e_invoice = df_e_invoice.iloc[1:].reset_index(drop=True)
                print(f"DEBUG: Shape of df_e_invoice (after removing column number row): {df_e_invoice.shape}")
            else:
                print("DEBUG: Column number row (row 10) not detected or not present in Bảng kê HĐĐT's first data row.")

        # --- Chuẩn hóa tên cột và tạo Chứng minh thư ---
        # Ánh xạ cột cho Bảng kê Thuế (dựa trên vị trí cột bạn cung cấp)
        tax_col_mapping = {
            'Ký hiệu mẫu số': 'invoice_template_symbol_tax',    # Cột B
            'Ký hiệu hóa đơn': 'invoice_series_symbol_tax',     # Cột C
            'Số hóa đơn': 'invoice_number_tax',                 # Cột D
            'Tên người mua/Tên người nhận hàng': 'buyer_name_tax', # Cột G
            'MST người mua/MST người nhận hàng': 'buyer_tax_id_tax', # Cột H
            'Tổng tiền chưa thuế': 'sub_total_amount_tax',      # Cột K
            'Tổng tiền thuế': 'tax_amount_tax',                  # Cột L
            'Tổng tiền thanh toán': 'total_amount_tax',          # Cột O
            # Giữ lại các cột khác nếu cần cho báo cáo chi tiết
            'STT': 'stt_tax',
            'Ngày lập': 'invoice_date_tax',
            'MST người bán/MST người xuất hàng': 'seller_tax_id_tax',
            'Tên người bán/Tên người xuất hàng': 'seller_name_tax',
            'Địa chỉ người mua': 'buyer_address_tax',
            'Tổng tiền chiết khấu thương mại': 'discount_amount_tax',
            'Tổng tiền phí': 'fee_amount_tax',
            'Đơn vị tiền tệ': 'currency_tax',
            'Tỷ giá': 'exchange_rate_tax',
            'Trạng thái hóa đơn': 'invoice_status_tax',
            'Kết quả kiểm tra hóa đơn': 'invoice_check_result_tax',
        }
        df_tax = df_tax.rename(columns=tax_col_mapping)
        print(f"DEBUG: Columns of Bảng kê Thuế (after initial renaming): {df_tax.columns.tolist()}")

        # Ánh xạ cột cho Bảng kê HĐĐT (dựa trên vị trí cột bạn cung cấp)
        e_invoice_col_mapping = {
            'Mẫu số': 'invoice_template_symbol_e_inv',          # Cột R
            'Ký hiệu': 'invoice_series_symbol_e_inv',           # Cột S
            'Số hóa đơn': 'invoice_number_e_inv',               # Cột T
            'Tên khách hàng': 'customer_name_e_inv',            # Cột D
            'MST khách hàng': 'buyer_tax_id_e_inv',             # Cột F
            'Thành tiền': 'sub_total_amount_e_inv',             # Cột N
            'Tiền thuế': 'tax_amount_e_inv',                     # Cột P
            'Phải thu': 'total_amount_e_inv',                    # Cột Q
            'Fkey': 'fkey_e_inv',                                # Cột Y
            # Giữ lại các cột khác nếu cần cho báo cáo chi tiết
            'STT': 'stt_e_inv',
            'Số công văn (Số tham chiếu)': 'reference_number_e_inv',
            'Mã KH (FAST)': 'customer_code_fast_e_inv',
            'Địa chỉ khách hàng': 'customer_address_e_inv',
            'Mặt hàng': 'item_description_e_inv',
            'Kho xuất hàng': 'warehouse_e_inv',
            'Số lượng': 'quantity_e_inv',
            'Đơn giá': 'unit_price_e_inv',
            'Đơn vị tính': 'unit_e_inv',
            'Đơn vị chuyển đổi': 'conversion_unit_e_inv',
            'Số lượng chuyển đổi': 'converted_quantity_e_inv',
            'VAT': 'vat_rate_e_inv',
            'Ngày hóa đơn': 'invoice_date_e_inv',
            'Trạng thái': 'status_e_inv', # IMPORTANT: Used for filtering "Đã phát hành"
            'Ghi chú (Hóa đơn)': 'invoice_notes_e_inv',
            'Ghi chú (Hàng hóa)': 'item_notes_e_inv',
            'PT Thanh toán': 'payment_method_e_inv',
            'Nguồn': 'source_e_inv',
            'Ngày ký hóa đơn': 'invoice_sign_date_e_inv',
            'Ngày giờ lập hóa đơn': 'invoice_creation_datetime_e_inv',
            'Trạng thái thuế': 'tax_status_e_inv',
            'Lý do thuế ': 'tax_reason_e_inv',
        }
        df_e_invoice = df_e_invoice.rename(columns=e_invoice_col_mapping)
        print(f"DEBUG: Columns of Bảng kê HĐĐT (after initial renaming): {df_e_invoice.columns.tolist()}")

        # --- Lọc df_e_invoice theo 'Trạng thái' == 'Đã phát hành' ---
        if 'status_e_inv' in df_e_invoice.columns:
            initial_e_invoice_count = len(df_e_invoice)
            df_e_invoice_published = df_e_invoice[
                df_e_invoice['status_e_inv'].astype(str).str.strip().str.lower() == 'đã phát hành'
            ].copy()
            print(f"DEBUG: Filtered Bảng kê HĐĐT to 'Đã phát hành' invoices. Original: {initial_e_invoice_count}, Filtered: {len(df_e_invoice_published)}")
        else:
            print("WARNING: 'Trạng thái' (status_e_inv) column not found in Bảng kê HĐĐT. All e-invoices will be considered for reconciliation.")
            df_e_invoice_published = df_e_invoice.copy()

        # --- Tạo cột 'Chứng minh thư' cho cả hai DataFrame ---
        # Đảm bảo các cột cần thiết để tạo 'Chứng minh thư' tồn tại
        required_tax_cols_for_id = ['invoice_template_symbol_tax', 'invoice_series_symbol_tax', 'invoice_number_tax']
        required_e_inv_cols_for_id = ['invoice_template_symbol_e_inv', 'invoice_series_symbol_e_inv', 'invoice_number_e_inv']

        for col in required_tax_cols_for_id:
            if col not in df_tax.columns:
                raise ValueError(f"Missing required column '{col}' in Bảng kê Thuế for creating 'Chứng minh thư'.")
        for col in required_e_inv_cols_for_id:
            if col not in df_e_invoice_published.columns:
                raise ValueError(f"Missing required column '{col}' in Bảng kê HĐĐT (Đã phát hành) for creating 'Chứng minh thư'.")

        df_tax['invoice_identity'] = df_tax.apply(
            lambda row: _create_invoice_identity(row, 'invoice_template_symbol_tax', 'invoice_series_symbol_tax', 'invoice_number_tax'),
            axis=1
        )
        df_e_invoice_published['invoice_identity'] = df_e_invoice_published.apply(
            lambda row: _create_invoice_identity(row, 'invoice_template_symbol_e_inv', 'invoice_series_symbol_e_inv', 'invoice_number_e_inv'),
            axis=1
        )
        
        # Loại bỏ các hàng có 'Chứng minh thư' rỗng/None
        df_tax.dropna(subset=['invoice_identity'], inplace=True)
        df_e_invoice_published.dropna(subset=['invoice_identity'], inplace=True)

        print(f"DEBUG: Sample of 'Chứng minh thư' from Bảng kê Thuế:")
        print(df_tax[['invoice_template_symbol_tax', 'invoice_series_symbol_tax', 'invoice_number_tax', 'invoice_identity']].head())
        print(f"DEBUG: Sample of 'Chứng minh thư' from Bảng kê HĐĐT (Đã phát hành):")
        print(df_e_invoice_published[['invoice_template_symbol_e_inv', 'invoice_series_symbol_e_inv', 'invoice_number_e_inv', 'invoice_identity']].head())
        
        print(f"DEBUG: Shape of df_tax (after creating identity & dropna): {df_tax.shape}")
        print(f"DEBUG: Shape of df_e_invoice_published (after creating identity & dropna): {df_e_invoice_published.shape}")


        # --- Hợp nhất (Merge) DataFrames dựa trên 'Chứng minh thư' ---
        # Sử dụng outer merge để tìm tất cả các hóa đơn từ cả hai nguồn
        merged_df = pd.merge(
            df_e_invoice_published, # Left: HĐĐT (đã phát hành)
            df_tax,                 # Right: Thuế
            on='invoice_identity',
            how='outer',            # Giữ tất cả các hàng từ cả hai DataFrame
            suffixes=('_e_inv', '_tax'),
            indicator=True          # Thêm cột '_merge' để chỉ ra nguồn gốc
        )
        print(f"DEBUG: Columns of merged_df: {merged_df.columns.tolist()}")
        print(f"DEBUG: Shape of merged_df (after outer merge): {merged_df.shape}")

        # --- Khởi tạo cột Lý do sai lệch chi tiết ---
        merged_df['Lý do sai lệch chi tiết'] = ''

        # --- Phát hiện sai lệch tồn tại (Existence Mismatches) ---
        # Hóa đơn có trong HĐĐT (đã phát hành) nhưng không có trong Thuế
        left_only_mask = merged_df['_merge'] == 'left_only'
        merged_df.loc[left_only_mask, 'Lý do sai lệch chi tiết'] = merged_df.loc[left_only_mask, 'Lý do sai lệch chi tiết'].apply(
            lambda x: (x + '; Chưa được đẩy lên Cơ quan Thuế') if x else 'Chưa được đẩy lên Cơ quan Thuế'
        )

        # Hóa đơn có trong Thuế nhưng không có trong HĐĐT (đã phát hành)
        right_only_mask = merged_df['_merge'] == 'right_only'
        merged_df.loc[right_only_mask, 'Lý do sai lệch chi tiết'] = merged_df.loc[right_only_mask, 'Lý do sai lệch chi tiết'].apply(
            lambda x: (x + '; Không có trong Bảng kê HĐĐT đã phát hành') if x else 'Không có trong Bảng kê HĐĐT đã phát hành'
        )

        # --- Phát hiện sai lệch giá trị (Value Discrepancies) cho các hóa đơn khớp ID ---
        both_present_df_indices = merged_df[merged_df['_merge'] == 'both'].index

        # Định nghĩa các trường cần so sánh giá trị
        # (cột HĐĐT, cột Thuế, tên hiển thị)
        comparison_fields = [
            ('customer_name_e_inv', 'buyer_name_tax', 'Tên khách hàng'),
            ('buyer_tax_id_e_inv', 'buyer_tax_id_tax', 'MST khách hàng'), 
            ('sub_total_amount_e_inv', 'sub_total_amount_tax', 'Tổng tiền chưa thuế'),
            ('tax_amount_e_inv', 'tax_amount_tax', 'Tiền thuế'),
            ('total_amount_e_inv', 'total_amount_tax', 'Tổng tiền thanh toán'),
        ]

        print("\nDEBUG: Sample of comparison values for 'both' merged invoices:")
        if not both_present_df_indices.empty:
            sample_rows = merged_df.loc[both_present_df_indices].head(5)
            for e_inv_col, tax_col, display_name in comparison_fields:
                if tax_col in sample_rows.columns and e_inv_col in sample_rows.columns:
                    print(f"  --- {display_name} ---")
                    print(f"    Tax (Original): {sample_rows[tax_col].tolist()}")
                    print(f"    E-Inv (Original): {sample_rows[e_inv_col].tolist()}")
                    
                    if 'Tên khách hàng' in display_name or 'MST khách hàng' in display_name:
                        val_tax_debug = sample_rows[tax_col].astype(str).str.strip().str.lower().tolist()
                        val_e_inv_debug = sample_rows[e_inv_col].astype(str).str.strip().str.lower().tolist()
                        print(f"    Tax (Cleaned String): {val_tax_debug}")
                        print(f"    E-Inv (Cleaned String): {val_e_inv_debug}")
                    else: # Numerical
                        val_tax_cleaned_str_debug = sample_rows[tax_col].apply(_clean_numeric_string_for_int)
                        val_e_inv_cleaned_str_debug = sample_rows[e_inv_col].apply(_clean_numeric_string_for_int)
                        print(f"    Tax (Cleaned Numeric Str): {val_tax_cleaned_str_debug.tolist()}")
                        print(f"    E-Inv (Cleaned Numeric Str): {val_e_inv_cleaned_str_debug.tolist()}")

                        # Chuyển đổi sang nullable integer (Int64) để so sánh chính xác
                        val_tax_int_debug = pd.Series(pd.to_numeric(val_tax_cleaned_str_debug, errors='coerce')).fillna(pd.NA).astype('Int64')
                        val_e_inv_int_debug = pd.Series(pd.to_numeric(val_e_inv_cleaned_str_debug, errors='coerce')).fillna(pd.NA).astype('Int64')
                        print(f"    Tax (Numeric Int64): {val_tax_int_debug.tolist()}")
                        print(f"    E-Inv (Numeric Int64): {val_e_inv_int_debug.tolist()}")
                        
                        diffs = []
                        for t, e in zip(val_tax_int_debug, val_e_inv_int_debug):
                            if pd.notna(t) and pd.notna(e):
                                diffs.append(abs(t - e))
                            else:
                                diffs.append(None) # Indicate no valid comparison
                        print(f"    Differences: {diffs}")


        for e_inv_col, tax_col, display_name in comparison_fields:
            # Chỉ so sánh cho các hàng có '_merge' là 'both'
            if tax_col in merged_df.columns and e_inv_col in merged_df.columns:
                # Lấy các giá trị từ các cột tương ứng chỉ cho các hàng 'both'
                tax_values = merged_df.loc[both_present_df_indices, tax_col]
                e_inv_values = merged_df.loc[both_present_df_indices, e_inv_col]

                mismatch_mask_for_field = pd.Series(False, index=both_present_df_indices)

                if 'Tên khách hàng' in display_name or 'MST khách hàng' in display_name:
                    # So sánh chuỗi: làm sạch và chuyển về chữ thường
                    val_tax_cleaned = tax_values.astype(str).str.strip().str.lower()
                    val_e_inv_cleaned = e_inv_values.astype(str).str.strip().str.lower()
                    
                    # Sai lệch nếu: (một bên là NaN và bên kia không) HOẶC (cả hai đều không NaN VÀ giá trị khác nhau)
                    mismatch_mask_for_field = (val_tax_cleaned.isna() != val_e_inv_cleaned.isna()) | \
                                              ((~val_tax_cleaned.isna()) & (~val_e_inv_cleaned.isna()) & (val_tax_cleaned != val_e_inv_cleaned))
                else: # So sánh số nguyên (Tổng tiền chưa thuế, Tiền thuế, Tổng tiền thanh toán)
                    # Áp dụng hàm làm sạch số nguyên
                    val_tax_cleaned_str = tax_values.apply(_clean_numeric_string_for_int)
                    val_e_inv_cleaned_str = e_inv_values.apply(_clean_numeric_string_for_int)

                    # Chuyển đổi sang nullable integer (Int64)
                    val_tax_int = pd.Series(pd.to_numeric(val_tax_cleaned_str, errors='coerce')).fillna(pd.NA).astype('Int64')
                    val_e_inv_int = pd.Series(pd.to_numeric(val_e_inv_cleaned_str, errors='coerce')).fillna(pd.NA).astype('Int64')
                    
                    # Sai lệch nếu: (một bên là NaN/NA và bên kia không) HOẶC (cả hai đều không NaN/NA VÀ giá trị khác nhau)
                    mismatch_mask_for_field = (val_tax_int.isna() != val_e_inv_int.isna()) | \
                                              ((~val_tax_int.isna()) & (~val_e_inv_int.isna()) & (val_tax_int != val_e_inv_int))
                
                # Cập nhật cột 'Lý do sai lệch chi tiết' cho các hàng có sai lệch
                current_reasons = merged_df.loc[both_present_df_indices[mismatch_mask_for_field], 'Lý do sai lệch chi tiết']
                merged_df.loc[both_present_df_indices[mismatch_mask_for_field], 'Lý do sai lệch chi tiết'] = current_reasons.apply(
                    lambda x: (x + f'; Sai lệch {display_name}') if x else f'Sai lệch {display_name}'
                )

        # --- Lọc các hóa đơn sai lệch để tạo báo cáo ---
        # Hóa đơn sai lệch là những hóa đơn có lý do sai lệch chi tiết (không rỗng)
        mismatched_details_df = merged_df[
            merged_df['Lý do sai lệch chi tiết'].astype(bool) # Lọc các hàng có lý do sai lệch
        ].copy()
        
        # Làm sạch cột 'Lý do sai lệch chi tiết'
        mismatched_details_df['Lý do sai lệch chi tiết'] = mismatched_details_df['Lý do sai lệch chi tiết'].str.strip(' ;')
        mismatched_details_df.loc[mismatched_details_df['Lý do sai lệch chi tiết'] == '', 'Lý do sai lệch chi tiết'] = 'Không xác định lý do chi tiết'

        print(f"DEBUG: Columns of mismatched_details_df: {mismatched_details_df.columns.tolist()}")
        print(f"DEBUG: Shape of mismatched_details_df: {mismatched_details_df.shape}")


        # --- Chuẩn bị tóm tắt cho frontend ---
        all_mismatched_invoice_numbers_for_summary = []
        if not mismatched_details_df.empty:
            # Sử dụng invoice_identity để nhóm các lý do sai lệch
            unique_mismatched_identities = mismatched_details_df['invoice_identity'].dropna().unique().tolist()
            unique_mismatched_identities.sort()

            for inv_id in unique_mismatched_identities:
                reasons = mismatched_details_df[mismatched_details_df['invoice_identity'] == inv_id]['Lý do sai lệch chi tiết'].unique()
                # Lấy số hóa đơn gốc từ HĐĐT để hiển thị trong tóm tắt
                original_e_inv_number = mismatched_details_df[mismatched_details_df['invoice_identity'] == inv_id]['invoice_number_e_inv'].iloc[0]
                all_mismatched_invoice_numbers_for_summary.append(f"Số HĐ: {original_e_inv_number} (CMT: {inv_id}) - Lý do: {', '.join(reasons)}")

        # Tính số lượng hóa đơn khớp: những hóa đơn có '_merge' là 'both' VÀ không có lý do sai lệch chi tiết
        matched_count = len(merged_df[
            (merged_df['_merge'] == 'both') & 
            (merged_df['Lý do sai lệch chi tiết'] == '') # Không có lý do sai lệch sau tất cả các kiểm tra
        ])
        print(f"DEBUG: Matched count: {matched_count}")

        comparison_summary = {
            'matched_count': matched_count,
            'mismatched_invoices': all_mismatched_invoice_numbers_for_summary
        }

        # --- Tạo file Excel kết quả chi tiết ---
        output_excel_stream = None
        if not mismatched_details_df.empty: # Chỉ tạo Excel nếu có sai lệch
            output_excel_stream = io.BytesIO()
            
            # Định nghĩa ánh xạ tên cột hiển thị đầy đủ cho file Excel đầu ra
            # Cấu trúc: 'tên cột nội bộ': 'tên hiển thị trong Excel'
            full_display_col_mapping = {
                'invoice_identity': 'Số/ký hiệu hóa đơn bị sai lệch',
                'customer_name_e_inv': 'Tên khách hàng (HĐĐT)',
                'buyer_name_tax': 'Tên khách hàng (Thuế)',
                'buyer_tax_id_e_inv': 'Mã số thuế (HĐĐT)',
                'buyer_tax_id_tax': 'Mã số thuế (Thuế)',
                'sub_total_amount_e_inv': 'Tổng tiền chưa thuế (HĐĐT)',
                'sub_total_amount_tax': 'Tổng tiền chưa thuế (Thuế)',
                'tax_amount_e_inv': 'Tiền thuế (HĐĐT)',
                'tax_amount_tax': 'Tiền thuế (Thuế)',
                'total_amount_e_inv': 'Tổng tiền thanh toán (HĐĐT)',
                'total_amount_tax': 'Tổng tiền thanh toán (Thuế)',
                'fkey_e_inv': 'Mã FKEY', # Thêm cột FKEY
                'Lý do sai lệch chi tiết': 'Lý do sai lệch',
                # Có thể thêm các cột khác từ df_e_invoice_published hoặc df_tax nếu cần hiển thị thêm thông tin gốc
                # Ví dụ: 'invoice_date_e_inv': 'Ngày hóa đơn (HĐĐT)',
                #         'invoice_date_tax': 'Ngày lập (Thuế)',
            }

            # Xây dựng danh sách các cột theo thứ tự mong muốn cho file Excel đầu ra
            # Đảm bảo thứ tự và các cột con như trong ảnh bạn gửi
            final_output_cols_order = [
                'invoice_identity',
                'customer_name_e_inv',
                'buyer_name_tax',
                'buyer_tax_id_e_inv',
                'buyer_tax_id_tax',
                'sub_total_amount_e_inv',
                'sub_total_amount_tax',
                'tax_amount_e_inv',
                'tax_amount_tax',
                'total_amount_e_inv',
                'total_amount_tax',
                'fkey_e_inv', # Thêm cột FKEY vào đây
                'Lý do sai lệch chi tiết',
            ]
            
            # Lọc DataFrame chỉ lấy các cột có trong danh sách và tồn tại trong df
            # Đảm bảo chỉ lấy các cột đã được định nghĩa trong full_display_col_mapping
            final_output_df = mismatched_details_df[[col for col in final_output_cols_order if col in mismatched_details_df.columns]].copy()
            
            # Đổi tên các cột để hiển thị đẹp hơn trong Excel
            # Chỉ đổi tên những cột có trong full_display_col_mapping
            final_output_df = final_output_df.rename(columns={k: v for k, v in full_display_col_mapping.items() if k in final_output_df.columns})

            print(f"DEBUG: Columns of final_output_df before saving: {final_output_df.columns.tolist()}")
            print("DEBUG: Head of final_output_df before saving:")
            print(final_output_df.head())

            final_output_df.to_excel(output_excel_stream, index=False, engine='openpyxl')
            output_excel_stream.seek(0)

        return comparison_summary, output_excel_stream, None

    except ValueError as e:
        print(f"Lỗi dữ liệu đầu vào hoặc cấu hình: {e}")
        raise ValueError(f"Lỗi dữ liệu đầu vào hoặc cấu hình: {e}")
    except Exception as e:
        print(f"Lỗi trong quá trình đối soát: {e}")
        raise ValueError(f"Lỗi trong quá trình đối soát: {e}")
