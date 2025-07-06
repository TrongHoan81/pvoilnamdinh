import pandas as pd
from openpyxl import load_workbook, Workbook
from datetime import datetime
import re
import io

# ==============================================================================
# HÀM NHẬN DIỆN LOẠI BẢNG KÊ (Đã ổn định)
# ==============================================================================
def detect_report_type(file_content_bytes):
    """
    Hàm nhận diện đơn giản và đáng tin cậy nhất, dựa trên gợi ý của người dùng.
    """
    try:
        wb = load_workbook(io.BytesIO(file_content_bytes), data_only=True)
        ws = wb.active
        # Kiểm tra cho file POS bằng ô B4
        if ws['B4'].value and 'seri' == str(ws['B4'].value).lower().strip():
            return 'POS'
        # Kiểm tra cho file HDDT bằng dòng 9
        for cell in ws[9]:
            if cell.value and 'số công văn (số tham chiếu)' in str(cell.value).lower():
                return 'HDDT'
    except Exception:
        return 'UNKNOWN'
    return 'UNKNOWN'

# ==============================================================================
# KHỐI 1: TOÀN BỘ LOGIC GỐC CỦA ỨNG DỤNG POS
# (SAO CHÉP 100% TỪ FILE GỐC CỦA BẠN - KHÔNG THAY ĐỔI)
# ==============================================================================
def process_pos_report(file_content_bytes, selected_chxd, price_periods, new_price_invoice_number, **kwargs):
    
    # --- Các hàm nội bộ của logic POS (từ file gốc) ---
    def _to_float_pos(value):
        try:
            if isinstance(value, str): value = value.replace(",", "").strip()
            return float(value)
        except (ValueError, TypeError): return 0.0

    def _clean_string_pos(s):
        if s is None: return ""
        return re.sub(r'\s+', ' ', str(s)).strip()

    def _format_tax_code_pos(vat_str):
        try:
            num = float(vat_str)
            if 0 < num < 1:
                return f"{int(num * 100):02d}"
            return f"{int(num):02d}"
        except (ValueError, TypeError):
            return "08"

    def _get_static_data_pos(file_path):
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active
            chxd_detail_map, store_specific_x_lookup = {}, {}
            for row_idx in range(4, ws.max_row + 1):
                row_values = [cell.value for cell in ws[row_idx]]
                if len(row_values) < 24: continue
                chxd_name = _clean_string_pos(row_values[3])
                if not chxd_name: continue
                chxd_detail_map[chxd_name] = {
                    'g4_val': _clean_string_pos(row_values[6]), 'g5_val': _clean_string_pos(row_values[4]),
                    't_val': _clean_string_pos(row_values[19]), 'u_val': _clean_string_pos(row_values[20]),
                    'v_val': _clean_string_pos(row_values[21]), 'w_val': _clean_string_pos(row_values[22]),
                    'x_val': _clean_string_pos(row_values[23])
                }
                store_specific_x_lookup[chxd_name] = {
                    _clean_string_pos(ws['O3'].value).lower(): _clean_string_pos(row_values[14]),
                    _clean_string_pos(ws['P3'].value).lower(): _clean_string_pos(row_values[15]),
                    _clean_string_pos(ws['Q3'].value).lower(): _clean_string_pos(row_values[16]),
                    _clean_string_pos(ws['R3'].value).lower(): _clean_string_pos(row_values[17])
                }
            tmt_map = {_clean_string_pos(ws[f'A{i}'].value): _to_float_pos(ws[f'B{i}'].value) for i in range(4, 8)}
            return {"chxd_details": chxd_detail_map, "x_lookup": store_specific_x_lookup, "tmt_map": tmt_map}, None
        except FileNotFoundError: return None, "Lỗi: Không tìm thấy file Data_POS.xlsx."
        except Exception as e: return None, f"Lỗi khi đọc file Data_POS.xlsx: {e}"

    def _create_upsse_workbook_pos():
        headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
        wb = Workbook()
        ws = wb.active
        for _ in range(4): ws.append([''] * len(headers))
        ws.append(headers)
        return wb

    def _create_pos_invoice_row(row, final_date, details, static_data):
        new_row = [''] * 37
        new_row[0] = details['g4_val']
        new_row[1] = new_row[31] = _clean_string_pos(row['Tên khách hàng'])
        new_row[2] = final_date
        new_row[3] = f"{_clean_string_pos(row['Séri'])}{int(_to_float_pos(row['Số']))}"
        new_row[4] = _clean_string_pos(row['Séri'])
        new_row[5] = f"Xuất bán hàng theo hóa đơn số {new_row[3]}"
        new_row[7] = _clean_string_pos(row['Hàng hóa'])
        new_row[8] = "Lít"
        new_row[9] = details['g5_val']
        new_row[12] = _to_float_pos(row['Số lượng'])
        tmt_value = static_data['tmt_map'].get(new_row[7], 0)
        new_row[13] = _to_float_pos(row['Đơn giá(đã có thuế GTGT)']) - tmt_value
        new_row[14] = _to_float_pos(row['Tổng tiền thanh toán']) - _to_float_pos(row['Tiền thuế GTGT']) - round(tmt_value * new_row[12], 0)
        new_row[17] = _format_tax_code_pos(str(row['Thuế GTGT']))
        new_row[18], new_row[19], new_row[20], new_row[21] = details['t_val'], details['u_val'], details['v_val'], details['w_val']
        new_row[33] = _clean_string_pos(row['Mã số thuế'])
        new_row[36] = _to_float_pos(row['Tiền thuế GTGT']) - round(tmt_value * new_row[12] * (_to_float_pos(new_row[17]) / 100.0), 0)
        return new_row

    def _create_pos_summary_row(product_name, data, final_date, details, static_data, is_new_price, selected_chxd):
        new_row = [''] * 37
        new_row[0] = details['g4_val']
        new_row[2] = final_date
        suffix = static_data['x_lookup'][selected_chxd].get(product_name.lower(), '')
        invoice_suffix = f"M{suffix}" if is_new_price else suffix
        new_row[3] = f"BK{pd.to_datetime(final_date, dayfirst=True).strftime('%d%m')}{invoice_suffix}"
        new_row[5] = f"Xuất bán hàng theo bảng kê số {new_row[3]}"
        new_row[7] = product_name
        new_row[8] = "Lít"
        new_row[9] = details['g5_val']
        new_row[12] = data['sl']
        tmt_value = static_data['tmt_map'].get(product_name, 0)
        new_row[13] = data['dongia'] - tmt_value
        new_row[14] = data['tt'] - data['tienthuegtgt'] - round(tmt_value * data['sl'], 0)
        new_row[17] = _format_tax_code_pos(str(data['thue']))
        new_row[18], new_row[19], new_row[20], new_row[21] = details['t_val'], details['u_val'], details['v_val'], details['w_val']
        new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"
        new_row[36] = data['tienthuegtgt'] - round(tmt_value * data['sl'] * (_to_float_pos(new_row[17]) / 100.0), 0)
        return new_row

    def _create_pos_tmt_row(original_row, tmt_value, details):
        tmt_row = list(original_row)
        tax_rate_decimal = _to_float_pos(original_row[17]) / 100.0
        tmt_row[6], tmt_row[7], tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "Lít"
        tmt_row[9] = details['g5_val']
        tmt_row[13] = tmt_value
        tmt_row[14] = round(tmt_value * _to_float_pos(original_row[12]), 0)
        tmt_row[18], tmt_row[19], tmt_row[20], tmt_row[21] = "1311", "51133", "6328", "33313"
        tmt_row[36] = round(tmt_row[14] * tax_rate_decimal)
        tmt_row[5], tmt_row[31], tmt_row[32], tmt_row[33] = '', '', '', ''
        return tmt_row

    # --- Bắt đầu logic chính của POS (từ file gốc) ---
    static_data, error = _get_static_data_pos("Data_POS.xlsx")
    if error: raise ValueError(error)
    df = pd.read_excel(io.BytesIO(file_content_bytes), header=4)
    df.columns = [_clean_string_pos(col) for col in df.columns]
    df = df.dropna(how='all')
    df = df[df['Số lượng'].apply(_to_float_pos) > 0]
    if df.empty: raise ValueError("Không có dữ liệu hợp lệ (Số lượng > 0) trong file POS.")
    final_date = pd.to_datetime(df['Ngày'].iloc[0]).strftime('%d/%m/%Y')
    chxd_details = static_data['chxd_details'].get(selected_chxd)
    if not chxd_details: raise ValueError(f"Không tìm thấy cấu hình cho CHXD: {selected_chxd}")
    def _process_rows(dataframe, is_new_price):
        processed_rows, summary_data = [], {}
        for _, row in dataframe.iterrows():
            product_name = _clean_string_pos(row['Hàng hóa'])
            if "không lấy hóa đơn" in _clean_string_pos(row['Tên khách hàng']):
                if product_name not in summary_data:
                    summary_data[product_name] = {'sl': 0, 'tt': 0, 'thue': 0, 'tienthuegtgt': 0, 'dongia': _to_float_pos(row['Đơn giá(đã có thuế GTGT)'])}
                summary_data[product_name]['sl'] += _to_float_pos(row['Số lượng'])
                summary_data[product_name]['tt'] += _to_float_pos(row['Tổng tiền thanh toán'])
                summary_data[product_name]['thue'] += _to_float_pos(row['Thuế GTGT'])
                summary_data[product_name]['tienthuegtgt'] += _to_float_pos(row['Tiền thuế GTGT'])
            else:
                processed_rows.append(_create_pos_invoice_row(row, final_date, chxd_details, static_data))
        for product, data in summary_data.items():
            processed_rows.append(_create_pos_summary_row(product, data, final_date, chxd_details, static_data, is_new_price, selected_chxd))
        final_df_rows = []
        for row_data in processed_rows:
            final_df_rows.append(row_data)
            tmt_value = static_data['tmt_map'].get(row_data[7], 0)
            if tmt_value > 0: final_df_rows.append(_create_pos_tmt_row(row_data, tmt_value, chxd_details))
        return final_df_rows
    if price_periods == '1':
        all_rows_for_df = _process_rows(df, is_new_price=False)
        final_wb = _create_upsse_workbook_pos()
        for r in all_rows_for_df: final_wb.active.append(r)
        output = io.BytesIO()
        final_wb.save(output)
        output.seek(0)
        return output
    else:
        if not new_price_invoice_number: raise ValueError("Vui lòng nhập số hóa đơn của giá mới.")
        split_idx = df.index[df['Số'].apply(_to_float_pos) == _to_float_pos(new_price_invoice_number)].tolist()
        if not split_idx: raise ValueError(f"Không tìm thấy số hóa đơn '{new_price_invoice_number}'")
        df_old, df_new = df.loc[:split_idx[0]-1], df.loc[split_idx[0]:]
        wb_old, wb_new = _create_upsse_workbook_pos(), _create_upsse_workbook_pos()
        for r in _process_rows(df_old, is_new_price=False): wb_old.active.append(r)
        for r in _process_rows(df_new, is_new_price=True): wb_new.active.append(r)
        output_old, output_new = io.BytesIO(), io.BytesIO()
        wb_old.save(output_old); wb_new.save(output_new)
        output_old.seek(0); output_new.seek(0)
        return {'old': output_old, 'new': output_new}

# ==============================================================================
# KHỐI 2: LOGIC HDDT (ĐÃ ỔN ĐỊNH - KHÔNG CHỈNH SỬA)
# ==============================================================================
def process_hddt_report(file_content_bytes, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=None):
    
    # --- Các hàm nội bộ của logic HDDT (từ file gốc) ---
    def _clean_string_hddt(s):
        if s is None: return ""
        cleaned_s = str(s).strip()
        if cleaned_s.startswith("'"): cleaned_s = cleaned_s[1:]
        return re.sub(r'\s+', ' ', cleaned_s)

    def _to_float_hddt(value):
        if value is None: return 0.0
        try:
            return float(str(value).replace(',', '').strip())
        except (ValueError, TypeError): return 0.0

    def _format_tax_code_hddt(raw_vat_value):
        if raw_vat_value is None: return ""
        try:
            s_value = str(raw_vat_value).replace('%', '').strip()
            f_value = float(s_value)
            if 0 < f_value < 1: f_value *= 100
            return f"{round(f_value):02d}"
        except (ValueError, TypeError): return ""
    
    def _load_static_data_hddt(data_file_path, mahh_file_path, dskh_file_path):
        try:
            static_data = {}
            wb = load_workbook(data_file_path, data_only=True)
            ws = wb.active
            chxd_list, tk_mk_map, khhd_map, chxd_to_khuvuc_map = [], {}, {}, {}
            vu_viec_map = {}
            vu_viec_headers = [_clean_string_hddt(cell.value) for cell in ws[2][4:9]]
            for row_values in ws.iter_rows(min_row=3, max_col=12, values_only=True):
                chxd_name = _clean_string_hddt(row_values[3])
                if chxd_name:
                    ma_kho, khhd, khu_vuc = _clean_string_hddt(row_values[9]), _clean_string_hddt(row_values[10]), _clean_string_hddt(row_values[11])
                    if chxd_name not in tk_mk_map: chxd_list.append(chxd_name)
                    if ma_kho: tk_mk_map[chxd_name] = ma_kho
                    if khhd: khhd_map[chxd_name] = khhd
                    if khu_vuc: chxd_to_khuvuc_map[chxd_name] = khu_vuc
                    vu_viec_map[chxd_name] = {}
                    vu_viec_data_row = row_values[4:9]
                    for i, header in enumerate(vu_viec_headers):
                        if header:
                            key = "Dầu mỡ nhờn" if i == len(vu_viec_headers) - 1 else header
                            vu_viec_map[chxd_name][key] = _clean_string_hddt(vu_viec_data_row[i])
            if not chxd_list: return None, "Không tìm thấy Tên CHXD nào trong cột D của file Data_HDDT.xlsx."
            static_data.update({"DS_CHXD": chxd_list, "tk_mk": tk_mk_map, "khhd_map": khhd_map, "chxd_to_khuvuc_map": chxd_to_khuvuc_map, "vu_viec_map": vu_viec_map})
            def get_lookup_map(min_r, max_r, min_c=1, max_c=2):
                return {_clean_string_hddt(row[0]): row[1] for row in ws.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=max_c, values_only=True) if row[0] and row[1] is not None}
            phi_bvmt_map_raw = get_lookup_map(10, 13)
            static_data["phi_bvmt_map"] = {_clean_string_hddt(k): _to_float_hddt(v) for k, v in phi_bvmt_map_raw.items()}
            static_data.update({
                "tk_no_map": get_lookup_map(29, 31), "tk_doanh_thu_map": get_lookup_map(33, 35),
                "tk_thue_co_map": get_lookup_map(38, 40), "tk_gia_von_value": ws['B36'].value,
                "tk_no_bvmt_map": get_lookup_map(44, 46), "tk_dt_thue_bvmt_map": get_lookup_map(48, 50),
                "tk_gia_von_bvmt_value": ws['B51'].value, "tk_thue_co_bvmt_map": get_lookup_map(53, 55)
            })
            wb_mahh = load_workbook(mahh_file_path, data_only=True)
            static_data["ma_hang_map"] = {_clean_string_hddt(r[0]): _clean_string_hddt(r[2]) for r in wb_mahh.active.iter_rows(min_row=2, max_col=3, values_only=True) if r[0] and r[2]}
            wb_dskh = load_workbook(dskh_file_path, data_only=True)
            static_data["mst_to_makh_map"] = {_clean_string_hddt(r[2]): _clean_string_hddt(r[3]) for r in wb_dskh.active.iter_rows(min_row=2, max_col=4, values_only=True) if r[2]}
            return static_data, None
        except FileNotFoundError as e: return None, f"Lỗi: Không tìm thấy file cấu hình. Chi tiết: {e.filename}"
        except Exception as e: return None, f"Lỗi khi đọc file cấu hình: {e}"

    def _create_hddt_bvmt_row(original_row, phi_bvmt, static_data, khu_vuc):
        bvmt_row = list(original_row)
        so_luong = _to_float_hddt(original_row[12])
        thue_suat = _to_float_hddt(original_row[17]) / 100.0 if original_row[17] else 0.0
        bvmt_row[6], bvmt_row[7] = "TMT", "Thuế bảo vệ môi trường"
        bvmt_row[13], bvmt_row[14] = phi_bvmt, round(phi_bvmt * so_luong)
        bvmt_row[18] = static_data.get('tk_no_bvmt_map', {}).get(khu_vuc)
        bvmt_row[19] = static_data.get('tk_dt_thue_bvmt_map', {}).get(khu_vuc)
        bvmt_row[20] = static_data.get('tk_gia_von_bvmt_value')
        bvmt_row[21] = static_data.get('tk_thue_co_bvmt_map', {}).get(khu_vuc)
        bvmt_row[36] = round(phi_bvmt * so_luong * thue_suat)
        for i in [5, 31, 32, 33]: bvmt_row[i] = ''
        return bvmt_row

    def _generate_upsse_from_hddt_rows(rows_to_process, static_data, selected_chxd, final_date, summary_suffix_map):
        if not rows_to_process: return None
        khu_vuc, ma_kho = static_data['chxd_to_khuvuc_map'].get(selected_chxd), static_data['tk_mk'].get(selected_chxd)
        tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co = static_data['tk_no_map'].get(khu_vuc), static_data['tk_doanh_thu_map'].get(khu_vuc), static_data['tk_gia_von_value'], static_data['tk_thue_co_map'].get(khu_vuc)
        original_invoice_rows, bvmt_rows, summary_data = [], [], {}
        first_invoice_prefix_source = ""
        for bkhd_row in rows_to_process:
            if _to_float_hddt(bkhd_row[8] if len(bkhd_row) > 8 else None) <= 0: continue
            ten_kh, ten_mat_hang = _clean_string_hddt(bkhd_row[3]), _clean_string_hddt(bkhd_row[6])
            is_anonymous, is_petrol = ("không lấy hóa đơn" in ten_kh.lower()), (ten_mat_hang in static_data['phi_bvmt_map'])
            if not is_anonymous or not is_petrol:
                new_upsse_row = [''] * 37
                new_upsse_row[9], new_upsse_row[1], new_upsse_row[31], new_upsse_row[2] = ma_kho, ten_kh, ten_kh, final_date
                so_hd_goc = str(bkhd_row[19] or '').strip()
                new_upsse_row[3] = f"HN{so_hd_goc[-6:]}" if selected_chxd == "Nguyễn Huệ" else f"{(str(bkhd_row[18] or '').strip())[-2:]}{so_hd_goc[-6:]}"
                new_upsse_row[4] = _clean_string_hddt(bkhd_row[17]) + _clean_string_hddt(bkhd_row[18])
                new_upsse_row[5], new_upsse_row[7], new_upsse_row[6] = f"Xuất bán hàng theo hóa đơn số {new_upsse_row[3]}", ten_mat_hang, static_data['ma_hang_map'].get(ten_mat_hang, '')
                new_upsse_row[8], new_upsse_row[12] = _clean_string_hddt(bkhd_row[10]), round(_to_float_hddt(bkhd_row[8]), 3)
                phi_bvmt = static_data['phi_bvmt_map'].get(ten_mat_hang, 0.0) if is_petrol else 0.0
                new_upsse_row[13] = _to_float_hddt(bkhd_row[9]) - phi_bvmt
                ma_thue = _format_tax_code_hddt(bkhd_row[14])
                new_upsse_row[17] = ma_thue
                thue_suat = _to_float_hddt(ma_thue) / 100.0 if ma_thue else 0.0
                tien_thue_goc, so_luong = _to_float_hddt(bkhd_row[15]), _to_float_hddt(bkhd_row[8])
                tien_thue_phi_bvmt = round(phi_bvmt * so_luong * thue_suat)
                new_upsse_row[36] = round(tien_thue_goc - tien_thue_phi_bvmt)
                new_upsse_row[14] = round(_to_float_hddt(bkhd_row[13]) if not is_petrol else _to_float_hddt(bkhd_row[16]) - tien_thue_goc - round(phi_bvmt * so_luong))
                new_upsse_row[18], new_upsse_row[19], new_upsse_row[20], new_upsse_row[21] = tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co
                chxd_vu_viec_map = static_data['vu_viec_map'].get(selected_chxd, {})
                new_upsse_row[23] = chxd_vu_viec_map.get(ten_mat_hang, chxd_vu_viec_map.get("Dầu mỡ nhờn", ''))
                new_upsse_row[32], mst_khach_hang = _clean_string_hddt(bkhd_row[4]), _clean_string_hddt(bkhd_row[5])
                new_upsse_row[33] = mst_khach_hang
                ma_kh_fast = _clean_string_hddt(bkhd_row[2])
                new_upsse_row[0] = ma_kh_fast if ma_kh_fast and len(ma_kh_fast) < 12 else static_data['mst_to_makh_map'].get(mst_khach_hang, ma_kho)
                original_invoice_rows.append(new_upsse_row)
                if is_petrol: bvmt_rows.append(_create_hddt_bvmt_row(new_upsse_row, phi_bvmt, static_data, khu_vuc))
            else:
                if not first_invoice_prefix_source: first_invoice_prefix_source = str(bkhd_row[18] or '').strip()
                if ten_mat_hang not in summary_data:
                    summary_data[ten_mat_hang] = {'sl': 0, 'thue': 0, 'phai_thu': 0, 'first_data': {'mau_so': _clean_string_hddt(bkhd_row[17]),'ky_hieu': _clean_string_hddt(bkhd_row[18]),'don_gia': _to_float_hddt(bkhd_row[9]),'vat_raw': bkhd_row[14]}}
                summary_data[ten_mat_hang]['sl'] += _to_float_hddt(bkhd_row[8])
                summary_data[ten_mat_hang]['thue'] += _to_float_hddt(bkhd_row[15])
                summary_data[ten_mat_hang]['phai_thu'] += _to_float_hddt(bkhd_row[16])
        prefix = first_invoice_prefix_source[-2:] if len(first_invoice_prefix_source) >= 2 else first_invoice_prefix_source
        for product, data in summary_data.items():
            summary_row = [''] * 37
            first_data, total_sl = data['first_data'], data['sl']
            phi_bvmt = static_data['phi_bvmt_map'].get(product, 0.0)
            ma_thue = _format_tax_code_hddt(first_data['vat_raw'])
            thue_suat = _to_float_hddt(ma_thue) / 100.0 if ma_thue else 0.0
            TDT, TTT = data['phai_thu'], data['thue']
            TH_TMT, TT_TMT = round(phi_bvmt * total_sl), round(phi_bvmt * total_sl * thue_suat)
            TT_goc, TH_goc = TTT - TT_TMT, TDT - TH_TMT - (TTT - TT_TMT) - TT_TMT
            summary_row[0], summary_row[1] = ma_kho, f"Khách hàng mua {product} không lấy hóa đơn"
            summary_row[31], summary_row[2] = summary_row[1], final_date
            summary_row[3] = f"{prefix}BK.{final_date.strftime('%d.%m')}.{summary_suffix_map.get(product, '')}"
            summary_row[4] = first_data['mau_so'] + first_data['ky_hieu']
            summary_row[5] = f"Xuất bán hàng theo hóa đơn số {summary_row[3]}"
            summary_row[7], summary_row[6], summary_row[8], summary_row[9] = product, static_data['ma_hang_map'].get(product, ''), "Lít", ma_kho
            summary_row[12], summary_row[13], summary_row[14], summary_row[17] = round(total_sl, 3), first_data['don_gia'] - phi_bvmt, round(TH_goc), ma_thue
            summary_row[18], summary_row[19], summary_row[20], summary_row[21] = tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co
            summary_row[23] = static_data['vu_viec_map'].get(selected_chxd, {}).get(product, '')
            summary_row[36] = round(TT_goc)
            original_invoice_rows.append(summary_row)
            bvmt_rows.append(_create_hddt_bvmt_row(summary_row, phi_bvmt, static_data, khu_vuc))
        upsse_wb = _create_upsse_workbook_shared()
        for row_data in original_invoice_rows + bvmt_rows: upsse_wb.active.append(row_data)
        output_buffer = io.BytesIO()
        upsse_wb.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    static_data, error = _load_static_data_hddt("Data_HDDT.xlsx", "MaHH.xlsx", "DSKH.xlsx")
    if error: raise ValueError(error)
    bkhd_wb = load_workbook(io.BytesIO(file_content_bytes), data_only=True)
    bkhd_ws = bkhd_wb.active
    final_date = None
    if confirmed_date_str:
        final_date = datetime.strptime(confirmed_date_str, '%Y-%m-%d')
    else:
        unique_dates = set()
        for row in bkhd_ws.iter_rows(min_row=11, values_only=True):
            if _to_float_shared(row[8] if len(row) > 8 else None) > 0:
                date_val = row[20] if len(row) > 20 else None
                if isinstance(date_val, datetime):
                    unique_dates.add(date_val.date())
                elif isinstance(date_val, (int, float)):
                    try:
                        converted_date_obj = pd.to_datetime(date_val, unit='D', origin='1899-12-30').to_pydatetime()
                        unique_dates.add(converted_date_obj.date())
                    except (ValueError, TypeError): pass
                elif isinstance(date_val, str):
                    try: unique_dates.add(datetime.strptime(date_val, '%d/%m/%Y').date())
                    except ValueError: continue
        if not unique_dates: raise ValueError("Không tìm thấy dữ liệu hóa đơn hợp lệ nào trong file Bảng kê.")
        if len(unique_dates) > 1: raise ValueError("Công cụ chỉ chạy được khi bạn kết xuất hóa đơn trong 1 ngày duy nhất.")
        the_date = unique_dates.pop()
        if the_date.day > 12: final_date = datetime(the_date.year, the_date.month, the_date.day)
        else:
            date1, date2 = datetime(the_date.year, the_date.month, the_date.day), datetime(the_date.year, the_date.day, the_date.month)
            if date1 != date2: return {'choice_needed': True, 'options': [{'text': date1.strftime('%d/%m/%Y'), 'value': date1.strftime('%Y-%m-%d')}, {'text': date2.strftime('%d/%m/%Y'), 'value': date2.strftime('%Y-%m-%d')}]}
            final_date = date1
    all_rows = list(bkhd_ws.iter_rows(min_row=11, values_only=True))
    if price_periods == '1':
        suffix_map = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}
        return _generate_upsse_from_hddt_rows(all_rows, static_data, selected_chxd, final_date, suffix_map)
    else:
        if not new_price_invoice_number: raise ValueError("Vui lòng nhập 'Số hóa đơn đầu tiên của giá mới'.")
        split_index = -1
        for i, row in enumerate(all_rows):
            if str(row[19] or '').strip() == new_price_invoice_number:
                split_index = i
                break
        if split_index == -1: raise ValueError(f"Không tìm thấy hóa đơn số '{new_price_invoice_number}'.")
        suffix_map_old = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}
        suffix_map_new = {"Xăng E5 RON 92-II": "5", "Xăng RON 95-III": "6", "Dầu DO 0,05S-II": "7", "Dầu DO 0,001S-V": "8"}
        result_old = _generate_upsse_from_hddt_rows(all_rows[:split_index], static_data, selected_chxd, final_date, suffix_map_old)
        result_new = _generate_upsse_from_hddt_rows(all_rows[split_index:], static_data, selected_chxd, final_date, suffix_map_new)
        if not result_old and not result_new: raise ValueError("Không có dữ liệu hợp lệ trong cả hai giai đoạn giá.")
        if not result_old: return result_new
        if not result_new: return result_old
        output_old.seek(0); output_new.seek(0)
        return {'old': result_old, 'new': result_new}

# ==============================================================================
# HÀM ĐIỀU PHỐI CHÍNH
# ==============================================================================
def process_unified_file(file_content, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=None):
    file_content_bytes = file_content
    report_type = detect_report_type(file_content_bytes)
    if report_type == 'POS':
        return process_pos_report(file_content_bytes, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=confirmed_date_str)
    elif report_type == 'HDDT':
        return process_hddt_report(file_content_bytes, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=confirmed_date_str)
    else:
        raise ValueError("Không thể tự động nhận diện loại Bảng kê. Vui lòng kiểm tra lại file Excel bạn đã tải lên.")
