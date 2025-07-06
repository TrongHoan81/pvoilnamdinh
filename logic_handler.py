import pandas as pd
from openpyxl import load_workbook, Workbook
from datetime import datetime
import re
import io

# ==============================================================================
# CÁC HÀM TIỆN ÍCH CHUNG
# ==============================================================================
def clean_string(s):
    if s is None: return ""
    cleaned_s = str(s).strip()
    if cleaned_s.startswith("'"): cleaned_s = cleaned_s[1:]
    return re.sub(r'\s+', ' ', cleaned_s)

def to_float(value):
    if value is None: return 0.0
    try:
        # Hỗ trợ cả dấu phẩy và không có dấu phẩy
        return float(str(value).replace(',', '').strip())
    except (ValueError, TypeError):
        return 0.0

def format_tax_code(raw_vat_value):
    if raw_vat_value is None: return ""
    try:
        s_value = str(raw_vat_value).replace('%', '').strip()
        f_value = float(s_value)
        if 0 < f_value < 1: f_value *= 100
        return f"{round(f_value):02d}"
    except (ValueError, TypeError):
        return ""

def _create_upsse_workbook():
    headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
    wb = Workbook()
    ws = wb.active
    for _ in range(4): ws.append([''] * len(headers))
    ws.append(headers)
    return wb

# ==============================================================================
# HÀM NHẬN DIỆN LOẠI BẢNG KÊ
# ==============================================================================
def detect_report_type(worksheet):
    try:
        pos_headers = {clean_string(cell.value) for cell in worksheet[5]}
        hddt_headers = {clean_string(cell.value) for cell in worksheet[10]}
        pos_characteristic_cols = {'Séri', 'Số', 'Đơn giá(đã có thuế GTGT)'}
        hddt_characteristic_cols = {'Số công văn (Số tham chiếu)', 'Mã KH (FAST)', 'MST khách hàng'}
        if pos_characteristic_cols.issubset(pos_headers): return 'POS'
        if hddt_characteristic_cols.issubset(hddt_headers): return 'HDDT'
    except IndexError: pass
    try:
        if clean_string(worksheet['B5'].value) == 'BẢNG THỐNG KÊ HÓA ĐƠN': return 'POS'
        if clean_string(worksheet['A1'].value) == 'BÁO CÁO': return 'HDDT'
    except Exception: return 'UNKNOWN'
    return 'UNKNOWN'

# ==============================================================================
# LOGIC RIÊNG CHO BẢNG KÊ POS
# ==============================================================================
def load_static_data_pos():
    """Hàm đọc file Data_POS.xlsx và các file cấu hình khác cho POS."""
    try:
        wb = load_workbook("Data_POS.xlsx", data_only=True)
        ws = wb.active
        chxd_detail_map = {}
        store_specific_x_lookup = {}
        for row_idx in range(4, ws.max_row + 1):
            row_values = [cell.value for cell in ws[row_idx]]
            if len(row_values) < 18: continue
            chxd_name = clean_string(row_values[3])
            if not chxd_name: continue
            chxd_detail_map[chxd_name] = {
                'g4_val': clean_string(row_values[6]), 'g5_val': clean_string(row_values[4]),
                't_val': clean_string(row_values[19]), 'u_val': clean_string(row_values[20]),
                'v_val': clean_string(row_values[21]), 'w_val': clean_string(row_values[22]),
                'x_val': clean_string(row_values[23])
            }
            store_specific_x_lookup[chxd_name] = {
                clean_string(ws['O3'].value).lower(): clean_string(row_values[14]),
                clean_string(ws['P3'].value).lower(): clean_string(row_values[15]),
                clean_string(ws['Q3'].value).lower(): clean_string(row_values[16]),
                clean_string(ws['R3'].value).lower(): clean_string(row_values[17])
            }
        tmt_map = {clean_string(ws[f'A{i}'].value): to_float(ws[f'B{i}'].value) for i in range(4, 8)}
        return {
            "chxd_details": chxd_detail_map,
            "x_lookup": store_specific_x_lookup,
            "tmt_map": tmt_map,
            "listbox_data": sorted(chxd_detail_map.keys())
        }, None
    except FileNotFoundError:
        return None, "Lỗi: Không tìm thấy file Data_POS.xlsx."
    except Exception as e:
        return None, f"Lỗi khi đọc file Data_POS.xlsx: {e}"

def process_pos_report(file_content, selected_chxd, price_periods, new_price_invoice_number, **kwargs):
    """Hàm xử lý cho file Bảng kê từ POS."""
    static_data, error = load_static_data_pos()
    if error: raise ValueError(error)

    df = pd.read_excel(io.BytesIO(file_content), header=4)
    df.columns = [clean_string(col) for col in df.columns]
    df = df.dropna(how='all')
    df = df[to_float(df['Số lượng']) > 0]
    if df.empty: raise ValueError("Không có dữ liệu hợp lệ (Số lượng > 0) trong file tải lên.")
    
    final_date = pd.to_datetime(df['Ngày'].iloc[0]).strftime('%d/%m/%Y')
    chxd_details = static_data['chxd_details'].get(selected_chxd)
    if not chxd_details: raise ValueError(f"Không tìm thấy cấu hình cho CHXD: {selected_chxd}")

    def process_rows(dataframe, is_new_price):
        processed_rows = []
        summary_data = {}
        for _, row in dataframe.iterrows():
            product_name = clean_string(row['Hàng hóa'])
            if "không lấy hóa đơn" in clean_string(row['Tên khách hàng']):
                if product_name not in summary_data:
                    summary_data[product_name] = {'sl': 0, 'tt': 0, 'thue': 0, 'tienthuegtgt': 0, 'dongia': to_float(row['Đơn giá(đã có thuế GTGT)'])}
                summary_data[product_name]['sl'] += to_float(row['Số lượng'])
                summary_data[product_name]['tt'] += to_float(row['Tổng tiền thanh toán'])
                summary_data[product_name]['thue'] += to_float(row['Thuế GTGT'])
                summary_data[product_name]['tienthuegtgt'] += to_float(row['Tiền thuế GTGT'])
            else:
                processed_rows.append(create_pos_invoice_row(row, final_date, chxd_details, static_data))
        
        for product, data in summary_data.items():
            processed_rows.append(create_pos_summary_row(product, data, final_date, chxd_details, static_data, is_new_price, selected_chxd))
        
        final_df_rows = []
        for row_data in processed_rows:
            final_df_rows.append(row_data)
            tmt_value = static_data['tmt_map'].get(row_data[7], 0)
            if tmt_value > 0:
                final_df_rows.append(create_pos_tmt_row(row_data, tmt_value, chxd_details))
        return final_df_rows

    all_rows_for_df = []
    if price_periods == '1':
        all_rows_for_df.extend(process_rows(df, is_new_price=False))
    else: # 2 periods
        if not new_price_invoice_number: raise ValueError("Vui lòng nhập số hóa đơn đầu tiên của giá mới.")
        split_idx = df.index[df['Số'] == to_float(new_price_invoice_number)].tolist()
        if not split_idx: raise ValueError(f"Không tìm thấy số hóa đơn '{new_price_invoice_number}'")
        
        df_old = df.loc[:split_idx[0]-1]
        df_new = df.loc[split_idx[0]:]
        
        old_rows = process_rows(df_old, is_new_price=False)
        new_rows = process_rows(df_new, is_new_price=True)
        
        wb_old = _create_upsse_workbook()
        for r in old_rows: wb_old.active.append(r)
        output_old = io.BytesIO()
        wb_old.save(output_old)

        wb_new = _create_upsse_workbook()
        for r in new_rows: wb_new.active.append(r)
        output_new = io.BytesIO()
        wb_new.save(output_new)
        
        return {'old': output_old, 'new': output_new}

    final_wb = _create_upsse_workbook()
    for r in all_rows_for_df: final_wb.active.append(r)
    output = io.BytesIO()
    final_wb.save(output)
    output.seek(0)
    return output

def create_pos_invoice_row(row, final_date, details, static_data):
    new_row = [''] * 37
    new_row[0] = details['g4_val']
    new_row[1] = new_row[31] = clean_string(row['Tên khách hàng'])
    new_row[2] = final_date
    new_row[3] = f"{clean_string(row['Séri'])}{int(to_float(row['Số']))}"
    new_row[4] = clean_string(row['Séri'])
    new_row[5] = f"Xuất bán hàng theo hóa đơn số {new_row[3]}"
    new_row[7] = clean_string(row['Hàng hóa'])
    new_row[8] = "Lít"
    new_row[9] = details['g5_val']
    new_row[12] = to_float(row['Số lượng'])
    tmt_value = static_data['tmt_map'].get(new_row[7], 0)
    new_row[13] = to_float(row['Đơn giá(đã có thuế GTGT)']) - tmt_value
    new_row[14] = to_float(row['Tổng tiền thanh toán']) - to_float(row['Tiền thuế GTGT']) - round(tmt_value * new_row[12], 0)
    new_row[17] = format_tax_code(to_float(row['Thuế GTGT']))
    new_row[18], new_row[19], new_row[20], new_row[21] = details['t_val'], details['u_val'], details['v_val'], details['w_val']
    new_row[33] = clean_string(row['Mã số thuế'])
    new_row[36] = to_float(row['Tiền thuế GTGT']) - round(tmt_value * new_row[12] * (to_float(new_row[17]) / 100.0), 0)
    return new_row

def create_pos_summary_row(product_name, data, final_date, details, static_data, is_new_price, selected_chxd):
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
    new_row[17] = format_tax_code(data['thue'] * 100)
    new_row[18], new_row[19], new_row[20], new_row[21] = details['t_val'], details['u_val'], details['v_val'], details['w_val']
    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"
    new_row[36] = data['tienthuegtgt'] - round(tmt_value * data['sl'] * (to_float(new_row[17]) / 100.0), 0)
    return new_row

def create_pos_tmt_row(original_row, tmt_value, details):
    tmt_row = list(original_row)
    tax_rate_decimal = to_float(original_row[17]) / 100.0
    tmt_row[6], tmt_row[7], tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "Lít"
    tmt_row[9] = details['g5_val']
    tmt_row[13] = tmt_value
    tmt_row[14] = round(tmt_value * to_float(original_row[12]), 0)
    tmt_row[18], tmt_row[19], tmt_row[20], tmt_row[21] = details['t_val'], details['u_val'], details['v_val'], details['w_val'] # Re-check these against data file
    tmt_row[36] = round(tmt_row[14] * tax_rate_decimal)
    tmt_row[5], tmt_row[31], tmt_row[32], tmt_row[33] = '', '', '', ''
    return tmt_row

# ==============================================================================
# LOGIC RIÊNG CHO BẢNG KÊ HDDT
# ==============================================================================
def load_static_data_hddt():
    """Hàm đọc file Data_HDDT.xlsx và các file cấu hình khác cho HDDT."""
    try:
        static_data = {}
        # --- Đọc file Data_HDDT.xlsx ---
        wb = load_workbook("Data_HDDT.xlsx", data_only=True)
        ws = wb.active
        chxd_list, tk_mk_map, khhd_map, chxd_to_khuvuc_map = [], {}, {}, {}
        vu_viec_map = {}
        vu_viec_headers = [clean_string(cell.value) for cell in ws[2][4:9]]
        for row_values in ws.iter_rows(min_row=3, max_col=12, values_only=True):
            chxd_name = clean_string(row_values[3])
            if chxd_name:
                ma_kho, khhd, khu_vuc = clean_string(row_values[9]), clean_string(row_values[10]), clean_string(row_values[11])
                if chxd_name not in tk_mk_map: chxd_list.append(chxd_name)
                if ma_kho: tk_mk_map[chxd_name] = ma_kho
                if khhd: khhd_map[chxd_name] = khhd
                if khu_vuc: chxd_to_khuvuc_map[chxd_name] = khu_vuc
                vu_viec_map[chxd_name] = {}
                vu_viec_data_row = row_values[4:9]
                for i, header in enumerate(vu_viec_headers):
                    if header:
                        key = "Dầu mỡ nhờn" if i == len(vu_viec_headers) - 1 else header
                        vu_viec_map[chxd_name][key] = clean_string(vu_viec_data_row[i])
        if not chxd_list: return None, "Không tìm thấy Tên CHXD nào trong cột D của file Data_HDDT.xlsx."
        static_data.update({"DS_CHXD": chxd_list, "tk_mk": tk_mk_map, "khhd_map": khhd_map, "chxd_to_khuvuc_map": chxd_to_khuvuc_map, "vu_viec_map": vu_viec_map})
        def get_lookup_map(min_r, max_r, min_c=1, max_c=2):
            return {clean_string(row[0]): row[1] for row in ws.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=max_c, values_only=True) if row[0] and row[1] is not None}
        phi_bvmt_map_raw = get_lookup_map(10, 13)
        static_data["phi_bvmt_map"] = {k: to_float(v) for k, v in phi_bvmt_map_raw.items()}
        static_data.update({
            "tk_no_map": get_lookup_map(29, 31), "tk_doanh_thu_map": get_lookup_map(33, 35),
            "tk_thue_co_map": get_lookup_map(38, 40), "tk_gia_von_value": ws['B36'].value,
            "tk_no_bvmt_map": get_lookup_map(44, 46), "tk_dt_thue_bvmt_map": get_lookup_map(48, 50),
            "tk_gia_von_bvmt_value": ws['B51'].value, "tk_thue_co_bvmt_map": get_lookup_map(53, 55)
        })
        # --- Đọc file MaHH.xlsx ---
        wb_mahh = load_workbook("MaHH.xlsx", data_only=True)
        static_data["ma_hang_map"] = {clean_string(r[0]): clean_string(r[2]) for r in wb_mahh.active.iter_rows(min_row=2, max_col=3, values_only=True) if r[0] and r[2]}
        # --- Đọc file DSKH.xlsx ---
        wb_dskh = load_workbook("DSKH.xlsx", data_only=True)
        static_data["mst_to_makh_map"] = {clean_string(r[2]): clean_string(r[3]) for r in wb_dskh.active.iter_rows(min_row=2, max_col=4, values_only=True) if r[2]}
        return static_data, None
    except FileNotFoundError as e:
        return None, f"Lỗi: Không tìm thấy file cấu hình. Chi tiết: {e.filename}"
    except Exception as e:
        return None, f"Lỗi khi đọc file cấu hình: {e}"

def process_hddt_report(file_content, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=None):
    """Hàm xử lý cho file Bảng kê từ Hóa đơn điện tử."""
    static_data, error = load_static_data_hddt()
    if error: raise ValueError(error)

    bkhd_wb = load_workbook(io.BytesIO(file_content), data_only=True)
    bkhd_ws = bkhd_wb.active
    
    final_date = None
    if confirmed_date_str:
        final_date = datetime.strptime(confirmed_date_str, '%Y-%m-%d')
    else:
        unique_dates = set()
        for row in bkhd_ws.iter_rows(min_row=11, values_only=True):
            if to_float(row[8] if len(row) > 8 else None) > 0:
                date_val = row[20] if len(row) > 20 else None
                if isinstance(date_val, datetime): unique_dates.add(date_val.date())
        if len(unique_dates) > 1: raise ValueError("Công cụ chỉ chạy được khi bạn kết xuất hóa đơn trong 1 ngày duy nhất.")
        if not unique_dates: raise ValueError("Không tìm thấy dữ liệu hóa đơn hợp lệ nào trong file Bảng kê.")
        the_date = unique_dates.pop()
        if the_date.day > 12: final_date = datetime(the_date.year, the_date.month, the_date.day)
        else:
            date1, date2 = datetime(the_date.year, the_date.month, the_date.day), datetime(the_date.year, the_date.day, the_date.month)
            return {'choice_needed': True, 'options': [{'text': date1.strftime('%d/%m/%Y'), 'value': date1.strftime('%Y-%m-%d')}, {'text': date2.strftime('%d/%m/%Y'), 'value': date2.strftime('%Y-%m-%d')}]}

    all_rows = list(bkhd_ws.iter_rows(min_row=11, values_only=True))
    if price_periods == '1':
        suffix_map = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}
        return _generate_upsse_from_hddt_rows(all_rows, static_data, selected_chxd, final_date, suffix_map)
    else: # 2 periods
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
        return {'old': result_old, 'new': result_new}

def _generate_upsse_from_hddt_rows(rows_to_process, static_data, selected_chxd, final_date, summary_suffix_map):
    if not rows_to_process: return None
    khu_vuc, ma_kho = static_data['chxd_to_khuvuc_map'].get(selected_chxd), static_data['tk_mk'].get(selected_chxd)
    tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co = static_data['tk_no_map'].get(khu_vuc), static_data['tk_doanh_thu_map'].get(khu_vuc), static_data['tk_gia_von_value'], static_data['tk_thue_co_map'].get(khu_vuc)
    
    original_invoice_rows, bvmt_rows, summary_data = [], [], {}
    first_invoice_prefix_source = ""

    for bkhd_row in rows_to_process:
        if to_float(bkhd_row[8] if len(bkhd_row) > 8 else None) <= 0: continue
        ten_kh, ten_mat_hang = clean_string(bkhd_row[3]), clean_string(bkhd_row[6])
        is_anonymous, is_petrol = ("không lấy hóa đơn" in ten_kh), (ten_mat_hang in static_data['phi_bvmt_map'])
        
        if not is_anonymous or not is_petrol:
            new_upsse_row = [''] * 37
            new_upsse_row[9], new_upsse_row[1], new_upsse_row[31], new_upsse_row[2] = ma_kho, ten_kh, ten_kh, final_date
            so_hd_goc = str(bkhd_row[19] or '').strip()
            new_upsse_row[3] = f"HN{so_hd_goc[-6:]}" if selected_chxd == "Nguyễn Huệ" else f"{(str(bkhd_row[18] or '').strip())[-2:]}{so_hd_goc[-6:]}"
            new_upsse_row[4] = clean_string(bkhd_row[17]) + clean_string(bkhd_row[18])
            new_upsse_row[5], new_upsse_row[7], new_upsse_row[6] = f"Xuất bán hàng theo hóa đơn số {new_upsse_row[3]}", ten_mat_hang, static_data['ma_hang_map'].get(ten_mat_hang, '')
            new_upsse_row[8], new_upsse_row[12] = clean_string(bkhd_row[10]), round(to_float(bkhd_row[8]), 3)
            phi_bvmt = static_data['phi_bvmt_map'].get(ten_mat_hang, 0.0) if is_petrol else 0.0
            new_upsse_row[13] = to_float(bkhd_row[9]) - phi_bvmt
            ma_thue = format_tax_code(bkhd_row[14])
            new_upsse_row[17] = ma_thue
            thue_suat = to_float(ma_thue) / 100.0 if ma_thue else 0.0
            tien_thue_goc, so_luong = to_float(bkhd_row[15]), to_float(bkhd_row[8])
            tien_thue_phi_bvmt = round(phi_bvmt * so_luong * thue_suat)
            new_upsse_row[36] = round(tien_thue_goc - tien_thue_phi_bvmt)
            new_upsse_row[14] = round(to_float(bkhd_row[13]) if not is_petrol else to_float(bkhd_row[16]) - tien_thue_goc - round(phi_bvmt * so_luong))
            new_upsse_row[18], new_upsse_row[19], new_upsse_row[20], new_upsse_row[21] = tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co
            chxd_vu_viec_map = static_data['vu_viec_map'].get(selected_chxd, {})
            new_upsse_row[23] = chxd_vu_viec_map.get(ten_mat_hang, chxd_vu_viec_map.get("Dầu mỡ nhờn", ''))
            new_upsse_row[32], mst_khach_hang = clean_string(bkhd_row[4]), clean_string(bkhd_row[5])
            new_upsse_row[33] = mst_khach_hang
            ma_kh_fast = clean_string(bkhd_row[2])
            new_upsse_row[0] = ma_kh_fast if ma_kh_fast and len(ma_kh_fast) < 12 else static_data['mst_to_makh_map'].get(mst_khach_hang, ma_kho)
            original_invoice_rows.append(new_upsse_row)
            if is_petrol: bvmt_rows.append(_create_hddt_bvmt_row(new_upsse_row, phi_bvmt, static_data, khu_vuc))
        else: # Gom dữ liệu khách vãng lai
            if not first_invoice_prefix_source: first_invoice_prefix_source = str(bkhd_row[18] or '').strip()
            if ten_mat_hang not in summary_data:
                summary_data[ten_mat_hang] = {'sl': 0, 'thue': 0, 'phai_thu': 0, 'first_data': {'mau_so': clean_string(bkhd_row[17]),'ky_hieu': clean_string(bkhd_row[18]),'don_gia': to_float(bkhd_row[9]),'vat_raw': bkhd_row[14]}}
            summary_data[ten_mat_hang]['sl'] += to_float(bkhd_row[8])
            summary_data[ten_mat_hang]['thue'] += to_float(bkhd_row[15])
            summary_data[ten_mat_hang]['phai_thu'] += to_float(bkhd_row[16])

    prefix = first_invoice_prefix_source[-2:] if len(first_invoice_prefix_source) >= 2 else first_invoice_prefix_source
    for product, data in summary_data.items():
        summary_row = [''] * 37
        first_data, total_sl = data['first_data'], data['sl']
        phi_bvmt = static_data['phi_bvmt_map'].get(product, 0.0)
        ma_thue = format_tax_code(first_data['vat_raw'])
        thue_suat = to_float(ma_thue) / 100.0 if ma_thue else 0.0
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

    upsse_wb = _create_upsse_workbook()
    for row_data in original_invoice_rows + bvmt_rows: upsse_wb.active.append(row_data)
    output_buffer = io.BytesIO()
    upsse_wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

def _create_hddt_bvmt_row(original_row, phi_bvmt, static_data, khu_vuc):
    bvmt_row = list(original_row)
    so_luong = to_float(original_row[12])
    thue_suat = to_float(original_row[17]) / 100.0 if original_row[17] else 0.0
    bvmt_row[6], bvmt_row[7] = "TMT", "Thuế bảo vệ môi trường"
    bvmt_row[13], bvmt_row[14] = phi_bvmt, round(phi_bvmt * so_luong)
    bvmt_row[18] = static_data.get('tk_no_bvmt_map', {}).get(khu_vuc)
    bvmt_row[19] = static_data.get('tk_dt_thue_bvmt_map', {}).get(khu_vuc)
    bvmt_row[20] = static_data.get('tk_gia_von_bvmt_value')
    bvmt_row[21] = static_data.get('tk_thue_co_bvmt_map', {}).get(khu_vuc)
    bvmt_row[36] = round(phi_bvmt * so_luong * thue_suat)
    for i in [5, 31, 32, 33]: bvmt_row[i] = ''
    return bvmt_row

# ==============================================================================
# HÀM ĐIỀU PHỐI CHÍNH
# ==============================================================================
def process_unified_file(file_content, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=None):
    """
    Hàm điều phối chính: Nhận diện loại file và gọi hàm xử lý tương ứng.
    """
    try:
        workbook = load_workbook(io.BytesIO(file_content), data_only=True)
        worksheet = workbook.active
    except Exception as e:
        raise ValueError(f"Không thể đọc file Excel. File có thể bị hỏng hoặc không đúng định dạng. Lỗi: {e}")

    report_type = detect_report_type(worksheet)

    if report_type == 'POS':
        return process_pos_report(file_content, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=confirmed_date_str)
    
    elif report_type == 'HDDT':
        return process_hddt_report(file_content, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=confirmed_date_str)

    else:
        raise ValueError("Không thể tự động nhận diện loại Bảng kê. Vui lòng kiểm tra lại file Excel bạn đã tải lên.")
