import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import openpyxl
import io
import random
import time as time_module

def str_to_timedelta(time_str):
    try:
        t = datetime.strptime(str(time_str).strip(), "%H:%M:%S")
        return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
    except:
        return timedelta(seconds=0)

def timedelta_to_str(td):
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def find_sequence(pool, target_seconds):
    files_30s = [f for f in pool if f['duration_secs'] == 30]
    if len(files_30s) == 0:
        return None, "Kho dữ liệu Excel không có file 30s nào để xếp 2 file đầu tiên!", 0
    
    # Chọn 2 file 30s đầu tiên (Cố gắng chọn 2 file khác nhau)
    first_file = random.choice(files_30s)
    other_30s = [f for f in files_30s if f['name'] != first_file['name']]
    second_file = random.choice(other_30s) if other_30s else first_file
    
    start_sum = 60
    if target_seconds < start_sum - 20:
        return None, f"Thời lượng yêu cầu quá ngắn, không đủ xếp 2 file 30s đầu!", 0
        
    # Xáo trộn random trước để tạo sự đa dạng, sau đó mới xếp ưu tiên file dài
    valid_pool = [f for f in pool if f['duration_secs'] > 0]
    random.shuffle(valid_pool)
    valid_pool.sort(key=lambda x: x['duration_secs'], reverse=True)
    
    best_path = None
    best_error = 999999
    
    start_time = time_module.time()
    
    def dfs(curr_sum, last_name, short_sum, path, usage_counts):
        nonlocal best_path, best_error
        
        # Ngắt thuật toán nếu tìm kiếm quá 2 giây để tránh treo ứng dụng
        if time_module.time() - start_time > 2.0:
            return best_error == 0
            
        error = curr_sum - target_seconds
        
        # CHẤP NHẬN SAI SỐ +- 20 GIÂY
        if -20 <= error <= 20:
            if abs(error) < best_error:
                best_error = abs(error)
                best_path = list(path)
            # Nếu vô tình khớp đúng 100% thì dừng luôn
            if best_error == 0:
                return True
                
        if curr_sum > target_seconds + 20:
            return False
            
        # RÀNG BUỘC: Tổng file <= 60s không được quá 15 phút (900s)
        if short_sum >= 900:
            return False
            
        for f in valid_pool:
            # RÀNG BUỘC: KHÔNG xếp 2 file giống hệt nhau đứng liền kề
            if f['name'] == last_name:
                continue
                
            dur = f['duration_secs']
            is_short = dur <= 60 
            
            if is_short and short_sum + dur >= 900:
                continue
                
            # RÀNG BUỘC HẠN CHẾ TRÙNG LẶP TOÀN CỤC: 
            # File ngắn chỉ được dùng max 3 lần, file dài max 2 lần
            limit = 3 if is_short else 2
            if usage_counts.get(f['name'], 0) >= limit:
                continue
                
            # Duyệt nhánh tiếp theo
            usage_counts[f['name']] = usage_counts.get(f['name'], 0) + 1
            path.append(f)
            
            if dfs(curr_sum + dur, f['name'], short_sum + (dur if is_short else 0), path, usage_counts):
                return True
                
            # Backtrack
            path.pop()
            usage_counts[f['name']] -= 1
            
        return False

    initial_uses = {first_file['name']: 1}
    initial_uses[second_file['name']] = initial_uses.get(second_file['name'], 0) + 1
    
    dfs(start_sum, second_file['name'], start_sum, [], initial_uses)
    
    if best_path is not None:
        actual_sum = start_sum + sum(f['duration_secs'] for f in best_path)
        sai_so = actual_sum - target_seconds
        return [first_file, second_file] + best_path, "Thành công", sai_so
    else:
        err_msg = ("Không tìm được tổ hợp nào khớp trong khoảng ±20s mà vẫn đảm bảo 'Tổng file ngắn < 15p'. "
                   "Hãy thêm nhiều file có thời lượng trung bình/dài vào file Excel mẫu để AI có nhiều lựa chọn hơn!")
        return None, err_msg, 0


# ==== GIAO DIỆN STREAMLIT ====
st.set_page_config(page_title="App Lịch Chèn Sự Cố", layout="wide")
st.title("Ứng dụng Lịch Chèn Sự Cố")
st.header("CÁC DỮ LIỆU NHẬP VÀO")
col1, col2, col3 = st.columns(3)
with col1: gio_input_cg = st.text_input("1. Giờ CG Phim hiện tại", "12:00:00")
with col2: gio_mat_tin_hieu = st.text_input("2. Giờ mất tín hiệu", "13:30:00")
with col3: gio_input_next = st.text_input("3. Giờ CG phim tiếp theo", "14:00:00")

td_cg = str_to_timedelta(gio_input_cg)
td_mat = str_to_timedelta(gio_mat_tin_hieu)
td_next = str_to_timedelta(gio_input_next)

tab1, tab2 = st.tabs(["I. MÀN INPUT", "II. MÀN OUTPUT"])

with tab1:
    diff = td_mat - td_cg
    if diff >= timedelta(hours=1):
        gio_chen = td_mat
        thoi_luong = td_next - td_mat
    else:
        gio_chen = td_cg
        thoi_luong = td_next - td_cg
    st.success(f"4. GIỜ CHÈN SỰ CỐ: **{timedelta_to_str(gio_chen)}**")
    st.success(f"5. THỜI LƯỢNG CHÈN: **{timedelta_to_str(thoi_luong)}**")
    if st.button("Áp dụng dữ liệu từ Màn INPUT cho File"):
        st.session_state['active_gio_chen'] = gio_chen
        st.session_state['active_thoi_luong'] = thoi_luong
        st.success("Đã chọn dữ liệu Màn INPUT!")

with tab2:
    gio_chen_out = td_mat + timedelta(minutes=5) - timedelta(hours=1)
    thoi_luong_out = td_next - gio_chen_out
    st.success(f"4. GIỜ CHÈN SỰ CỐ: **{timedelta_to_str(gio_chen_out)}**")
    st.success(f"5. THỜI LƯỢNG CHÈN: **{timedelta_to_str(thoi_luong_out)}**")
    if st.button("Áp dụng dữ liệu từ Màn OUTPUT cho File"):
        st.session_state['active_gio_chen'] = gio_chen_out
        st.session_state['active_thoi_luong'] = thoi_luong_out
        st.success("Đã chọn dữ liệu Màn OUTPUT!")

st.divider()
st.header("6. NHẬP FILE MẪU & XUẤT FILE")
if 'active_gio_chen' not in st.session_state:
    st.warning("Vui lòng bấm 'Áp dụng dữ liệu' ở Màn INPUT hoặc OUTPUT trước khi xuất file.")
else:
    active_gio = st.session_state['active_gio_chen']
    active_tl = st.session_state['active_thoi_luong']
    
    st.write(f"**Giờ chèn sự cố đang chọn:** {timedelta_to_str(active_gio)}")
    st.write(f"**Thời lượng chèn yêu cầu:** {timedelta_to_str(active_tl)}")
    
    uploaded_file = st.file_uploader("Tải lên file Excel mẫu (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        wb = openpyxl.load_workbook(uploaded_file)
        sheet = wb.active
        file_pool = []
        for row in range(7, sheet.max_row + 1):
            file_name = sheet.cell(row=row, column=3).value
            file_dur = sheet.cell(row=row, column=4).value
            file_type = sheet.cell(row=row, column=5).value
            if file_name and file_dur is not None:
                try:
                    if isinstance(file_dur, str):
                        dur_td = str_to_timedelta(file_dur)
                    elif isinstance(file_dur, timedelta):
                        dur_td = file_dur
                    elif isinstance(file_dur, time):
                        dur_td = timedelta(hours=file_dur.hour, minutes=file_dur.minute, seconds=file_dur.second)
                    else:
                        dur_td = str_to_timedelta(str(file_dur))
                except:
                    continue
                    
                sec = int(dur_td.total_seconds())
                if sec > 0:
                    file_pool.append({
                        'name': file_name,
                        'duration_str': timedelta_to_str(dur_td),
                        'duration_secs': sec,
                        'type': file_type if file_type else "CM"
                    })
        
        if st.button("7. XUẤT FILE CHÈN", type="primary"):
            target_secs = int(active_tl.total_seconds())
            with st.spinner("Đang tính toán... (Ưu tiên File dài, không lặp, cho phép ±20s)"):
                sequence, msg, sai_so = find_sequence(file_pool, target_secs)
                
            if sequence is None:
                st.error(msg)
            else:
                st.success("🎉 Tuyệt vời! Đã tìm được tổ hợp file chèn tối ưu nhất!")
                
                sheet["B2"] = datetime.now().strftime("%d/%m/%Y")
                sheet["A6"] = f"{timedelta_to_str(active_gio)}"
                
                # Cập nhật tổng thời lượng thực tế AI tính ra vào D6 (Có thể lệch +- 20s so với yêu cầu)
                thuc_te_secs = target_secs + sai_so
                sheet["D6"] = f"{timedelta_to_str(timedelta(seconds=thuc_te_secs))}"
                
                sheet.delete_rows(7, sheet.max_row)
                for i, item in enumerate(sequence):
                    row_idx = 7 + i
                    sheet.cell(row=row_idx, column=3, value=item['name'])
                    sheet.cell(row=row_idx, column=4, value=item['duration_str'])
                    sheet.cell(row=row_idx, column=5, value=item['type'])
                    
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.download_button(
                    label="⬇️ BẤM VÀO ĐÂY TẢI XUỐNG FILE CHÈN HOÀN CHỈNH",
                    data=output,
                    file_name=f"FILE_CHEN_XUAT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # HIỂN THỊ CẢNH BÁO SAI SỐ TRỰC QUAN NGAY DƯỚI NÚT TẢI
                if sai_so != 0:
                    trang_thai = "THỪA" if sai_so > 0 else "THIẾU"
                    st.warning(f"⚠️ **CẢNH BÁO SAI SỐ:** Tổ hợp chèn xuất ra bị **{trang_thai} {abs(sai_so)} giây** so với yêu cầu ban đầu do cơ chế làm tròn ±20s.")
