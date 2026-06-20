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

def find_sequence(pool, target_seconds, max_sai_so, max_lap):
    valid_pool = [f for f in pool if f['duration_secs'] > 0]
    if not valid_pool:
        return None, "Kho dữ liệu Excel của bạn đang trống hoặc không đọc được thời lượng!", 0
        
    # Xáo trộn random để tạo sự đa dạng giữa các lần xuất
    random.shuffle(valid_pool)
    
    # Sắp xếp mặc định: Ưu tiên file dài lên trước để lấp đầy nhanh chóng
    valid_pool.sort(key=lambda x: x['duration_secs'], reverse=True)
    
    # Tạo một Pool đặc biệt để xử lý yêu cầu: "Ưu tiên 2 file < 60s ở đầu"
    short_files_pool = [f for f in valid_pool if f['duration_secs'] <= 60]
    long_files_pool = [f for f in valid_pool if f['duration_secs'] > 60]
    priority_pool = short_files_pool + long_files_pool # Đưa file ngắn lên đầu mâm
    
    best_path = None
    best_error = 999999
    
    start_time = time_module.time()
    
    def dfs(curr_sum, last_name, short_sum, path, usage_counts):
        nonlocal best_path, best_error
        
        # Ngắt thuật toán nếu tìm kiếm quá 3 giây để app không bị treo
        if time_module.time() - start_time > 3.0:
            return best_error <= max_sai_so
            
        error = curr_sum - target_seconds
        
        # Kiểm tra sai số theo giới hạn người dùng chọn
        if -max_sai_so <= error <= max_sai_so:
            if abs(error) < best_error:
                best_error = abs(error)
                best_path = list(path)
            if best_error == 0:
                return True
                
        if curr_sum > target_seconds + max_sai_so:
            return False
            
        # RÀNG BUỘC: Tổng file <= 60s không được quá 15 phút (900s)
        if short_sum >= 900:
            return False
            
        # NẾU ĐANG CHỌN 2 FILE ĐẦU TIÊN: Dùng mâm ưu tiên file ngắn
        # TỪ FILE THỨ 3 TRỞ ĐI: Dùng mâm ưu tiên file dài
        current_pool = priority_pool if len(path) < 2 else valid_pool
            
        for f in current_pool:
            # Ràng buộc: KHÔNG xếp 2 file giống hệt nhau liền kề
            if f['name'] == last_name:
                continue
                
            dur = f['duration_secs']
            is_short = dur <= 60 
            
            if is_short and short_sum + dur >= 900:
                continue
                
            # Ràng buộc: Giới hạn lặp lại
            limit = max_lap if is_short else 2
            if usage_counts.get(f['name'], 0) >= limit:
                continue
                
            # Ghi nhận và duyệt nhánh tiếp theo
            usage_counts[f['name']] = usage_counts.get(f['name'], 0) + 1
            path.append(f)
            
            if dfs(curr_sum + dur, f['name'], short_sum + (dur if is_short else 0), path, usage_counts):
                return True
                
            # Backtrack (Hoàn tác nếu nhánh này đi vào ngõ cụt)
            path.pop()
            usage_counts[f['name']] -= 1
            
        return False

    # Chạy thuật toán tìm kiếm từ trạng thái rỗng
    dfs(0, None, 0, [], {})
    
    if best_path is not None:
        actual_sum = sum(f['duration_secs'] for f in best_path)
        sai_so = actual_sum - target_seconds
        return best_path, "Thành công", sai_so
    else:
        err_msg = (f"Không tìm được tổ hợp nào khớp trong khoảng ±{max_sai_so}s mà vẫn đảm bảo 'Tổng file ngắn < 15p'.\n"
                   f"👉 Kéo thanh 'Sai số tối đa' lên cao hơn (ví dụ: 60s) hoặc thêm nhiều file có thời lượng đa dạng vào Excel!")
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
st.header("6. CẤU HÌNH AI & XUẤT FILE")

if 'active_gio_chen' not in st.session_state:
    st.warning("Vui lòng bấm 'Áp dụng dữ liệu' ở Màn INPUT hoặc OUTPUT trước khi xuất file.")
else:
    active_gio = st.session_state['active_gio_chen']
    active_tl = st.session_state['active_thoi_luong']
    
    st.write(f"**Giờ chèn đang chọn:** {timedelta_to_str(active_gio)} | **Thời lượng yêu cầu:** {timedelta_to_str(active_tl)}")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        user_sai_so = st.slider("Sai số tối đa cho phép (giây) - Kéo lên nếu bị lỗi", min_value=0, max_value=120, value=20, step=5)
    with col_c2:
        user_max_lap = st.slider("Số lần lặp tối đa của 1 file ngắn", min_value=2, max_value=10, value=3, step=1)
    
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
            with st.spinner("Đang xử lý thuật toán AI... (Ưu tiên File <60s ở đầu)"):
                sequence, msg, sai_so = find_sequence(file_pool, target_secs, user_sai_so, user_max_lap)
                
            if sequence is None:
                st.error(msg)
            else:
                st.success("🎉 Tuyệt vời! Đã tìm được tổ hợp tối ưu nhất!")
                
                sheet["B2"] = datetime.now().strftime("%d/%m/%Y")
                sheet["A6"] = f"{timedelta_to_str(active_gio)}"
                
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
                
                if sai_so != 0:
                    trang_thai = "THỪA" if sai_so > 0 else "THIẾU"
                    st.warning(f"⚠️ **CẢNH BÁO SAI SỐ:** File chèn thực tế bị **{trang_thai} {abs(sai_so)} giây** so với yêu cầu ban đầu.")
