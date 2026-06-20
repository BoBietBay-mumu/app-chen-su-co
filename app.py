import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import openpyxl
import io
import random
from collections import deque

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
        return None, "Kho dữ liệu không có file 30s nào để xếp 2 file đầu!"
    
    first_file = random.choice(files_30s)
    second_file = random.choice([f for f in files_30s if f['name'] != first_file['name']] or files_30s)
    
    start_sum = 60
    if target_seconds < start_sum:
        return None, f"Thời lượng yêu cầu quá ngắn ({target_seconds}s), không đủ xếp 2 file 30s đầu tiên!"
    if target_seconds == start_sum:
        return [first_file, second_file], "Thành công"
        
    valid_pool = [f for f in pool if f['duration_secs'] > 0]
    
    # HÀM ƯU TIÊN: Ưu tiên file 20-30 phút (1200s - 1800s), sau đó đến file dài khác, cuối cùng mới đến file ngắn
    def sort_priority(x):
        dur = x['duration_secs']
        if 1200 <= dur <= 1800:
            return (2, dur) # Ưu tiên cao nhất
        elif dur > 60:
            return (1, dur) # Ưu tiên trung bình
        else:
            return (0, dur) # Ít ưu tiên (các file ngắn)

    sorted_pool = sorted(valid_pool, key=sort_priority, reverse=True)
    
    # Cấu trúc hàng đợi (Queue): (Tổng_hiện_tại, Tên_file_cuối, Danh_sách_file, Tổng_thời_lượng_file_ngắn)
    # File ngắn được định nghĩa là từ 30s đến 60s
    queue = deque([(start_sum, second_file['name'], [first_file, second_file], start_sum)])
    
    # Visited dict mapping (current_sum, last_file_name) -> min_short_sum (lưu lại tổng file ngắn nhỏ nhất để tối ưu)
    visited = {(start_sum, second_file['name']): start_sum}
    
    while queue:
        curr_sum, last_name, path, short_sum = queue.popleft()
        
        for f in sorted_pool:
            if f['name'] == last_name:
                continue
                
            nxt_sum = curr_sum + f['duration_secs']
            
            # Cắt nhánh: nếu tổng vượt quá yêu cầu thì bỏ qua luôn
            if nxt_sum > target_seconds:
                continue
                
            # Kiểm tra và tính toán tổng thời lượng các file ngắn (30s - 1m)
            is_short_file = 30 <= f['duration_secs'] <= 60
            nxt_short_sum = short_sum + f['duration_secs'] if is_short_file else short_sum
            
            # RÀNG BUỘC: Tổng thời lượng các file 30s đến 1 phút phải nhỏ hơn 15 phút (900s)
            if nxt_short_sum >= 900:
                continue
                
            if nxt_sum == target_seconds:
                return path + [f], "Thành công"
                
            if nxt_sum < target_seconds:
                state_key = (nxt_sum, f['name'])
                # Cập nhật visited nếu chưa từng tới trạng thái này hoặc tìm được cách tới với ít file ngắn hơn
                if state_key not in visited or nxt_short_sum < visited[state_key]:
                    visited[state_key] = nxt_short_sum
                    queue.append((nxt_sum, f['name'], path + [f], nxt_short_sum))
                    
    return None, "Không thể ghép chính xác thời lượng hoặc vi phạm giới hạn file quảng cáo ngắn. Hãy thử thêm file có thời lượng đa dạng hơn vào Excel mẫu!"

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
    st.write(f"**Thời lượng chèn đang chọn:** {timedelta_to_str(active_tl)}")
    
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
            with st.spinner("Đang tính toán thuật toán AI (Bao gồm ưu tiên ghép các file dài 20-30 phút)..."):
                sequence, msg = find_sequence(file_pool, target_secs)
            if sequence is None:
                st.error(msg)
            else:
                st.success("Tuyệt vời! Đã tìm được tổ hợp file chèn khớp 100% thời gian!")
                sheet["B2"] = datetime.now().strftime("%d/%m/%Y")
                sheet["A6"] = f"{timedelta_to_str(active_gio)}"
                sheet["D6"] = f"{timedelta_to_str(active_tl)}"
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
