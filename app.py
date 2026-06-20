import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import io
import random
import time

# ================= CÁC HÀM XỬ LÝ THỜI GIAN =================
def str_to_timedelta(time_str):
    """Chuyển đổi chuỗi HH:MM:SS hoặc HH:MM:SS:FF sang timedelta"""
    try:
        # Ép kiểu về chuỗi và xoá khoảng trắng thừa
        s = str(time_str).strip()
        # Nếu chuỗi dài hơn 8 ký tự (ví dụ: 00:03:00:00), chỉ lấy 8 ký tự đầu tiên
        if len(s) > 8:
            s = s[:8]
            
        t = datetime.strptime(s, "%H:%M:%S")
        return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
    except:
        return timedelta(seconds=0)

def timedelta_to_str(td):
    """Chuyển đổi timedelta sang chuỗi HH:MM:SS"""
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

# ================= HÀM XỬ LÝ GHÉP FILE SỰ CỐ =================
def find_sequence(pool, target_seconds, require_30s_first=True, no_adjacent_dup=True):
    """Tìm chuỗi các file sao cho tổng thời lượng đúng bằng target_seconds"""
    start_time = time.time()
    files_30s = [f for f in pool if f['duration_secs'] == 30]
    
    if require_30s_first and len(files_30s) == 0:
        return None, "Kho dữ liệu không có file 30s nào để xếp 2 vị trí đầu. Vui lòng bỏ chọn điều kiện '2 file đầu 30s'."

    # Sử dụng thuật toán Random Greedy thử nhiều tổ hợp trong tối đa 3 giây để chống treo App
    while time.time() - start_time < 3.0:
        seq = []
        current_sum = 0
        
        if require_30s_first:
            first_file = random.choice(files_30s)
            second_file = random.choice([f for f in files_30s if f['name'] != first_file['name']] or files_30s)
            seq = [first_file, second_file]
            current_sum = 60
            
            if current_sum == target_seconds:
                return seq, "Thành công"
            elif current_sum > target_seconds:
                return None, "Thời lượng chèn quá ngắn (dưới 60s), không thể chèn 2 file 30s. Vui lòng bỏ chọn điều kiện '2 file đầu 30s'."
                
        # Random pick các file tiếp theo để lắp đầy thời gian
        available_pool = list(pool)
        while current_sum < target_seconds:
            # Lọc các file nhét vừa khoảng trống thời gian còn lại
            valid_files = [f for f in available_pool if current_sum + f['duration_secs'] <= target_seconds]
            
            # Áp dụng tuỳ chọn: Lọc bỏ các file giống file liền kề trước đó
            if no_adjacent_dup and len(seq) > 0:
                valid_files = [f for f in valid_files if f['name'] != seq[-1]['name']]
                
            if not valid_files:
                break # Cụt đường (không có file nào ghép vừa), thoát vòng lặp nhỏ để random lại từ đầu
                
            chosen = random.choice(valid_files)
            seq.append(chosen)
            current_sum += chosen['duration_secs']
            
            if current_sum == target_seconds:
                return seq, "Thành công"
                
    return None, "App đã thử nhiều tổ hợp nhưng không thể khớp chính xác TỪNG GIÂY. Bạn hãy thử:\n1. Bỏ chọn các 'Điều kiện ghép file' ở bên trên.\n2. Bổ sung thêm các file thời lượng ngắn (5s, 10s, 15s) vào Excel mẫu."

# ================= GIAO DIỆN APP =================
st.set_page_config(page_title="App Lịch Chèn Sự Cố", layout="wide")
st.title("Ứng dụng Lịch Chèn Sự Cố")

# Nhập liệu chung
st.header("CÁC DỮ LIỆU NHẬP VÀO")
col1, col2, col3 = st.columns(3)
with col1:
    gio_input_cg = st.text_input("1. Giờ Input (CG) phim đang phát (HH:MM:SS)", "12:00:00")
with col2:
    gio_mat_tin_hieu = st.text_input("2. Giờ mất tín hiệu (HH:MM:SS)", "13:30:00")
with col3:
    gio_input_next = st.text_input("3. Giờ Input (CG) phim tiếp theo (HH:MM:SS)", "14:00:00")

# Chuyển đổi sang timedelta
td_cg = str_to_timedelta(gio_input_cg)
td_mat = str_to_timedelta(gio_mat_tin_hieu)
td_next = str_to_timedelta(gio_input_next)

# Tabs Màn INPUT và Màn OUTPUT
tab1, tab2 = st.tabs(["I. MÀN INPUT", "II. MÀN OUTPUT"])

gio_chen_final = ""
thoi_luong_final = ""

with tab1:
    st.subheader("Cách tính cho Màn INPUT")
    diff = td_mat - td_cg
    
    if diff >= timedelta(hours=1):
        st.info("Trường hợp 1: Giờ mất tín hiệu - Giờ Input (CG) >= 1 giờ")
        gio_chen = td_mat
        thoi_luong = td_next - td_mat
    else:
        st.info("Trường hợp 2: Giờ mất tín hiệu - Giờ Input (CG) < 1 giờ")
        gio_chen = td_cg
        thoi_luong = td_next - td_cg
        
    st.success(f"4. GIỜ CHÈN SỰ CỐ: **{timedelta_to_str(gio_chen)}**")
    st.success(f"5. THỜI LƯỢNG CHÈN: **{timedelta_to_str(thoi_luong)}**")
    
    if st.button("Áp dụng dữ liệu từ Màn INPUT cho File"):
        st.session_state['active_gio_chen'] = gio_chen
        st.session_state['active_thoi_luong'] = thoi_luong
        st.success("Đã chọn dữ liệu Màn INPUT!")

with tab2:
    st.subheader("Cách tính cho Màn OUTPUT")
    # Giờ chèn = Giờ mất + 5 phút - 1 giờ
    gio_chen_out = td_mat + timedelta(minutes=5) - timedelta(hours=1)
    thoi_luong_out = td_next - gio_chen_out
    
    st.success(f"4. GIỜ CHÈN SỰ CỐ: **{timedelta_to_str(gio_chen_out)}**")
    st.success(f"5. THỜI LƯỢNG CHÈN: **{timedelta_to_str(thoi_luong_out)}**")
    
    if st.button("Áp dụng dữ liệu từ Màn OUTPUT cho File"):
        st.session_state['active_gio_chen'] = gio_chen_out
        st.session_state['active_thoi_luong'] = thoi_luong_out
        st.success("Đã chọn dữ liệu Màn OUTPUT!")

# ================= XỬ LÝ FILE EXCEL =================
st.divider()
st.header("6. NHẬP FILE EXCEL MẪU & XUẤT FILE CHÈN")

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
        
        # Đọc kho dữ liệu file từ dòng 7 trở đi
        file_pool = []
        for row in range(7, sheet.max_row + 1):
            file_name = sheet.cell(row=row, column=3).value
            file_dur = sheet.cell(row=row, column=4).value
            file_type = sheet.cell(row=row, column=5).value
            
            if file_name and file_dur:
                if isinstance(file_dur, str):
                    dur_td = str_to_timedelta(file_dur)
                else:
                    dur_td = timedelta(hours=file_dur.hour, minutes=file_dur.minute, seconds=file_dur.second)
                
                sec = int(dur_td.total_seconds())
                if sec > 0:
                    file_pool.append({
                        'name': file_name,
                        'duration_str': timedelta_to_str(dur_td),
                        'duration_secs': sec,
                        'type': file_type if file_type else "CM"
                    })
        
        st.info(f"Đã đọc được {len(file_pool)} file chèn từ Excel mẫu để làm kho dữ liệu.")
        
        # --- THÊM TUỲ CHỌN ĐIỀU KIỆN TẠI ĐÂY ---
        st.markdown("### ⚙️ TUỲ CHỈNH ĐIỀU KIỆN GHÉP FILE")
        st.caption("*(Nếu App báo lỗi không ghép được do kho dữ liệu không đủ mảnh ghép thời lượng, hãy thử bỏ chọn các ô dưới đây và bấm Xuất lại)*")
        
        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            req_30s = st.checkbox("Bắt buộc 2 file đầu tiên phải dài 30s", value=True)
        with col_opt2:
            no_dup = st.checkbox("Không cho phép các file giống nhau xếp liền kề", value=True)
        # ---------------------------------------
        
        if st.button("7. XUẤT FILE CHÈN", type="primary"):
            target_secs = int(active_tl.total_seconds())
            
            with st.spinner("Đang tính toán ghép file sao cho khớp thời gian..."):
                sequence, msg = find_sequence(file_pool, target_secs, require_30s_first=req_30s, no_adjacent_dup=no_dup)
                
            if sequence is None:
                st.error(msg)
            else:
                st.success("Đã tìm được tổ hợp file chèn khớp hoàn toàn!")
                
                sheet["B2"] = datetime.now().strftime("%d/%m/%Y")
                sheet["A6"] = timedelta_to_str(active_gio)
                sheet["D6"] = timedelta_to_str(active_tl)
                
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
                    label="Tải xuống File Chèn hoàn chỉnh",
                    data=output,
                    file_name=f"FILE_CHEN_XUAT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
