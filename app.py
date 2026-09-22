import streamlit as st
import pandas as pd
import numpy as np
import random
import io

# ==========================================
# CẤU HÌNH GIAO DIỆN TRANG WEB
# ==========================================
st.set_page_config(page_title="Hệ Thống Phân Bổ Rải Hàng", page_icon="🛒", layout="wide")

# Áp dụng Custom CSS để làm đẹp giao diện
st.markdown("""
    <style>
    .main-title {
        font-size: 50px !important; 
        color: #FF4B4B;
        text-align: center;
        font-weight: 900;
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 24px !important; 
        color: #0068C9;
        text-align: center;
        font-weight: 600;
        margin-bottom: 30px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-size: 18px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: scale(1.02);
    }
    /* Làm to và in đậm tiêu đề của khung Hướng dẫn */
    [data-testid="stExpander"] summary p {
        font-size: 22px !important;
        font-weight: 900 !important;
        color: #D32F2F !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">🛒 HỆ THỐNG TỰ ĐỘNG PHÂN BỔ RẢI HÀNG</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Tối ưu hóa Khuyến Mãi & Cân bằng tải Vận hành</p>', unsafe_allow_html=True)

# ==========================================
# HƯỚNG DẪN SỬ DỤNG VÀ CHUẨN BỊ FILE
# ==========================================
with st.expander("📋 HƯỚNG DẪN CHUẨN BỊ FILE DỮ LIỆU (Bấm để xem chi tiết)", expanded=True):
    st.markdown("""
    Để hệ thống xử lý chính xác, file Excel tải lên **bắt buộc phải có các cột sau** (viết đúng chính tả, có dấu):
    *   **Kho**: Mã hoặc tên Kho xử lý.
    *   **Mã siêu thị**: Mã của Siêu thị nhận hàng.
    *   **Lịch về hàng**: Định dạng chuỗi số cách nhau bằng dấu phẩy (VD: `0,2,4,6`, `1,3,5` hoặc `0,1,2,3,4,5,6`).
    *   **Đợt KM**: Phân loại khuyến mãi (VD: `1 Lần/Tuần`, `2 Lần/Tuần`, `2 ngày KM/lần`...).
    *   **Quy cách mua**: Số lượng hàng hóa trong 1 quy cách/lô (kiểu số).
    *   **Tổng mới**: Tổng số lượng hàng cần rải trong tuần (kiểu số).
    
    *(Các cột thông tin khác như `Tên siêu thị`, `Mã sản phẩm`, `Tên sản phẩm` nên được giữ nguyên để thuận tiện cho việc trích xuất và đối chiếu kết quả sau khi rải).*
    """)

# ==========================================
# KHU VỰC TẢI FILE
# ==========================================
st.markdown("### 📂 1. Tải lên dữ liệu gốc")
uploaded_file = st.file_uploader("Vui lòng chọn file Excel (Định dạng .xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success("✅ Đã nhận file dữ liệu! Bạn có thể bắt đầu quá trình rải hàng.")
    
    st.markdown("### 🚀 2. Xử lý thuật toán")
    if st.button("BẮT ĐẦU RẢI SỐ MUA", type="primary"):
        with st.spinner("⏳ Hệ thống đang tính toán chia nhóm và cân bằng tải... Vui lòng chờ..."):
            
            # Đọc file từ upload
            df = pd.read_excel(uploaded_file)
            
            # ==========================================
            # LOGIC THUẬT TOÁN (BẢN FINAL V22)
            # ==========================================
            df.rename(columns={'Tổng mới': 'Tổng tuần cũ'}, inplace=True)

            days_mapping_new = {0: 'sun', 1: 'mon', 2: 'tue', 3: 'wed', 4: 'thu', 5: 'fri', 6: 'sat'}
            for d in days_mapping_new.values():
                if d not in df.columns:
                    df[d] = 0

            def check_case1(row):
                lich = str(row['Lịch về hàng']).strip()
                km = str(row['Đợt KM']).strip()
                days = [x.strip() for x in lich.split(',') if x.strip()]
                return len(days) == 7 and km in ['2 ngày KM/lân', '2 ngày KM/lần', '2 Lần/Tuần']

            def check_km1(row):
                return str(row['Đợt KM']).strip() == '1 Lần/Tuần'

            df['Is_Case1'] = df.apply(check_case1, axis=1)
            df['Is_KM1'] = df.apply(check_km1, axis=1)

            def calculate_tong_moi_final(row):
                lich = str(row['Lịch về hàng']).strip()
                tm = row['Tổng tuần cũ']
                qc = row['Quy cách mua']
                
                if pd.isna(qc): qc = 0
                if pd.isna(tm): tm = 0
                    
                days = [x.strip() for x in lich.split(',') if x.strip()]
                so_nhip = len(days)
                
                if row['Is_Case1']: return max(tm, qc * 4)
                elif 0 < so_nhip < 7: return max(tm, qc * so_nhip)
                return tm

            df['Tổng mới final'] = df.apply(calculate_tong_moi_final, axis=1)

            # 2.1 - Chia Siêu thị cho Case 1 (Tỷ lệ 31:24)
            case1_df = df[df['Is_Case1']].copy()
            st_totals_case1 = case1_df.groupby(['Kho', 'Mã siêu thị'])['Tổng mới final'].sum().reset_index()

            group_assignment_case1 = {}
            for kho, group_df in st_totals_case1.groupby('Kho'):
                st_list = group_df.sort_values('Tổng mới final', ascending=False).to_dict('records')
                bin1, bin2 = {'sts': [], 'sum': 0}, {'sts': [], 'sum': 0}
                
                for st_item in st_list:
                    total = bin1['sum'] + bin2['sum'] + st_item['Tổng mới final']
                    if total == 0:
                        bin1['sts'].append(st_item['Mã siêu thị']); bin1['sum'] += st_item['Tổng mới final']
                        continue
                        
                    if (total * (40/71) - bin1['sum']) >= (total * (31/71) - bin2['sum']):
                        bin1['sts'].append(st_item['Mã siêu thị']); bin1['sum'] += st_item['Tổng mới final']
                    else:
                        bin2['sts'].append(st_item['Mã siêu thị']); bin2['sum'] += st_item['Tổng mới final']
                            
                for st_id in bin1['sts']: group_assignment_case1[st_id] = '0,1,3,5'
                for st_id in bin2['sts']: group_assignment_case1[st_id] = '0,2,4,6'

            # 2.2 - Chia Siêu thị cho Nhóm 1 Lần/Tuần (Tỷ lệ 50:50)
            km1_df = df[df['Is_KM1']].copy()
            st_totals_km1 = km1_df.groupby(['Kho', 'Mã siêu thị'])['Tổng mới final'].sum().reset_index()

            group_assignment_km1 = {}
            for kho, group_df in st_totals_km1.groupby('Kho'):
                st_list = group_df.sort_values('Tổng mới final', ascending=False).to_dict('records')
                bin1, bin2 = {'sts': [], 'sum': 0}, {'sts': [], 'sum': 0}
                
                for st_item in st_list:
                    if bin1['sum'] <= bin2['sum']:
                        bin1['sts'].append(st_item['Mã siêu thị']); bin1['sum'] += st_item['Tổng mới final']
                    else:
                        bin2['sts'].append(st_item['Mã siêu thị']); bin2['sum'] += st_item['Tổng mới final']
                        
                for st_id in bin1['sts']: group_assignment_km1[st_id] = 'T3'
                for st_id in bin2['sts']: group_assignment_km1[st_id] = 'T5'

            def assign_lich_moi(row):
                if row['Is_Case1']: 
                    return group_assignment_case1.get(row['Mã siêu thị'], "Giữ nguyên lịch cũ")
                return "Giữ nguyên lịch cũ"
                
            def assign_lich_tinh_toan(row):
                lm = row['Lịch về hàng mới']
                if lm != "Giữ nguyên lịch cũ" and pd.notna(lm) and lm != "": return lm
                return row['Lịch về hàng']

            df['Lịch về hàng mới'] = df.apply(assign_lich_moi, axis=1)
            df['Lịch về hàng tính toán'] = df.apply(assign_lich_tinh_toan, axis=1)

            def assign_lich_km(row):
                km = str(row['Đợt KM']).strip()
                
                if km in ['2 ngày KM/lân', '2 ngày KM/lần', '2 Lần/Tuần']:
                    lich_ve = str(row['Lịch về hàng tính toán']).strip()
                    days = [int(x.strip()) for x in lich_ve.split(',') if x.strip().isdigit()]
                    
                    le_count = sum(1 for d in days if d in [1, 3, 5])
                    chan_count = sum(1 for d in days if d in [0, 2, 4, 6])
                    
                    if km in ['2 ngày KM/lân', '2 ngày KM/lần']:
                        return 'T3, T5, T7' if le_count > chan_count else 'T2, T4, T6'
                    elif km == '2 Lần/Tuần':
                        return 'T3, T7' if le_count > chan_count else 'T2, T6'
                        
                elif row['Is_KM1']:
                    return group_assignment_km1.get(row['Mã siêu thị'], "")
                    
                return ""

            df['Lịch KM ST'] = df.apply(assign_lich_km, axis=1)

            def distribute_by_quycach(row): 
                lich_ve = row['Lịch về hàng tính toán']
                try:
                    qck = float(row['Quy cách mua'])
                    tong_moi = float(row['Tổng mới final'])
                except: return row
                
                if pd.isnull(lich_ve) or qck == 0 or pd.isnull(tong_moi) or tong_moi == 0: return row
                try: 
                    original_days = sorted([int(x.strip()) for x in str(lich_ve).split(',') if x.strip() != ""])
                except: return row
                
                days = original_days.copy()
                len_days = len(days)
                if len_days == 0: return row
                
                for d in days_mapping_new.values(): row[d] = 0
                    
                so_quy_cach_rai = int(tong_moi // qck)
                phan_le = tong_moi - so_quy_cach_rai * qck
                
                is_km1 = row['Is_KM1']
                is_new_schedule = (row['Lịch về hàng mới'] != "Giữ nguyên lịch cũ")

                # ----------------------------------------
                # LUỒNG 1: NHÓM KHUYẾN MÃI 1 LẦN/TUẦN
                # ----------------------------------------
                if is_km1:
                    if len(original_days) == 7:
                        km_st = str(row['Lịch KM ST']).strip()
                        day_after_km = 3 if km_st == 'T3' else (5 if km_st == 'T5' else -1)
                        
                        vong = so_quy_cach_rai // 7
                        so_phan_con_lai = so_quy_cach_rai % 7
                        
                        for _ in range(vong):
                            for d in days:
                                row[days_mapping_new[d]] += qck
                                
                        if so_phan_con_lai > 0:
                            buoc_nhay = 1 if so_phan_con_lai == 6 else (2 if so_phan_con_lai in [5,4,3] else (4 if so_phan_con_lai == 2 else 1))
                            
                            weights = [0.1 if (d == 3 or d == 5) else 1.0 for d in days]
                            start_day = random.choices(days, weights=weights, k=1)[0]
                            start_idx = days.index(start_day)
                            
                            chosen_days = [days[(start_idx + i * buoc_nhay) % 7] for i in range(so_phan_con_lai)]
                            
                            if day_after_km != -1 and day_after_km not in chosen_days:
                                chosen_days[-1] = day_after_km
                                
                            for d in chosen_days:
                                row[days_mapping_new[d]] += qck
                                
                        if phan_le > 0:
                            if so_quy_cach_rai == 0 and day_after_km != -1:
                                row[days_mapping_new[day_after_km]] += phan_le
                            else:
                                weights_le = [0.1 if (d == 3 or d == 5) else 1.0 for d in days]
                                day_le = random.choices(days, weights=weights_le, k=1)[0]
                                row[days_mapping_new[day_le]] += phan_le

                    else:
                        km_st = str(row['Lịch KM ST']).strip()
                        target_day = 3 if km_st == 'T3' else (5 if km_st == 'T5' else -1)
                        suppress_day = 5 if km_st == 'T3' else (3 if km_st == 'T5' else -1)
                        
                        if len(days) >= 6 and suppress_day in days:
                            if random.random() >= 0.10: 
                                days.remove(suppress_day)
                                
                        len_days_current = len(days)
                        if len_days_current == 0: return row
                        
                        vong = so_quy_cach_rai // len_days_current
                        n_phan = so_quy_cach_rai % len_days_current
                        
                        has_target = target_day in days
                        is_piggybank = False
                        
                        if vong == 0 and has_target:
                            if so_quy_cach_rai >= 1:
                                row[days_mapping_new[target_day]] += qck
                                so_quy_cach_rai -= 1
                                n_phan = so_quy_cach_rai
                                is_piggybank = True
                            elif phan_le > 0:
                                row[days_mapping_new[target_day]] += phan_le
                                phan_le = 0
                                
                        if vong > 0:
                            for d in days:
                                row[days_mapping_new[d]] += vong * qck
                                
                        extra_eligible_days = days.copy()
                        if has_target and (vong > 0 or is_piggybank):
                            extra_eligible_days.remove(target_day)
                            
                        if n_phan > 0:
                            if not extra_eligible_days: extra_eligible_days = days.copy()
                            chosen_days = random.sample(extra_eligible_days, n_phan)
                            for d in chosen_days:
                                row[days_mapping_new[d]] += qck
                                
                        if phan_le > 0:
                            if not extra_eligible_days: extra_eligible_days = days.copy()
                            row[days_mapping_new[random.choice(extra_eligible_days)]] += phan_le

                # ----------------------------------------
                # LUỒNG 2: CÁC NHÓM CÒN LẠI (V22 Tinh chỉnh Random phần thừa)
                # ----------------------------------------
                else:
                    # Xử lý đặc biệt cho lịch mới 0,2,4,6 (T3, T5, T7, CN)
                    if str(lich_ve).strip() == '0,2,4,6' and is_new_schedule:
                        keep_sunday = random.random() >= 0.875
                        
                        if keep_sunday:
                            actual_days = [0, 2, 4, 6] # Rải 4 ngày: CN, T3, T5, T7
                        else:
                            actual_days = [2, 4, 6]    # Gom 3 ngày: T3, T5, T7
                            
                        len_actual = len(actual_days)
                        vong = int(so_quy_cach_rai // len_actual)
                        n_phan_le = int(so_quy_cach_rai % len_actual)
                        
                        # 1. Rải đều phần nguyên (đảm bảo số chia không lệch)
                        for d in actual_days:
                            row[days_mapping_new[d]] += vong * qck
                            
                        # 2. Rải phần lẻ NGẪU NHIÊN để tránh dồn số lượng vào T3
                        if n_phan_le > 0:
                            chosen_le_days = random.sample(actual_days, n_phan_le)
                            for d in chosen_le_days:
                                row[days_mapping_new[d]] += qck
                            
                        # 3. Xử lý phần dư lẻ (nhỏ hơn 1 QCM)
                        if phan_le > 0:
                            row[days_mapping_new[random.choice(actual_days)]] += phan_le

                    # Xử lý cho các lịch khác trong luồng 2 (giữ nguyên logic cũ)
                    else:
                        def add_day_idx_v7(idx):
                            day_index = days[idx % len_days]
                            row[days_mapping_new[day_index]] += qck

                        def rai_buoc_nhay_v7(n, buoc, start):
                            for i in range(n): add_day_idx_v7((start + i*buoc) % len_days)
                            
                        if len_days == 7:
                            vong = int(so_quy_cach_rai // 7)
                            n_phan = int(so_quy_cach_rai % 7)
                            for _ in range(vong):
                                for i in range(7): add_day_idx_v7(i)
                                
                            if n_phan > 0:
                                buoc_nhay = 1 if n_phan == 6 else (2 if n_phan in [5,4,3] else (4 if n_phan == 2 else 1))
                                day_start = random.randint(0, len_days - 1)
                                rai_buoc_nhay_v7(n_phan, buoc_nhay, day_start)
                            else:
                                day_start = random.randint(0, len_days - 1)
                                
                            if phan_le > 0:
                                day_cho = days[day_start] 
                                row[days_mapping_new[day_cho]] += phan_le
                        else: 
                            start_pos = random.randint(0, len_days - 1)
                            for i in range(int(so_quy_cach_rai)):
                                add_day_idx_v7((start_pos + i) % len_days)
                                
                            if phan_le > 0:
                                row[days_mapping_new[random.choice(days)]] += phan_le
                                
                return row
            
            df = df.apply(distribute_by_quycach, axis=1)

            final_columns = [
                'Mã siêu thị', 'Tên siêu thị', 'Mã sản phẩm', 'Tên sản phẩm', 'Kho', 
                'Lịch về hàng', 'Quy cách mua', 'Đợt KM', 'Lịch KM ST',
                'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN', 
                'Tổng tuần cũ', 'Lịch về hàng mới', 'Lịch về hàng tính toán', 'Tổng mới final', 
                'mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'
            ]

            df_output = df[[c for c in final_columns if c in df.columns]]
            
            # Xuất file ra bộ nhớ đệm (RAM) để tải về
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                df_output.to_excel(writer, index=False, sheet_name='DuLieuRai')
            output_data = output_buffer.getvalue()

        # Hiển thị kết quả
        st.markdown("### 📊 3. Kết quả phân bổ")
        st.info(f"✅ Quá trình rải hàng đã hoàn tất! Xử lý thành công **{len(df_output)}** dòng dữ liệu.")
        st.dataframe(df_output.head(100), use_container_width=True)
        
        st.markdown(
            """
            <style>
            .download-btn {
                display: flex;
                justify-content: center;
                margin-top: 20px;
            }
            </style>
            """, unsafe_allow_html=True
        )
        
        st.download_button(
            label="⬇️ TẢI FILE EXCEL KẾT QUẢ",
            data=output_data,
            file_name="Ket_Qua_Rai_Hang_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
else:
    st.info("Chưa có dữ liệu. Vui lòng tải file Excel lên hệ thống ở mục 1.")
