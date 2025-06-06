import streamlit as st
import pandas as pd
from jinja2 import Template
import base64
import re
import io
import unicodedata
import openpyxl

st.set_page_config(page_title="Báo cáo kết quả đào tạo - VIAGS", layout="wide")

st.title("📋 Quản lý lớp học - VIAGS")

# Hàm chuẩn hóa thời gian
def chuan_hoa_thoi_gian(time_str):
    # 26-27/5/2025 -> 26,27/5/2025
    match = re.match(r"(\d{1,2})-(\d{1,2})/(\d{1,2}/\d{4})", str(time_str))
    if match:
        ngay1, ngay2, thangnam = match.groups()
        return f"{ngay1},{ngay2}/{thangnam}"
    return str(time_str).strip()

# Hàm loại bỏ dấu tiếng Việt và chuẩn hóa chuỗi
def remove_vietnamese_accents(s):
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize('NFD', s)
    s = s.encode('ascii', 'ignore').decode('utf-8')
    s = s.replace(' ', '').lower()
    return s

def normalize_name(s):
    s = str(s) if s is not None else ""
    s = s.split('-')[0].strip()
    # Bổ sung thay thế Đ/đ thành D/d
    s = s.replace('Đ', 'D').replace('đ', 'd')
    import unicodedata, re
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    s = re.sub(r'\s+', '', s)
    s = s.lower()
    return s


def strip_accents(s):
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize('NFD', s)
    s = s.replace('Đ', 'D').replace('đ', 'd')
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')


# ========== Quản lý nhiều lớp ==========
if "danh_sach_lop" not in st.session_state:
    st.session_state["danh_sach_lop"] = {}
if "ten_lop_hien_tai" not in st.session_state:
    st.session_state["ten_lop_hien_tai"] = ""
if "hien_nhap_excel" not in st.session_state:
    st.session_state["hien_nhap_excel"] = False

# Sắp xếp danh sách lớp theo thứ tự tiếng Việt

ds_lop = sorted(list(st.session_state["danh_sach_lop"].keys()), key=strip_accents)

chuc_nang = st.columns([5, 2, 2, 3])
with chuc_nang[0]:
    ten_lop = st.selectbox(
        "🗂️ Chọn lớp",
        ds_lop + ["+ Tạo lớp mới"],
        index=ds_lop.index(st.session_state["ten_lop_hien_tai"]) if st.session_state["ten_lop_hien_tai"] in ds_lop else len(ds_lop),
    )
with chuc_nang[1]:
    ten_moi = st.text_input("Tên lớp mới", value="", placeholder="VD: ATHK 01/2025")
    tao_lop = st.button("➕ Tạo lớp mới")
with chuc_nang[2]:
    if ds_lop and st.button("🗑️ Xóa lớp đang chọn"):
        if st.session_state["ten_lop_hien_tai"] in st.session_state["danh_sach_lop"]:
            del st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]
            st.session_state["ten_lop_hien_tai"] = ds_lop[0] if ds_lop else ""
with chuc_nang[3]:
    if st.button("📥 Nhập nhiều lớp từ Excel", key="open_excel_modal"):
        st.session_state["hien_nhap_excel"] = True


# Hiển thị khối nhập file Excel khi bấm nút (giả popup)
if st.session_state.get("hien_nhap_excel", False):
    with st.expander("📥 Nhập nhiều lớp từ file Excel (mỗi sheet 1 lớp)", expanded=True):
        file_excel = st.file_uploader(
            "Chọn file Excel danh sách lớp",
            type=["xlsx"],
            key="multi_class_uploader_import"
        )
        col_excel = st.columns([2, 1])
        with col_excel[0]:
            nhap_excel = st.button("Nhập các lớp vào hệ thống", key="btn_nhap_excel")
        with col_excel[1]:
            huy_excel = st.button("❌ Đóng nhập nhiều lớp", key="btn_huy_excel")
        # Xử lý nhập và đóng form
        if huy_excel:
            st.session_state["hien_nhap_excel"] = False
            st.rerun()

    if file_excel is not None and nhap_excel:
        wb = openpyxl.load_workbook(file_excel, data_only=True)
        so_lop_them = 0
        lop_moi_vua_them = None
        log_sheets = []

        for sheetname in wb.sheetnames:
            sheet_check = remove_vietnamese_accents(sheetname)
            if sheet_check == "mucluc":
                log_sheets.append(f"⏩ Bỏ qua sheet '{sheetname}' (Mục lục).")
                continue

            ws = wb[sheetname]
            ten_lop_goc = ws["D7"].value
            if not ten_lop_goc or str(ten_lop_goc).strip() == "":
                log_sheets.append(f"❌ Sheet '{sheetname}': Thiếu tên lớp ở D7.")
                continue

            thoi_gian = ws["D9"].value or ""
            thoi_gian_chuan = chuan_hoa_thoi_gian(thoi_gian)
            # Tạo tên lớp như code bố đang dùng
            ten_lop = f"{str(ten_lop_goc).strip()}_{str(thoi_gian).strip()}"
            orig_ten_lop = ten_lop
            cnt = 1
            while ten_lop in st.session_state["danh_sach_lop"]:
                ten_lop = f"{orig_ten_lop}_{cnt}"
                cnt += 1

            # Loại hình/hình thức đào tạo
            loai_hinh_full = ws["B8"].value or ""
            if ":" in str(loai_hinh_full):
                loai_hinh = str(loai_hinh_full).split(":", 1)[-1].strip()
            else:
                loai_hinh = str(loai_hinh_full).strip()
            dia_diem = ws["D10"].value or ""

            # Đọc danh sách học viên từ dòng 14 trở đi (C14 - Mã NV, D14 - Họ tên, E14 - Đơn vị)
            data = []
            row = 14
            while True:
                ma_nv = ws[f"C{row}"].value
                ho_ten = ws[f"D{row}"].value
                don_vi = ws[f"E{row}"].value
                # Nếu cả 3 ô đều trống thì dừng
                if (not ma_nv or str(ma_nv).strip() == "") and (not ho_ten or str(ho_ten).strip() == ""):
                    break
                # Nếu 1 trong các ô chứa từ khóa "Trưởng", "Trung tâm", "Ký tên" thì dừng
                if any((
                    (isinstance(ma_nv, str) and ("trưởng" in ma_nv.lower() or "trung tâm" in ma_nv.lower() or "ký tên" in ma_nv.lower())),
                    (isinstance(ho_ten, str) and ("trưởng" in ho_ten.lower() or "trung tâm" in ho_ten.lower() or "ký tên" in ho_ten.lower())),
                    (isinstance(don_vi, str) and ("trưởng" in don_vi.lower() or "trung tâm" in don_vi.lower() or "ký tên" in don_vi.lower()))
                )):
                    break
                if (ma_nv and str(ma_nv).strip() != "") or (ho_ten and str(ho_ten).strip() != ""):
                    data.append({
                        "Mã NV": str(ma_nv or "").strip(),
                        "Họ tên": str(ho_ten or "").strip(),
                        "Đơn vị": str(don_vi or "").strip(),
                        "Điểm": ""
                    })
                row += 1


            if len(data) > 0:
                df = pd.DataFrame(data)
                st.session_state["danh_sach_lop"][ten_lop] = {
                    "class_info": {
                        "course_name": ten_lop_goc,
                        "training_type": loai_hinh,
                        "time": thoi_gian_chuan,
                        "location": dia_diem,
                        "num_attended": "",
                        "num_total": "",
                    },
                    "ds_hocvien": df
                }
                lop_moi_vua_them = ten_lop
                so_lop_them += 1
                log_sheets.append(f"✅ Sheet '{sheetname}' ({ten_lop_goc}) đã nhập {len(data)} học viên (tên lớp: {ten_lop})")
            else:
                log_sheets.append(f"❌ Sheet '{sheetname}': Không có học viên ở C14-E14 trở đi.")

        if so_lop_them:
            st.session_state["ten_lop_hien_tai"] = lop_moi_vua_them
                # RESET biến tạm về đúng lớp vừa nhập cuối cùng
            cur_lop_data = st.session_state["danh_sach_lop"][lop_moi_vua_them]
            st.session_state["class_info_tmp"] = cur_lop_data.get("class_info", {}).copy()
            st.session_state["ds_hocvien_tmp"] = cur_lop_data.get("ds_hocvien", pd.DataFrame()).copy()
            st.session_state["diem_tmp"] = cur_lop_data.get("ds_hocvien", pd.DataFrame()).copy()
            st.success(f"Đã nhập xong {so_lop_them} lớp! Vào phần 'Chọn lớp' để kiểm tra.")
            for log in log_sheets:
                st.write(log)
            st.session_state["hien_nhap_excel"] = False
            st.rerun()
        else:
            for log in log_sheets:
                st.write(log)
            st.warning("Không tìm thấy sheet nào hợp lệ (phải có D7 là tên lớp và học viên từ C14-E14).")

# Xử lý tạo lớp mới
if tao_lop and ten_moi.strip():
    if ten_moi not in st.session_state["danh_sach_lop"]:
        st.session_state["danh_sach_lop"][ten_moi] = {
            "class_info": {
                "course_name": "",
                "training_type": "",
                "time": "",
                "location": "", 
                "num_attended": "",
                "num_total": "",
            },
            "ds_hocvien": pd.DataFrame({
                "Mã NV": [""] * 30,
                "Họ tên": [""] * 30,
                "Đơn vị": [""] * 30,
                "Điểm": [""] * 30
            }),
        }
        st.session_state["ten_lop_hien_tai"] = ten_moi

        # RESET các biến tạm cho lớp mới tạo
        cur_lop_data = st.session_state["danh_sach_lop"][ten_moi]
        st.session_state["class_info_tmp"] = cur_lop_data.get("class_info", {}).copy()
        st.session_state["ds_hocvien_tmp"] = cur_lop_data.get("ds_hocvien", pd.DataFrame()).copy()
        st.session_state["diem_tmp"] = cur_lop_data.get("ds_hocvien", pd.DataFrame()).copy()

        st.rerun()
    else:
        st.warning("Tên lớp đã tồn tại!")
elif ten_lop and ten_lop != "+ Tạo lớp mới":
    st.session_state["ten_lop_hien_tai"] = ten_lop
    # RESET các biến tạm theo lớp mới chọn
    cur_lop_data = st.session_state["danh_sach_lop"][ten_lop]
    st.session_state["class_info_tmp"] = cur_lop_data.get("class_info", {}).copy()
    st.session_state["ds_hocvien_tmp"] = cur_lop_data.get("ds_hocvien", pd.DataFrame()).copy()
    st.session_state["diem_tmp"] = cur_lop_data.get("ds_hocvien", pd.DataFrame()).copy()

# Nếu chưa có lớp nào, yêu cầu tạo trước
if not st.session_state["ten_lop_hien_tai"]:
    st.info("🔔 Hãy tạo lớp mới để bắt đầu nhập liệu và quản lý!")
    st.stop()

# Lấy dữ liệu lớp hiện tại
lop_data = st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]
class_info = lop_data.get("class_info", {})
ds_hocvien = lop_data.get("ds_hocvien", pd.DataFrame({
    "Mã NV": [""] * 30,
    "Họ tên": [""] * 30,
    "Đơn vị": [""] * 30,
    "Điểm": [""] * 30
}))

# ==== Chuẩn bị biến tạm cho cơ chế lưu khi chuyển tab ====
if "active_tab" not in st.session_state:
    st.session_state["active_tab"] = "1️⃣ Thông tin lớp học"

if "ds_hocvien_tmp" not in st.session_state:
    st.session_state["ds_hocvien_tmp"] = ds_hocvien.copy()
if "diem_tmp" not in st.session_state:
    st.session_state["diem_tmp"] = ds_hocvien.copy()

def save_data_when_switch_tab(new_tab):
    
    # Tab 1: Lưu thông tin lớp học khi chuyển tab
    if st.session_state["active_tab"] == "1️⃣ Thông tin lớp học" and new_tab != "1️⃣ Thông tin lớp học":
        st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]["class_info"] = st.session_state["class_info_tmp"].copy()# Tab 2: Lưu danh sách học viên (không điểm)
    if st.session_state["active_tab"] == "2️⃣ Danh sách học viên" and new_tab != "2️⃣ Danh sách học viên":
        ds = st.session_state["ds_hocvien_tmp"].copy()
        # Nếu đang có cột điểm thì reset, nếu không thì thôi (hoặc giữ lại, tùy bố muốn)
        for col in ["Điểm LT", "Điểm TH"]:
            if col in ds.columns:
                ds = ds.drop(columns=[col])
        # Reset lại điểm khi danh sách thay đổi
        ds["Điểm LT"] = ""
        ds["Điểm TH"] = ""
        st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]["ds_hocvien"] = ds.copy()
    # Tab 3: Lưu điểm
    if st.session_state["active_tab"] == "3️⃣ Cập nhật điểm" and new_tab != "3️⃣ Cập nhật điểm":
        st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]["ds_hocvien"] = st.session_state["diem_tmp"].copy()
    st.session_state["active_tab"] = new_tab

# ========= Tabs =========
tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Thông tin lớp học", 
    "2️⃣ Danh sách học viên",
    "3️⃣ Cập nhật điểm",
    "4️⃣ Chữ ký & xuất báo cáo"
])
# ========== Tab nội dung ==========

    
with tab1:
    save_data_when_switch_tab("1️⃣ Thông tin lớp học")
    st.subheader("Nhập thông tin lớp học")
    class_info_sample = '''An toàn hàng không
Định kỳ/Elearning+Trực tiếp
02/01/2025
TTĐT MB
VNBA25-ĐKVH04'''
    # Lấy dữ liệu từ biến tạm, nếu chưa có thì copy từ dữ liệu gốc
    if st.session_state["active_tab"] != "1️⃣ Thông tin lớp học":
        st.session_state["class_info_tmp"] = st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]["class_info"].copy()

    class_info_tmp = st.session_state["class_info_tmp"]
    class_info_input = st.text_area(
        "Dán vào 5 dòng gồm: Môn học, Loại hình, Thời gian, Địa điểm, [Mã lớp/ghi chú nếu có]", 
        value="\n".join([
            class_info_tmp.get("course_name", ""),
            class_info_tmp.get("training_type", ""),
            class_info_tmp.get("time", ""),
            class_info_tmp.get("location", ""),
            class_info_tmp.get("class_code", "")
        ]) if any(class_info_tmp.values()) else class_info_sample, height=130
    )

    class_info_lines = class_info_input.strip().split("\n")
    course_name = class_info_lines[0] if len(class_info_lines) > 0 else ""
    training_type = class_info_lines[1] if len(class_info_lines) > 1 else ""
    time = class_info_lines[2] if len(class_info_lines) > 2 else ""
    location = class_info_lines[3] if len(class_info_lines) > 3 else ""
    class_code_note = class_info_lines[4].strip() if len(class_info_lines) > 4 else ""

    st.session_state["class_info_tmp"] = {
        "course_name": course_name,
        "training_type": training_type,
        "time": time,
        "location": location,
        "class_code": class_code_note,
    }
    st.info("Thông tin sẽ được lưu khi chuyển sang tab khác.")


with tab2:
    save_data_when_switch_tab("2️⃣ Danh sách học viên")
    st.subheader("Danh sách học viên")
    st.caption("📌 Dán hoặc nhập danh sách học viên, chỉ chỉnh sửa thông tin cá nhân ở đây (KHÔNG nhập điểm ở tab này).")

    # Khởi tạo lại biến tạm nếu vừa chuyển sang tab hoặc danh sách học viên tạm bị rỗng
    if st.session_state["active_tab"] != "2️⃣ Danh sách học viên" or st.session_state["ds_hocvien_tmp"].empty:
        ds_hocvien_tmp = st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]["ds_hocvien"].copy()
        # Loại bỏ các cột điểm nếu có (đảm bảo tab 2 chỉ quản lý thông tin cá nhân)
        for col in ["Điểm LT", "Điểm TH"]:
            if col in ds_hocvien_tmp.columns:
                ds_hocvien_tmp = ds_hocvien_tmp.drop(columns=[col])
        st.session_state["ds_hocvien_tmp"] = ds_hocvien_tmp.copy()

    ds_hocvien_tmp = st.session_state["ds_hocvien_tmp"]

    # Đảm bảo đủ 3 cột
    for col in ["Mã NV", "Họ tên", "Đơn vị"]:
        if col not in ds_hocvien_tmp.columns:
            ds_hocvien_tmp[col] = ""

    ds_hocvien_tmp = ds_hocvien_tmp[["Mã NV", "Họ tên", "Đơn vị"]]

    ds_hocvien_tmp_new = st.data_editor(
        ds_hocvien_tmp,
        num_rows="dynamic",
        hide_index=False,
        use_container_width=True,
        column_order=["Mã NV", "Họ tên", "Đơn vị"],
        column_config={
            "Mã NV": st.column_config.TextColumn(width="x-small"),
            "Họ tên": st.column_config.TextColumn(width="large"),
            "Đơn vị": st.column_config.TextColumn(width="medium"),
        },
        key="data_editor_ds"
    )

    # Luôn lưu vào biến tạm, KHÔNG ghi session_state chính cho đến khi chuyển tab!
    st.session_state["ds_hocvien_tmp"] = ds_hocvien_tmp_new.copy()

    st.info("Mọi thay đổi sẽ được lưu khi chuyển sang tab khác.")


with tab3:
    save_data_when_switch_tab("3️⃣ Cập nhật điểm")
    st.subheader("Nhập điểm (từ file hoặc nhập tay)")
    # LUÔN lấy data mới nhất
    ds_hocvien = st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]["ds_hocvien"].copy()
    st.session_state["diem_tmp"] = ds_hocvien.copy()
    ds_hocvien = st.session_state["diem_tmp"]

    if ds_hocvien.empty or "Họ tên" not in ds_hocvien.columns:
        st.error("❌ Chưa có danh sách học viên. Vui lòng nhập ở tab 2 trước.")
        st.stop()

    # Upload file điểm tự động GHÉP
    st.markdown("**Tải file điểm dạng LMS hoặc Đợt thi để GHÉP tự động vào cột Điểm LT:**")
    uploaded_lms = st.file_uploader("📥 File điểm LMS", type=["xlsx"], key="uploader_lms_tab3")
    uploaded_dotthi = st.file_uploader("📥 File điểm Đợt thi", type=["xlsx"], key="uploader_dotthi_tab3")

    if uploaded_lms is not None:
        df_diem = pd.read_excel(uploaded_lms)
        col_name_hoten = df_diem.columns[3]
        col_name_lanthi = df_diem.columns[6]
        df_diem["HoTenChuan"] = df_diem[col_name_hoten].apply(normalize_name)
        def extract_diem_lanthi(text):
            if not isinstance(text, str):
                return ""
            scores = re.findall(r"Lần \d+\s*:\s*(\d+)", text)
            return "/".join(scores)
        df_diem["DiemDaXuLy"] = df_diem[col_name_lanthi].apply(extract_diem_lanthi)
        diem_map = dict(zip(df_diem["HoTenChuan"], df_diem["DiemDaXuLy"]))
        matched = 0
        ds_hocvien["HoTenChuan"] = ds_hocvien["Họ tên"].apply(normalize_name)
        for i, row in ds_hocvien.iterrows():
            key = row["HoTenChuan"]
            if key in diem_map and diem_map[key]:
                ds_hocvien.at[i, "Điểm LT"] = diem_map[key]
                matched += 1
        ds_hocvien = ds_hocvien.drop(columns=["HoTenChuan"])
        if matched > 0:
            st.success(f"✅ Đã ghép điểm LT cho {matched} học viên.")
        else:
            st.warning("⚠️ Không ghép được điểm. Hãy kiểm tra lại tên học viên.")
        st.session_state["diem_tmp"] = ds_hocvien.copy()

    if uploaded_dotthi is not None:
        df_dotthi = pd.read_excel(uploaded_dotthi)
        col_name_hoten = df_dotthi.columns[2]
        col_name_diem_1lan = df_dotthi.columns[4]
        col_name_diem_nlan = df_dotthi.columns[6]
        def extract_score_dotthi(row):
            diem_1lan = row[col_name_diem_1lan]
            diem_nlan = row[col_name_diem_nlan]
            if pd.notnull(diem_nlan) and str(diem_nlan).strip() != "":
                scores = re.findall(r"Lần\s*\d+\s*:\s*(\d+)", str(diem_nlan))
                return "/".join(scores) if scores else str(diem_nlan).strip()
            elif pd.notnull(diem_1lan) and str(diem_1lan).strip() != "":
                return str(diem_1lan).strip()
            return ""
        df_dotthi["HoTenChuan"] = df_dotthi[col_name_hoten].apply(normalize_name)
        df_dotthi["DiemDaXuLy"] = df_dotthi.apply(extract_score_dotthi, axis=1)
        diem_map = dict(zip(df_dotthi["HoTenChuan"], df_dotthi["DiemDaXuLy"]))
        matched = 0
        ds_hocvien["HoTenChuan"] = ds_hocvien["Họ tên"].apply(normalize_name)
        for i, row in ds_hocvien.iterrows():
            key = row["HoTenChuan"]
            if key in diem_map and diem_map[key]:
                ds_hocvien.at[i, "Điểm LT"] = diem_map[key]
                matched += 1
        ds_hocvien = ds_hocvien.drop(columns=["HoTenChuan"])
        if matched > 0:
            st.success(f"✅ Đã ghép điểm LT cho {matched} học viên.")
        else:
            st.warning("⚠️ Không ghép được điểm.")
        st.session_state["diem_tmp"] = ds_hocvien.copy()

    # ĐẢM BẢO đủ 2 cột "Điểm LT", "Điểm TH" trước khi hiển thị data_editor
    for col in ["Điểm LT", "Điểm TH"]:
        if col not in ds_hocvien.columns:
            ds_hocvien[col] = ""

    # Hiển thị và cho phép NHẬP/SỬA trực tiếp điểm LT, TH (KHÔNG cho sửa danh tính)
    st.markdown("**Hoặc nhập điểm LT, điểm TH trực tiếp:**")
    cols_show = ["Mã NV", "Họ tên", "Đơn vị", "Điểm LT", "Điểm TH"]
    ds_hocvien_edit = st.data_editor(
        ds_hocvien[cols_show],
        num_rows="fixed",
        hide_index=False,
        use_container_width=True,
        column_order=cols_show,
        disabled=["Mã NV", "Họ tên", "Đơn vị"],
        key="diem_editor_tab3"
    )

    # Cập nhật điểm vào biến tạm
    for col in ["Điểm LT", "Điểm TH"]:
        ds_hocvien[col] = ds_hocvien_edit[col]
    st.session_state["diem_tmp"] = ds_hocvien.copy()

    st.info("Mọi thay đổi điểm sẽ được lưu khi chuyển sang tab khác.")

with tab4:
    save_data_when_switch_tab("4️⃣ Chữ ký & xuất báo cáo")
    st.subheader("Thông tin chữ ký báo cáo & Xuất báo cáo")

    # Lấy danh sách học viên từ đúng lớp đang chọn
    ds_hocvien = st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]["ds_hocvien"].copy()

    gv_huong_dan = st.text_input("Họ tên Giáo viên hướng dẫn", value="Nguyễn Đức Nghĩa")
    truong_bo_mon = st.text_input("Họ tên Trưởng bộ môn", value="Ngô Trung Thành")
    truong_tt = st.text_input("Họ tên Trưởng TTĐT", value="Nguyễn Chí Kiên")

    # Lấy thông tin lớp
    lop_data = st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]]
    class_info = lop_data["class_info"]
    course_name = class_info.get("course_name", "")
    training_type = class_info.get("training_type", "")
    time = class_info.get("time", "")
    location = class_info.get("location", "")
    num_attended = class_info.get("num_attended", "")
    num_total = class_info.get("num_total", "")

    def extract_days(time_str):
        if not time_str:
            return []
        time_str = time_str.replace('S', '').replace('s', '')
        match = re.search(r'([\d,]+)/(\d{1,2})/(\d{4})', time_str)
        if match:
            days = [d.strip() for d in match.group(1).split(',')]
            month = match.group(2)
            return [f"{d}/{month}" for d in days if d]
        match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', time_str)
        if match:
            return [f"{match.group(1)}/{match.group(2)}"]
        return []

    def ma_hoa_ten_lop(course_name, time_str):
        #Mã hóa tên lớp thành tên ngắn gọn (ví dụ: YTCN_040625)
        from datetime import datetime
        words = re.findall(r'\w+', str(course_name))
        initials = ''.join([w[0].upper() for w in words])[:10]
        s = str(time_str)
        match = re.match(r'(\d{1,2})[, -](\d{1,2})/(\d{1,2})/(\d{4})', s)
        if match:
            dd1 = match.group(1).zfill(2)
            dd2 = match.group(2).zfill(2)
            mm = match.group(3).zfill(2)
            yy = match.group(4)[-2:]
            time_part = f"{dd1}{dd2}{mm}{yy}"
        else:
            match2 = re.match(r'(\d{1,2})[, -](\d{1,2})/(\d{1,2})$', s)
            if match2:
                dd1 = match2.group(1).zfill(2)
                dd2 = match2.group(2).zfill(2)
                mm = match2.group(3).zfill(2)
                yy = str(datetime.now().year)[-2:]
                time_part = f"{dd1}{dd2}{mm}{yy}"
            else:
                match3 = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
                if match3:
                    dd = match3.group(1).zfill(2)
                    mm = match3.group(2).zfill(2)
                    yy = match3.group(3)[-2:]
                    time_part = f"{dd}{mm}{yy}"
                else:
                    nums = re.findall(r'\d+', s)
                    if len(nums) >= 4:
                        dd1 = nums[0].zfill(2)
                        dd2 = nums[2].zfill(2)
                        mm = nums[-2].zfill(2)
                        yy = nums[-1][-2:]
                        time_part = f"{dd1}{dd2}{mm}{yy}"
                    elif len(nums) == 3:
                        dd = nums[0].zfill(2)
                        mm = nums[1].zfill(2)
                        yy = nums[2][-2:]
                        time_part = f"{dd}{mm}{yy}"
                    else:
                        time_part = "000000"
        return f"{initials}_{time_part}"

    with open("logo_viags.png", "rb") as image_file:
        logo_base64 = base64.b64encode(image_file.read()).decode()

    col1, col2, _ = st.columns([1, 1, 4])
    with col1:
        bckq = st.button("📄In báo cáo kết quả")
    with col2:
        diem_danh = st.button("Tạo bảng điểm danh")

    if bckq:
        if ds_hocvien.empty:
            st.warning("Vui lòng nhập danh sách học viên!")
        else:
            ds_hocvien_filtered = ds_hocvien[
                (ds_hocvien["Mã NV"].astype(str).str.strip() != "") | (ds_hocvien["Họ tên"].astype(str).str.strip() != "")]
            data = []

            diem_lt_nonempty = ds_hocvien_filtered["Điểm LT"].astype(str).str.strip().replace("-", "").replace("nan", "").replace("None", "").ne("").sum()
            diem_th_nonempty = ds_hocvien_filtered["Điểm TH"].astype(str).str.strip().replace("-", "").replace("nan", "").replace("None", "").ne("").sum()
            use_5b = diem_lt_nonempty > 0 and diem_th_nonempty > 0
            template_file = "report_template_5b.html" if use_5b else "report_template_5a.html"

            for i, row in ds_hocvien_filtered.iterrows():
                ma_nv = str(row.get("Mã NV", "") or "").strip()
                ho_ten = str(row.get("Họ tên", "") or "").strip()
                if (not ma_nv or ma_nv.lower() == "none") and (not ho_ten or ho_ten.lower() == "none"):
                    continue

                diem_lt = str(row.get("Điểm LT", "") or "").strip()
                diem_th = str(row.get("Điểm TH", "") or "").strip()

                if use_5b:
                    diem_lt = diem_lt if diem_lt not in ["", "nan", "None", None] else "-"
                    diem_th = diem_th if diem_th not in ["", "nan", "None", None] else "-"

                    def get_last_score(s):
                        if s in ["", "-", "nan", "None", None]:
                            return 0
                        parts = [p.strip() for p in str(s).replace(",", ".").split("/") if p.strip().replace(".", "", 1).isdigit()]
                        return float(parts[-1]) if parts else 0

                    d_lt = get_last_score(diem_lt)
                    d_th = get_last_score(diem_th)
                    try:
                        diem_tb = round((d_lt + 2 * d_th) / 3)
                    except:
                        diem_tb = 0

                    scores_lt = [s for s in str(diem_lt).split("/") if s.strip().isdigit()]
                    scores_th = [s for s in str(diem_th).split("/") if s.strip().isdigit()]
                    num_tests = max(len(scores_lt), len(scores_th))
                    # Xếp loại: >1 lần chỉ "Đạt" nếu >=80, không có "Xuất sắc"
                    if diem_lt == "-" and diem_th == "-":
                        xep_loai = "-"
                    elif num_tests > 1:
                        xep_loai = "Đạt" if diem_tb >= 80 else "Không đạt"
                    else:
                        if diem_tb >= 95:
                            xep_loai = "Xuất sắc"
                        elif diem_tb >= 80:
                            xep_loai = "Đạt"
                        else:
                            xep_loai = "Không đạt"
                    # Ghi chú
                    if diem_lt == "-" and diem_th == "-":
                        note = "Vắng"
                    else:
                        note = ""
                        main_scores = scores_lt if len(scores_lt) > 1 else scores_th
                        if len(main_scores) > 1:
                            note = f"Kiểm tra lần {'/'.join(str(i+1) for i in range(len(main_scores)))}"
                    data.append({
                        "id": ma_nv,
                        "name": ho_ten,
                        "unit": str(row.get("Đơn vị", "") or "").strip(),
                        "score_lt": diem_lt,
                        "score_th": diem_th,
                        "score_tb": diem_tb,
                        "rank": xep_loai,
                        "note": note,
                        "num_tests": num_tests
                    })
                else:
                    if diem_lt and diem_lt not in ["", "-", "nan", "None", None]:
                        diem_chinh = diem_lt
                    elif diem_th and diem_th not in ["", "-", "nan", "None", None]:
                        diem_chinh = diem_th
                    else:
                        diem_chinh = "-"
                    try:
                        parts = [p.strip() for p in str(diem_chinh).replace(",", ".").split("/") if p.strip().replace(".", "", 1).isdigit()]
                        diem_num = float(parts[-1]) if parts else 0
                    except:
                        diem_num = 0
                    scores = [s for s in str(diem_chinh).split("/") if s.strip().isdigit()]
                    num_tests = len(scores)
                    if diem_chinh in ["", "nan", "None", None]:
                        diem_chinh = "-"
                    if diem_chinh == "-":
                        xep_loai = "-"
                    elif num_tests > 1:
                        xep_loai = "Đạt" if diem_num >= 80 else "Không đạt"
                    else:
                        if diem_num >= 95:
                            xep_loai = "Xuất sắc"
                        elif diem_num >= 80:
                            xep_loai = "Đạt"
                        elif diem_num > 0:
                            xep_loai = "Không đạt"
                        else:
                            xep_loai = "-"
                    # Ghi chú
                    if diem_chinh == "-":
                        note = "Vắng"
                    else:
                        if num_tests > 1:
                            note = f"Kiểm tra lần {'/'.join(str(i+1) for i in range(num_tests))}"
                        else:
                            note = ""
                    data.append({
                        "id": ma_nv,
                        "name": ho_ten,
                        "unit": str(row.get("Đơn vị", "") or "").strip(),
                        "score": diem_chinh,
                        "rank": xep_loai,
                        "note": note,
                        "num_tests": num_tests
                    })

            # Sắp xếp
            def calc_group_numtests_score1(student):
                if "score_tb" in student:
                    try:
                        score_1 = float(student.get("score_tb", 0))
                    except:
                        score_1 = 0
                    num_tests = student.get("num_tests", 1)
                    if student.get("score_lt", "-") == "-" and student.get("score_th", "-") == "-":
                        group = 6
                    elif num_tests == 1:
                        if score_1 >= 95:
                            group = 1
                        elif score_1 >= 80:
                            group = 2
                        else:
                            group = 4
                    else:
                        if score_1 >= 80:
                            group = 3
                        else:
                            group = 5
                else:
                    try:
                        score_1 = float(student.get("score", 0))
                    except:
                        score_1 = 0
                    num_tests = student.get("num_tests", 1)
                    if student.get("score", "-") in ["", "-", "nan", "None", None]:
                        group = 6
                    elif num_tests == 1:
                        if score_1 >= 95:
                            group = 1
                        elif score_1 >= 80:
                            group = 2
                        else:
                            group = 4
                    else:
                        if score_1 >= 80:
                            group = 3
                        else:
                            group = 5
                return group, num_tests, score_1

            for student in data:
                group, num_tests, score_1 = calc_group_numtests_score1(student)
                student["group"] = group
                student["num_tests"] = num_tests
                student["score_1"] = score_1

            data_sorted = sorted(
                data,
                key=lambda row: (
                    row["group"],
                    row["num_tests"],
                    -row["score_1"],
                    row["name"]
                )
            )

            # Tính lại số lượng nếu chưa có
            if not num_attended or not num_total:
                num_total = len(data_sorted)
                if use_5b:
                    num_attended = sum(1 for x in data_sorted if str(x.get("score_lt", "") or "").strip() not in ["", "-", "nan", "None"] or str(x.get("score_th", "") or "").strip() not in ["", "-", "nan", "None"])
                else:
                    num_attended = sum(1 for x in data_sorted if str(x.get("score", "") or "").strip() not in ["", "-", "nan", "None"])

            # Xử lý ngày
            days = extract_days(time)
            for i, student in enumerate(data_sorted):
                student["day1"] = days[0] if len(days) > 0 else ""
                student["day2"] = days[1] if len(days) > 1 else ""
                student["day3"] = days[2] if len(days) > 2 else ""

            with open(template_file, "r", encoding="utf-8") as f:
                template_str = f.read()
            template = Template(template_str)
            min_height = 120 if len(data_sorted) <= 14 else 90

            rendered = template.render(
                students=data_sorted,
                course_name=course_name,
                training_type=training_type,
                time=time,
                location=location,
                num_attended=num_attended,
                num_total=num_total,
                class_code=class_info.get("class_code", ""),
                gv_huong_dan=gv_huong_dan,
                truong_bo_mon=truong_bo_mon,
                truong_tt=truong_tt,
                logo_base64=logo_base64,
                min_height=min_height
            )

            # Xuất Excel đúng tên mã hóa
            file_basename = ma_hoa_ten_lop(course_name, time)
            file_excel = f"{file_basename}.xlsx"
            df_baocao = pd.DataFrame(data_sorted)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_baocao.to_excel(writer, index=False, sheet_name="Báo cáo")
            output.seek(0)
            excel_b64 = base64.b64encode(output.read()).decode()

            html_report = f"""
            <div style="text-align:right; margin-bottom:12px;" class="no-print">
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}"
                download="{file_excel}"
                style="display:inline-block; font-size:18px; padding:6px 18px; margin-right:16px; background:#f0f0f0; border-radius:4px; text-decoration:none; border:1px solid #ccc;">
                📥 Tải báo cáo Excel
                </a>
                <button onclick="window.print()" style="font-size:18px;padding:6px 18px;">🖨️ In báo cáo kết quả</button>
            </div>
            {rendered}
            """
            st.subheader("📄 Xem trước báo cáo")
            st.components.v1.html(html_report, height=1200, scrolling=True)

    # Nếu có nút điểm danh, tạo bảng điểm danh
    if diem_danh:
        df = ds_hocvien[(ds_hocvien["Mã NV"].astype(str).str.strip() != "") | (ds_hocvien["Họ tên"].astype(str).str.strip() != "")]
        df = df.reset_index(drop=True)
        days = extract_days(time)
        students = []
        for i, row in df.iterrows():
            diem_lt = str(row.get("Điểm LT", "") or "").strip()
            check = "X" if diem_lt and diem_lt not in ["", "-", "None"] else "V"
            students.append({
                "stt": i + 1,
                "id": str(row.get("Mã NV", "") or "").strip(),
                "name": str(row.get("Họ tên", "") or "").strip(),
                "unit": str(row.get("Đơn vị", "") or "").strip(),
                "day1": check if len(days) > 0 else "",
                "day2": check if len(days) > 1 else "",
                "day3": check if len(days) > 2 else "",
                "note": ""
            })
        students = [s for s in students if s["id"] or s["name"]]
        num_attended = sum(
            1 for s in students if "X" in [s.get("day1", ""), s.get("day2", ""), s.get("day3", "")]
        )
        with open("attendance_template.html", "r", encoding="utf-8") as f:
            template_str = f.read()
        template = Template(template_str)
        attendance_html = template.render(
            students=students,
            course_name=course_name,
            training_type=training_type,
            time=time,
            location=location,
            num_total=len(students),
            num_attended=num_attended,
            class_code=class_info.get("class_code", ""),
            gv_huong_dan=gv_huong_dan,
            days=days,
            logo_base64=logo_base64,
            min_height=120 if len(students) <= 14 else 90
        )
        attendance_html_with_print = """
        <div style="text-align:right; margin-bottom:12px;" class="no-print">
            <button onclick="window.print()" style="font-size:18px;padding:6px 18px;">🖨️ In bảng điểm danh</button>
        </div>
        """ + attendance_html
        st.components.v1.html(attendance_html_with_print, height=1000, scrolling=True)


