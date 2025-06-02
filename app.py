import streamlit as st
import pandas as pd
from jinja2 import Template
import base64
import re
import io
import unicodedata
import openpyxl

st.set_page_config(page_title="Báo cáo kết quả đào tạo - VIAGS", layout="wide")

st.title("📋 Tạo báo cáo kết quả đào tạo - VIAGS (Nhiều lớp)")

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

def strip_accents(s):
    if not isinstance(s, str):
        return ""
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

# ========== Quản lý nhiều lớp ==========
if "danh_sach_lop" not in st.session_state:
    st.session_state["danh_sach_lop"] = {}
if "ten_lop_hien_tai" not in st.session_state:
    st.session_state["ten_lop_hien_tai"] = ""
if "hien_nhap_excel" not in st.session_state:
    st.session_state["hien_nhap_excel"] = False

# Sắp xếp danh sách lớp theo thứ tự tiếng Việt

ds_lop = sorted(df_muc_luc["MaLop"].tolist(), key=strip_accents) if not df_muc_luc.empty else []


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
            st.success(f"Đã nhập xong {so_lop_them} lớp! Vào phần 'Chọn lớp' để kiểm tra.")
            for log in log_sheets:
                st.write(log)
            st.session_state["hien_nhap_excel"] = False
            st.rerun()
        else:
            for log in log_sheets:
                st.write(log)
            st.warning("Không tìm thấy sheet nào hợp lệ (phải có D7 là tên lớp và học viên từ C14-E14).")

# Tạo lớp mới
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
        st.rerun()
    else:
        st.warning("Tên lớp đã tồn tại!")
elif ten_lop and ten_lop != "+ Tạo lớp mới":
    st.session_state["ten_lop_hien_tai"] = ten_lop

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

# ========= Tabs =========
tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Thông tin lớp học", 
    "2️⃣ Danh sách học viên & điểm",
    "3️⃣ Upload file điểm",
    "4️⃣ Chữ ký & xuất báo cáo"
])
# ========== Tab nội dung ==========

    
with tab1:
    st.subheader("Nhập thông tin lớp học")
    class_info_sample = '''An toàn hàng không
Định kỳ/Elearning+Trực tiếp
02/01/2025
TTĐT MB'''
    class_info_input = st.text_area("Dán vào 4 dòng gồm: Môn học, Loại hình, Thời gian, Địa điểm", value="\n".join([
        class_info.get("course_name", ""),
        class_info.get("training_type", ""),
        class_info.get("time", ""),
        class_info.get("location", ""),
        class_info.get("num_attended", "") + "/" + class_info.get("num_total", "") if class_info.get("num_attended", "") else ""
    ]) if any(class_info.values()) else class_info_sample, height=120)
    class_info_lines = class_info_input.strip().split("\n")
    course_name = class_info_lines[0] if len(class_info_lines) > 0 else ""
    training_type = class_info_lines[1] if len(class_info_lines) > 1 else ""
    time = class_info_lines[2] if len(class_info_lines) > 2 else ""
    location = class_info_lines[3] if len(class_info_lines) > 3 else ""
    if len(class_info_lines) > 4 and "/" in class_info_lines[4]:
        num_info = class_info_lines[4].split("/")
        num_attended = num_info[0].strip()
        num_total = num_info[1].strip()
    else:
        num_attended = ""
        num_total = ""
    # Lưu lại
    class_info = {
        "course_name": course_name,
        "training_type": training_type,
        "time": time,
        "location": location,
        "num_attended": num_attended,
        "num_total": num_total,
    }
    lop_data["class_info"] = class_info

with tab2:
    st.subheader("Danh sách học viên và nhập điểm")
    # Hiển thị bảng cho nhập liệu
    st.caption("📌 Dán danh sách học viên (copy từ Excel), điểm dạng 70/90 nếu thi nhiều lần.")
    ds_hocvien = st.data_editor(
        ds_hocvien,
        num_rows="dynamic",
        hide_index=False,
        use_container_width=True,
        column_order=["Mã NV", "Họ tên", "Đơn vị", "Điểm"],
        column_config={
            "Mã NV": st.column_config.TextColumn(width="small"),
            "Họ tên": st.column_config.TextColumn(width="large"),
            "Đơn vị": st.column_config.TextColumn(width="medium"),
            "Điểm": st.column_config.TextColumn(width="small"),
        },
        key="data_editor_ds"
    )
    # Làm sạch dữ liệu nhập
    for col in ["Mã NV", "Họ tên", "Đơn vị", "Điểm"]:
        ds_hocvien[col] = ds_hocvien[col].astype(str).str.strip()
    # Kiểm tra & làm tròn điểm
    error_rows = []
    for i, row in ds_hocvien.iterrows():
        diem_str = row.get("Điểm", "").strip()
        if diem_str:
            parts = [x.strip() for x in diem_str.split("/")]
            diem_moi = []
            for p in parts:
                try:
                    p_float = float(p.replace(",", "."))
                    p_int = int(round(p_float))
                    if 0 <= p_int <= 100:
                        diem_moi.append(str(p_int))
                    else:
                        error_rows.append((i, row["Họ tên"], p))
                except:
                    error_rows.append((i, row["Họ tên"], p))
            ds_hocvien.at[i, "Điểm"] = "/".join(diem_moi)
        else:
            ds_hocvien.at[i, "Điểm"] = ""
    # Lưu lại vào session_state
    lop_data["ds_hocvien"] = ds_hocvien
    # Cảnh báo nếu có lỗi
    if error_rows:
        st.warning("⚠️ Có điểm không hợp lệ:\n" + "\n".join([f"{idx+1} - {name} (giá trị: {val})" for idx, name, val in error_rows]))
    else:
        st.info("✅ Toàn bộ điểm đã được kiểm tra và làm tròn đúng định dạng.")

with tab3:
    st.subheader("Upload file điểm và tự động ghép điểm vào danh sách")
    uploaded_lms = st.file_uploader("📥 Tải file điểm dạng lớp học (LMS_RPT)", type=["xlsx"], key="uploader_lms")
    if uploaded_lms is not None:
        df_diem = pd.read_excel(uploaded_lms)
        col_name_hoten = df_diem.columns[3]    # Cột D
        col_name_lanthi = df_diem.columns[6]   # Cột G

        def normalize_name(s):
            import re
            return re.sub(r"\s+", "", str(s).strip().lower())
        df_diem["HoTenChuan"] = df_diem[col_name_hoten].apply(normalize_name)

        def extract_diem_lanthi(text):
            if not isinstance(text, str):
                return ""
            scores = re.findall(r"Lần \d+\s*:\s*(\d+)", text)
            return "/".join(scores)
        df_diem["DiemDaXuLy"] = df_diem[col_name_lanthi].apply(extract_diem_lanthi)

        ds_hocvien["HoTenChuan"] = ds_hocvien["Họ tên"].apply(normalize_name)
        diem_map = dict(zip(df_diem["HoTenChuan"], df_diem["DiemDaXuLy"]))
        ds_hocvien["Điểm"] = ds_hocvien["HoTenChuan"].map(diem_map).fillna(ds_hocvien["Điểm"])
        st.success("Đã cập nhật điểm từ file LMS (theo họ tên)!")
        st.dataframe(ds_hocvien[["Mã NV", "Họ tên", "Điểm"]], use_container_width=True)
        lop_data["ds_hocvien"] = ds_hocvien

    uploaded_dotthi = st.file_uploader("📥 Tải file điểm dạng đợt thi", type=["xlsx"], key="uploader_dotthi")
    if uploaded_dotthi is not None:
        df_dotthi = pd.read_excel(uploaded_dotthi)
        col_name_hoten = df_dotthi.columns[2]       # Cột C
        col_name_diem_1lan = df_dotthi.columns[4]   # Cột E
        col_name_diem_nlan = df_dotthi.columns[6]   # Cột G

        def normalize_name(s):
            return re.sub(r"\s+", "", str(s).strip().lower())

        def extract_score_dotthi(row):
            diem_1lan = row[col_name_diem_1lan]
            diem_nlan = row[col_name_diem_nlan]
            if pd.notnull(diem_nlan) and str(diem_nlan).strip() != "":
                scores = re.findall(r"Lần\s*\d+\s*:\s*(\d+)", str(diem_nlan))
                if scores:
                    return "/".join(scores)
                else:
                    return str(diem_nlan).strip()
            elif pd.notnull(diem_1lan) and str(diem_1lan).strip() != "":
                return str(diem_1lan).strip()
            else:
                return ""

        df_dotthi["HoTenChuan"] = df_dotthi[col_name_hoten].apply(normalize_name)
        df_dotthi["DiemDaXuLy"] = df_dotthi.apply(extract_score_dotthi, axis=1)
        ds_hocvien["HoTenChuan"] = ds_hocvien["Họ tên"].apply(normalize_name)
        diem_map = dict(zip(df_dotthi["HoTenChuan"], df_dotthi["DiemDaXuLy"]))
        ds_hocvien["Điểm"] = ds_hocvien["HoTenChuan"].map(diem_map).fillna(ds_hocvien["Điểm"])

        st.success("Đã tự động cập nhật điểm từ file đợt thi!")
        st.dataframe(ds_hocvien[["Mã NV", "Họ tên", "Điểm"]], use_container_width=True)
        lop_data["ds_hocvien"] = ds_hocvien

with tab4:
    st.subheader("Thông tin chữ ký báo cáo & Xuất báo cáo")
    gv_huong_dan = st.text_input("Họ tên Giáo viên hướng dẫn", value="Nguyễn Đức Nghĩa")
    truong_bo_mon = st.text_input("Họ tên Trưởng bộ môn", value="Ngô Trung Thành")
    truong_tt = st.text_input("Họ tên Trưởng TTĐT", value="Nguyễn Chí Kiên")
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

    with open("logo_viags.png", "rb") as image_file:
        logo_base64 = base64.b64encode(image_file.read()).decode()

    col1, col2,_ = st.columns([1,1,4])
    with col1:
        bckq = st.button("📄In báo cáo kết quả")
    with col2:
        diem_danh = st.button("Tạo bảng điểm danh")

    if bckq:
        ds_hocvien = lop_data["ds_hocvien"]
        course_name = class_info.get("course_name", "")
        training_type = class_info.get("training_type", "")
        time = class_info.get("time", "")
        location = class_info.get("location", "")
        num_attended = class_info.get("num_attended", "")
        num_total = class_info.get("num_total", "")

        if ds_hocvien.empty:
            st.warning("Vui lòng nhập danh sách học viên!")
        else:
            ds_hocvien_filtered = ds_hocvien[(ds_hocvien["Mã NV"].str.strip() != "") | (ds_hocvien["Họ tên"].str.strip() != "")]
            data = []
            for i, row in ds_hocvien_filtered.iterrows():
                if (
                    not row["Mã NV"].strip() or row["Mã NV"].strip().lower() == "none"
                ) and (
                    not row["Họ tên"].strip() or row["Họ tên"].strip().lower() == "none"
                ):
                    continue
                data.append({
                    "id": row["Mã NV"],
                    "name": row["Họ tên"],
                    "unit": row["Đơn vị"],
                    "raw_score": row.get("Điểm", "")
                })
            def process_student(row):
                score_str = row["raw_score"]
                if not score_str or score_str.strip() in ["-", ""]:
                    return "-", "-", "Vắng", 99, 0, 0, 6
                try:
                    scores = [int(s.strip()) for s in score_str.split("/") if s.strip().isdigit()]
                    num_tests = len(scores)
                    score_1 = scores[0] if scores else 0
                    final_score = scores[-1] if scores else 0
                    note = ""
                    if num_tests > 1:
                        note = f"Kiểm tra lần {'/'.join(str(i+1) for i in range(num_tests))}"
                    if num_tests == 1:
                        if final_score >= 95:
                            group = 1
                            rank = "Xuất sắc"
                        elif final_score >= 80:
                            group = 2
                            rank = "Đạt"
                        else:
                            group = 4
                            rank = "Không đạt"
                    elif num_tests >= 2:
                        if final_score >= 80:
                            group = 3
                            rank = "Đạt"
                        else:
                            group = 5
                            rank = "Không đạt"
                    else:
                        group = 6
                        rank = "-"
                    return score_str, rank, note, num_tests, -score_1, score_1, group
                except:
                    return "-", "-", "Vắng", 99, 0, 0, 6

            for row in data:
                row["score"], row["rank"], row["note"], row["num_tests"], row["sort_score"], row["score_1"], row["group"] = process_student(row)

            def full_sort_key(row):
                return (
                    row["group"],
                    row["num_tests"],
                    -row["score_1"]
                )

            data_sorted = sorted(data, key=full_sort_key)

            if not num_attended or not num_total:
                num_total = len(data_sorted)
                num_attended = sum(1 for x in data_sorted if x["score"] not in ["-", ""])

            num_students = len(data_sorted)
            if num_students <= 14:
                min_height = 120
            else:
                min_height = 90

            days = extract_days(time)
            for i, student in enumerate(data_sorted):
                student["day1"] = days[0] if len(days) > 0 else ""
                student["day2"] = days[1] if len(days) > 1 else ""
                student["day3"] = days[2] if len(days) > 2 else ""

            with open("report_template.html", "r", encoding="utf-8") as f:
                template_str = f.read()
            template = Template(template_str)
            rendered = template.render(
                students=data_sorted,
                course_name=course_name,
                training_type=training_type,
                time=time,
                location=location,
                num_attended=num_attended,
                num_total=num_total,
                gv_huong_dan=gv_huong_dan,
                truong_bo_mon=truong_bo_mon,
                truong_tt=truong_tt,
                logo_base64=logo_base64,
                min_height=min_height,
                num_students=num_students
            )
            df_baocao = pd.DataFrame(data_sorted)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_baocao.to_excel(writer, index=False, sheet_name="Báo cáo")
            output.seek(0)
            excel_b64 = base64.b64encode(output.read()).decode()
            excel_link = f'''
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}"
            download="Bao_cao_ket_qua_dao_tao.xlsx"
            style="display:inline-block; font-size:18px; padding:6px 18px; margin-right:16px; background:#f0f0f0; border-radius:4px; text-decoration:none; border:1px solid #ccc;">
            📥 Tải báo cáo Excel
            </a>'''

            html_report = f"""
            <div style="text-align:right; margin-bottom:12px;" class="no-print">
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}"
                download="Bao_cao_ket_qua_dao_tao.xlsx"
                style="display:inline-block; font-size:18px; padding:6px 18px; margin-right:16px; background:#f0f0f0; border-radius:4px; text-decoration:none; border:1px solid #ccc;">
                📥 Tải báo cáo Excel
                </a>
                <button onclick="window.print()" style="font-size:18px;padding:6px 18px;">🖨️ In báo cáo kết quả</button>
            </div>
            {rendered}
            """

            st.subheader("📄 Xem trước báo cáo")
            st.components.v1.html(html_report, height=1200, scrolling=True)
    if diem_danh:
        ds_hocvien = lop_data["ds_hocvien"]
        df = ds_hocvien[(ds_hocvien["Mã NV"].str.strip() != "") | (ds_hocvien["Họ tên"].str.strip() != "")]
        df = df.reset_index(drop=True)
        days = extract_days(class_info.get("time", ""))
        students = []
        for i, row in df.iterrows():
            diem = row.get("Điểm", "").strip()
            check = "X" if diem and diem not in ["", "-", "None"] else "V"
            students.append({
                "stt": i + 1,
                "id": row["Mã NV"],
                "name": row["Họ tên"],
                "unit": row["Đơn vị"],
                "day1": check if len(days) > 0 else "",
                "day2": check if len(days) > 1 else "",
                "day3": check if len(days) > 2 else "",
                "note": ""
            })
        num_attended = sum(
            1 for s in students if "X" in [s.get("day1", ""), s.get("day2", ""), s.get("day3", "")]
        )
        with open("attendance_template.html", "r", encoding="utf-8") as f:
            template_str = f.read()
        template = Template(template_str)
        attendance_html = template.render(
            students=students,
            course_name=class_info.get("course_name", ""),
            training_type=class_info.get("training_type", ""),
            time=class_info.get("time", ""),
            location=class_info.get("location", ""),
            num_total=len(students),
            num_attended=num_attended,
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

# Ghi lại dữ liệu lớp về session_state (cực kỳ quan trọng)
st.session_state["danh_sach_lop"][st.session_state["ten_lop_hien_tai"]] = {
    "class_info": class_info,
    "ds_hocvien": ds_hocvien,
}
