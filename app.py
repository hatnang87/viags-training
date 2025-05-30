import streamlit as st
import pandas as pd
from jinja2 import Template
import base64
import re
import io

st.set_page_config(page_title="Báo cáo kết quả đào tạo - VIAGS", layout="wide")

st.title("📋 Tạo báo cáo kết quả đào tạo - VIAGS")

tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Thông tin lớp học", 
    "2️⃣ Danh sách học viên & điểm",
    "3️⃣ Upload file điểm",
    "4️⃣ Chữ ký & xuất báo cáo"
])

with tab1:
    st.subheader("Nhập thông tin lớp học")
    class_info_sample = '''Điều khiển xe dầu kéo - Thực hành nâng cao
Bồi dưỡng kiến thức/Trực tiếp
02/01/2025
VNBA'''
    class_info_input = st.text_area("Dán vào 4 dòng gồm: Môn học, Loại hình, Thời gian, Địa điểm", value=class_info_sample, height=120)
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
    st.session_state["course_name"] = course_name
    st.session_state["training_type"] = training_type
    st.session_state["time"] = time
    st.session_state["location"] = location
    st.session_state["num_attended"] = num_attended
    st.session_state["num_total"] = num_total

with tab2:
    st.subheader("Danh sách học viên và nhập điểm")

    # 1. Khởi tạo bảng nếu chưa có
    if "ds_hocvien" not in st.session_state:
        ds_hocvien = pd.DataFrame({
            "Mã NV": [""] * 30,
            "Họ tên": [""] * 30,
            "Đơn vị": [""] * 30,
            "Điểm": [""] * 30
        })
    else:
        ds_hocvien = st.session_state["ds_hocvien"]

    # 2. Hiển thị bảng cho nhập liệu
    st.caption("📌 Dán danh sách học viên (copy từ Excel), điểm dạng 70/90 nếu thi nhiều lần.")
    ds_hocvien = st.data_editor(
        ds_hocvien,
        num_rows="dynamic",
        hide_index=False,  # Hiện index làm STT
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

    # 3. Làm sạch dữ liệu nhập
    #ds_hocvien = ds_hocvien.fillna("")
    for col in ["Mã NV", "Họ tên", "Đơn vị", "Điểm"]:
        ds_hocvien[col] = ds_hocvien[col].astype(str).str.strip()

    # 4. Tự động đánh lại STT sau khi người dùng thêm dòng
    #ds_hocvien["STT"] = list(range(1, len(ds_hocvien) + 1))

    # 5. Kiểm tra & làm tròn điểm luôn
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

    # 6. Lưu lại vào session
    st.session_state["ds_hocvien"] = ds_hocvien

    # 7. Cảnh báo nếu có lỗi
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

        ds_hocvien = st.session_state["ds_hocvien"]
        ds_hocvien["HoTenChuan"] = ds_hocvien["Họ tên"].apply(normalize_name)
        diem_map = dict(zip(df_diem["HoTenChuan"], df_diem["DiemDaXuLy"]))
        ds_hocvien["Điểm"] = ds_hocvien["HoTenChuan"].map(diem_map).fillna(ds_hocvien["Điểm"])
        st.success("Đã cập nhật điểm từ file LMS (theo họ tên)!")
        st.dataframe(ds_hocvien[["Mã NV", "Họ tên", "Điểm"]], use_container_width=True)
        st.session_state["ds_hocvien"] = ds_hocvien

    uploaded_dotthi = st.file_uploader("📥 Tải file điểm dạng đợt thi (E: điểm 1 lần, G: điểm nhiều lần)", type=["xlsx"], key="uploader_dotthi")
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
        ds_hocvien = st.session_state["ds_hocvien"]
        ds_hocvien["HoTenChuan"] = ds_hocvien["Họ tên"].apply(normalize_name)
        diem_map = dict(zip(df_dotthi["HoTenChuan"], df_dotthi["DiemDaXuLy"]))
        ds_hocvien["Điểm"] = ds_hocvien["HoTenChuan"].map(diem_map).fillna(ds_hocvien["Điểm"])

        st.success("Đã tự động cập nhật điểm từ file đợt thi!")
        st.dataframe(ds_hocvien[["Mã NV", "Họ tên", "Điểm"]], use_container_width=True)
        st.session_state["ds_hocvien"] = ds_hocvien

with tab4:
    st.subheader("Thông tin chữ ký báo cáo & Xuất báo cáo")
    gv_huong_dan = st.text_input("Họ tên Giáo viên hướng dẫn", value="Nguyễn Đức Nghĩa")
    truong_bo_mon = st.text_input("Họ tên Trưởng bộ môn", value="Ngô Trung Thành")
    truong_tt = st.text_input("Họ tên Trưởng TTĐT", value="Nguyễn Chí Kiên")

    if st.button("📄 Xem trước & In báo cáo"):
        ds_hocvien = st.session_state.get("ds_hocvien", pd.DataFrame())
        course_name = st.session_state.get("course_name", "")
        training_type = st.session_state.get("training_type", "")
        time = st.session_state.get("time", "")
        location = st.session_state.get("location", "")
        num_attended = st.session_state.get("num_attended", "")
        num_total = st.session_state.get("num_total", "")

        if ds_hocvien.empty:
            st.warning("Vui lòng nhập danh sách học viên!")
        else:
            # Lọc bỏ dòng trống
            ds_hocvien_filtered = ds_hocvien[(ds_hocvien["Mã NV"].str.strip() != "") | (ds_hocvien["Họ tên"].str.strip() != "")]
            data = []
            for i, row in ds_hocvien_filtered.iterrows():
                if (
                    not row["Mã NV"].strip() or row["Mã NV"].strip().lower() == "none"
                ) and (
                    not row["Họ tên"].strip() or row["Họ tên"].strip().lower() == "none"
                ):
                    continue  # Bỏ dòng trống hoặc None
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

            # Tính lại số học viên nếu cần
            if not num_attended or not num_total:
                num_total = len(data_sorted)
                num_attended = sum(1 for x in data_sorted if x["score"] not in ["-", ""])

            # Tính min_height cho bảng (ví dụ mỗi dòng ~10mm, tối thiểu 120mm)
            num_students = len(data_sorted)
            if num_students <= 13:
                min_height = 150
            else:
                min_height = max(150, num_students * 15)

            # Đọc file logo và chuyển sang base64
            with open("logo_viags.png", "rb") as image_file:
                logo_base64 = base64.b64encode(image_file.read()).decode()

            # Đọc template HTML
            with open("report_template.html", "r", encoding="utf-8") as f:
                template_str = f.read()
            template = Template(template_str)

            # Render template với đầy đủ biến
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
                min_height=min_height
            )

            # Tải Excel
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

            # Thêm nút in vào đầu HTML
            html_report = f"""
            <div style="text-align:right; margin-bottom:12px;" class="no-print">
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}"
                download="Bao_cao_ket_qua_dao_tao.xlsx"
                style="display:inline-block; font-size:18px; padding:6px 18px; margin-right:16px; background:#f0f0f0; border-radius:4px; text-decoration:none; border:1px solid #ccc;">
                📥 Tải báo cáo Excel
                </a>
                <button onclick="window.print()" style="font-size:18px;padding:6px 18px;">🖨️ In báo cáo</button>
            </div>
            {rendered}
"""

            st.subheader("📄 Xem trước báo cáo")
            st.components.v1.html(html_report, height=1200, scrolling=True)

            # Tải HTML
            #b64 = base64.b64encode(rendered.encode()).decode()
            #href = f'<a href="data:text/html;base64,{b64}" download="bao_cao.html">📥 Tải báo cáo HTML</a>'
            #st.markdown(href, unsafe_allow_html=True)
