import streamlit as st
import pandas as pd
from jinja2 import Template
import base64
import re
import io
# Thiết lập cấu hình trang
st.set_page_config(page_title="Báo cáo kết quả đào tạo", layout="wide")

st.title("📋 Tạo báo cáo kết quả đào tạo - VIAGS")

# --- Nhập thông tin lớp học dạng dán hàng loạt ---
st.subheader("1️⃣ Dán thông tin lớp học (5 dòng, mỗi dòng 1 mục)")
class_info_sample = '''Điều khiển xe dầu kéo - Thực hành nâng cao
Bồi dưỡng kiến thức/Trực tiếp
02/01/2025
VNBA
'''
class_info_input = st.text_area("Dán vào 5 dòng gồm: Môn học, Loại hình, Thời gian, Địa điểm, Số tham dự/Tổng", value=class_info_sample, height=140)

# Xử lý thông tin lớp học
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
    # Tạm để trống, sẽ gán lại sau nếu có bảng điểm
    num_attended = ""
    num_total = ""

st.title("📋 Nhập danh sách học viên và điểm")

# Gợi ý bảng mẫu trống để dán danh sách
if "ds_hocvien" not in st.session_state:
    ds_hocvien = pd.DataFrame({
        "Mã NV": [""]*10,
        "Họ tên": [""]*10,
        "Đơn vị": [""]*10,
        "Điểm": [""]*10
    })
    st.session_state["ds_hocvien"] = ds_hocvien
else:
    ds_hocvien = st.session_state["ds_hocvien"]

# --- Sắp xếp ---
sort_col = st.selectbox("Sắp xếp theo cột:", ["Họ tên", "Mã NV"])
sort_asc = st.radio("Chiều sắp xếp:", ["Tăng dần", "Giảm dần"], horizontal=True)
ascending = sort_asc == "Tăng dần"
ds_hocvien = ds_hocvien.sort_values(by=sort_col, ascending=ascending, ignore_index=True)

# --- Bảng nhập danh sách và điểm ---
st.info("Có thể dán toàn bộ danh sách học viên từ Excel vào bảng này, rồi nhập/copy điểm sau.")
ds_hocvien = st.data_editor(
    ds_hocvien,
    num_rows="dynamic",
    hide_index=True,
    use_container_width=True,
    column_order=["Mã NV", "Họ tên", "Đơn vị", "Điểm"],
    column_config={
        "Mã NV": st.column_config.TextColumn(width="small"),
        "Họ tên": st.column_config.TextColumn(width="large"),
        "Đơn vị": st.column_config.TextColumn(width="medium"),
        "Điểm": st.column_config.TextColumn(width="small"),
    }
)
ds_hocvien = ds_hocvien.fillna("")
for col in ["Mã NV", "Họ tên", "Đơn vị", "Điểm"]:
    ds_hocvien[col] = ds_hocvien[col].astype(str).str.strip().replace("None", "")
# Lưu lại danh sách học viên vào session state
st.session_state["ds_hocvien"] = ds_hocvien

# --- Tự động kiểm tra, làm tròn điểm ---
error_rows = []
for i, row in ds_hocvien.iterrows():
    diem_str = str(row.get("Điểm", "")).strip()
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

if error_rows:
    st.warning("Có giá trị điểm không hợp lệ (ngoài 0-100, hoặc ký tự lạ) ở các dòng:\n" +
        "\n".join([f"{idx+1} - {name} (giá trị: {val})" for idx, name, val in error_rows]))
else:
    st.info("Toàn bộ điểm đã được kiểm tra và làm tròn đúng định dạng!")

st.session_state["ds_hocvien"] = ds_hocvien

#Map điểm từ file Excel
uploaded_file = st.file_uploader("📥 Tải lên file Excel điểm (LMS_RPT_001A.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    df_diem = pd.read_excel(uploaded_file, header=0)
    #st.write("Tên các cột file điểm:", df_diem.columns.tolist())
    # Lấy cột họ tên và cột G (Lần thi)
    col_name_hoten = df_diem.columns[3]   # Cột D: Họ và tên
    col_name_lanthi = df_diem.columns[6]  # Cột G: Lần thi

    def normalize_name(s):
        return re.sub(r"\s+", "", str(s).strip().lower())
    df_diem["HoTenChuan"] = df_diem[col_name_hoten].apply(normalize_name)

    def extract_diem_lanthi(text):
        if not isinstance(text, str):
            return ""
        scores = re.findall(r"Lần \d+\s*:\s*(\d+)", text)
        return "/".join(scores)
    df_diem["Điểm đã xử lý"] = df_diem[col_name_lanthi].apply(extract_diem_lanthi)

    # Chuẩn hóa họ tên bảng học viên
    ds_hocvien["HoTenChuan"] = ds_hocvien["Họ tên"].apply(normalize_name)
    diem_map = dict(zip(df_diem["HoTenChuan"], df_diem["Điểm đã xử lý"]))
    ds_hocvien["Điểm"] = ds_hocvien["HoTenChuan"].map(diem_map).fillna(ds_hocvien["Điểm"])

    st.success("Đã tự động cập nhật điểm từ file Excel dựa theo họ tên!")
    st.dataframe(ds_hocvien[["Mã NV", "Họ tên", "Điểm"]], use_container_width=True)
    st.session_state["ds_hocvien"] = ds_hocvien


# --- Thông tin chữ ký báo cáo ---
st.subheader("3️⃣ Thông tin chữ ký báo cáo")
gv_huong_dan = st.text_input("Họ tên Giáo viên hướng dẫn", value="Nguyễn Đức Nghĩa")
truong_bo_mon = st.text_input("Họ tên Trưởng bộ môn", value="Ngô Trung Thành")
truong_tt = st.text_input("Họ tên Trưởng TTĐT", value="Nguyễn Chí Kiên")

# --- Nút tạo báo cáo ---
if st.button("📄 Tạo báo cáo"):
    ds_hocvien = st.session_state.get("ds_hocvien", pd.DataFrame())
    if ds_hocvien.empty:
        st.warning("Vui lòng nhập danh sách học viên!")
    else:
        data = []
        for i, row in ds_hocvien.iterrows():
            data.append({
                "id": row["Mã NV"],
                "name": row["Họ tên"],
                "unit": row["Đơn vị"],
                "raw_score": row.get("Điểm", "")
            })

        # ... (process_student, full_sort_key như cũ)
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

        # Sau khi xử lý data_sorted:
        if not num_attended or not num_total:
            num_total = len(data_sorted)
            num_attended = sum(1 for x in data_sorted if x["score"] not in ["-", ""])

        # Render HTML
        with open("report_template.html", "r", encoding="utf-8") as f:
            template_str = f.read()
        template = Template(template_str)
        rendered = template.render(
            course_name=course_name,
            training_type=training_type,
            time=time,
            location=location,
            num_attended=num_attended,
            num_total=num_total,
            students=data_sorted,
            gv_huong_dan=gv_huong_dan,
            truong_bo_mon=truong_bo_mon,
            truong_tt=truong_tt
        )

        st.subheader("📄 Xem trước báo cáo")
        st.components.v1.html(rendered, height=900, scrolling=True)

        # Tải HTML
        b64 = base64.b64encode(rendered.encode()).decode()
        href = f'<a href="data:text/html;base64,{b64}" download="bao_cao.html">📥 Tải báo cáo HTML</a>'
        st.markdown(href, unsafe_allow_html=True)

        # Tải Excel
        df_baocao = pd.DataFrame(data_sorted)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_baocao.to_excel(writer, index=False, sheet_name="Báo cáo")
        output.seek(0)
        st.download_button(
            label="📥 Tải báo cáo Excel",
            data=output,
            file_name="Bao_cao_ket_qua_dao_tao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
