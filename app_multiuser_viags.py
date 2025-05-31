# app_multiuser_viags.py
import streamlit as st
import pandas as pd
import sqlite3
from sqlalchemy import create_engine
import streamlit_authenticator as stauth
import base64
import io
import re
from jinja2 import Template

# ---------- CẤU HÌNH ĐĂNG NHẬP ----------
names = ["Nguyễn A", "Trần B"]
usernames = ["user_a", "user_b"]
passwords = stauth.Hasher(["pass123", "pass456"]).generate()

authenticator = stauth.Authenticate(
    names, usernames, passwords,
    "viags_app", "abcdef", cookie_expiry_days=30
)

name, authentication_status, username = authenticator.login("🔐 Đăng nhập", "main")

# ---------- CƠ SỞ DỮ LIỆU ----------
engine = create_engine("sqlite:///viags.db")
conn = engine.connect()

def init_db():
    conn.execute('''CREATE TABLE IF NOT EXISTS classes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        created_by TEXT,
        course_name TEXT,
        training_type TEXT,
        time TEXT,
        location TEXT
    )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        class_id INTEGER,
        employee_id TEXT,
        name TEXT,
        unit TEXT,
        score TEXT
    )''')

init_db()

# ---------- ĐỌC TEMPLATE HTML ----------
def get_template(template_name):
    with open(template_name, "r", encoding="utf-8") as f:
        return Template(f.read())

with open("logo_viags.png", "rb") as image_file:
    logo_base64 = base64.b64encode(image_file.read()).decode()

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

# ---------- GIAO DIỆN CHÍNH ----------
if authentication_status:
    authenticator.logout("🚪 Đăng xuất", "sidebar")
    st.title("📋 Quản lý lớp học & kết quả đào tạo - VIAGS")

    menu = st.sidebar.selectbox("Chọn chức năng", ["📘 Tạo lớp học", "📄 Nhập/sửa danh sách học viên", "📊 Báo cáo & Điểm danh"])

    if menu == "📘 Tạo lớp học":
        st.subheader("📘 Nhập thông tin lớp học mới")
        course_name = st.text_input("Môn học")
        training_type = st.text_input("Loại hình đào tạo")
        time = st.text_input("Thời gian")
        location = st.text_input("Địa điểm")
        if st.button("➕ Lưu lớp học"):
            conn.execute("INSERT INTO classes (created_by, course_name, training_type, time, location) VALUES (?, ?, ?, ?, ?)",
                         (username, course_name, training_type, time, location))
            st.success("✅ Đã lưu lớp học mới!")

    elif menu == "📄 Nhập/sửa danh sách học viên":
        st.subheader("📄 Chọn lớp để nhập danh sách học viên")
        df_lop = pd.read_sql("SELECT * FROM classes", conn)
        lop_dict = {f"[{r['id']}] {r['course_name']} ({r['time']})": r['id'] for i, r in df_lop.iterrows()}
        ten_lop = st.selectbox("Chọn lớp", list(lop_dict.keys()))
        class_id = lop_dict[ten_lop]

        df_sv = pd.read_sql(f"SELECT * FROM students WHERE class_id = {class_id}", conn)

        edited_df = st.data_editor(df_sv.drop(columns=["id", "class_id"], errors="ignore"),
                                   num_rows="dynamic",
                                   use_container_width=True,
                                   key="editor")
        if st.button("💾 Lưu danh sách"):
            conn.execute(f"DELETE FROM students WHERE class_id = {class_id}")
            for _, row in edited_df.iterrows():
                conn.execute("INSERT INTO students (class_id, employee_id, name, unit, score) VALUES (?, ?, ?, ?, ?)",
                             (class_id, row['employee_id'], row['name'], row['unit'], row.get('score', '')))
            st.success("✅ Đã lưu danh sách học viên!")

    elif menu == "📊 Báo cáo & Điểm danh":
        st.subheader("📊 Chọn lớp và mẫu báo cáo")
        df_lop = pd.read_sql("SELECT * FROM classes", conn)
        lop_dict = {f"[{r['id']}] {r['course_name']} ({r['time']})": r['id'] for i, r in df_lop.iterrows()}
        ten_lop = st.selectbox("Chọn lớp để in", list(lop_dict.keys()), key="report")
        class_id = lop_dict[ten_lop]
        class_info = df_lop[df_lop['id'] == class_id].iloc[0]
        df_sv = pd.read_sql(f"SELECT * FROM students WHERE class_id = {class_id}", conn)

        df_sv_filtered = df_sv[(df_sv['employee_id'].str.strip() != '') | (df_sv['name'].str.strip() != '')].copy()

        template_type = st.radio("Chọn loại mẫu báo cáo", ["Báo cáo kết quả đào tạo", "Bảng điểm danh"])

        students = []
        days = extract_days(class_info['time'])

        for i, row in df_sv_filtered.iterrows():
            diem = row.get("score", "").strip()
            check = "X" if diem and diem not in ["", "-", "None"] else "V"
            s = {
                "stt": i + 1,
                "id": row['employee_id'],
                "name": row['name'],
                "unit": row['unit'],
                "score": diem,
                "day1": check if len(days) > 0 else "",
                "day2": check if len(days) > 1 else "",
                "day3": check if len(days) > 2 else "",
                "note": ""
            }
            students.append(s)

        if template_type == "Bảng điểm danh":
            template = get_template("attendance_template.html")
            rendered = template.render(
                students=students,
                course_name=class_info['course_name'],
                training_type=class_info['training_type'],
                time=class_info['time'],
                location=class_info['location'],
                num_attended=sum(1 for s in students if "X" in [s['day1'], s['day2'], s['day3']]),
                num_total=len(students),
                gv_huong_dan=name,
                days=days,
                logo_base64=logo_base64,
                min_height=120
            )
        else:
            def process_student(row):
                score_str = row['score']
                if not score_str or score_str.strip() in ['-', '']:
                    return '-', '-', 'Vắng', 99, 0, 0, 6
                try:
                    scores = [int(s.strip()) for s in score_str.split("/") if s.strip().isdigit()]
                    num_tests = len(scores)
                    score_1 = scores[0] if scores else 0
                    final_score = scores[-1] if scores else 0
                    note = f"Kiểm tra lần {'/'.join(str(i+1) for i in range(num_tests))}" if num_tests > 1 else ""
                    if num_tests == 1:
                        group, rank = (1, "Xuất sắc") if final_score >= 95 else (2, "Đạt") if final_score >= 80 else (4, "Không đạt")
                    else:
                        group, rank = (3, "Đạt") if final_score >= 80 else (5, "Không đạt")
                    return score_str, rank, note, num_tests, -score_1, score_1, group
                except:
                    return '-', '-', 'Vắng', 99, 0, 0, 6

            for s in students:
                s['raw_score'] = s['score']
                s['score'], s['rank'], s['note'], s['num_tests'], s['sort_score'], s['score_1'], s['group'] = process_student(s)

            students = sorted(students, key=lambda r: (r['group'], r['num_tests'], -r['score_1']))

            template = get_template("report_template.html")
            rendered = template.render(
                students=students,
                course_name=class_info['course_name'],
                training_type=class_info['training_type'],
                time=class_info['time'],
                location=class_info['location'],
                num_attended=sum(1 for x in students if x['score'] != '-'),
                num_total=len(students),
                gv_huong_dan=name,
                truong_bo_mon="Ngô Trung Thành",
                truong_tt="Nguyễn Chí Kiên",
                logo_base64=logo_base64,
                min_height=120
            )

        st.components.v1.html(rendered, height=1100, scrolling=True)

elif authentication_status is False:
    st.error("❌ Sai tài khoản hoặc mật khẩu.")
else:
    st.warning("⏳ Vui lòng đăng nhập để tiếp tục.")
