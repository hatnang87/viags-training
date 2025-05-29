# Ứng dụng Báo cáo Kết quả Đào tạo - VIAGS

Ứng dụng hỗ trợ nhập danh sách học viên, xử lý điểm thi, tạo báo cáo HTML và xuất Excel **theo mẫu VIAGS** (chuẩn phòng đào tạo), triển khai trên nền tảng [Streamlit](https://streamlit.io).

---

## 🚀 **Tính năng nổi bật**

- Nhập danh sách học viên (dạng bảng như Excel, copy-paste từ file ngoài).
- Nhập/sửa điểm trực tiếp, tự động kiểm tra, làm tròn, cảnh báo điểm sai.
- **Upload file điểm** Excel và tự động ghép điểm vào từng học viên (theo mã NV hoặc họ tên, chuẩn hóa).
- Sinh báo cáo **HTML** in chuẩn mẫu VIAGS, xuất file **Excel** theo kết quả đã xử lý.
- Tùy biến thông tin lớp học, chữ ký, số lượng học viên…
- Triển khai dùng trực tuyến trên **Streamlit Cloud**.

---

## 🛠️ **Hướng dẫn sử dụng (Local)**

### **1. Cài Python và các thư viện cần thiết**
```sh
pip install streamlit pandas jinja2 openpyxl xlsxwriter
