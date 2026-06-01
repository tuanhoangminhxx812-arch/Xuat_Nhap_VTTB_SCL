# 📊 Hệ Thống Tổng Hợp Vật Tư Thiết Bị & Tách Chi Phí SCL

Ứng dụng Streamlit Python chuyên nghiệp dùng để tự động xử lý, chuẩn hóa dữ liệu và xuất hai báo cáo kế toán vật tư độc lập cho Công ty Điện lực Vũng Tàu từ hai file nguồn **Nhập (`INV-007A`)** và **Xuất (`INV-009`)**:

1. **Báo cáo gộp tổng hợp (`Xuat_Nhap_TongHop.xlsx`)**: Gộp toàn bộ giao dịch kế toán theo form mẫu chuẩn `Xuat_Nhap(mau).xlsx`.
2. **Báo cáo tách chi phí Trung/Hạ thế (`Tach_ChiPhi_PP_BL.xlsx`)**: Phân tách vật tư sửa chữa lớn (SCL) thành **Trung áp (Trung thế)** và **Hạ áp (Hạ thế)** theo từng tháng, sau đó áp dụng công thức phân bổ Khâu Phân Phối (80.79%) và Khâu Bán Lẻ (19.21%) dưới dạng công thức Excel động theo mẫu chuẩn `Tách PP-BL.xlsx`.

---

## ✨ Các Tính Năng Vượt Trội

### 1. Phân Tách Trung/Hạ Thế Bằng Trí Tuệ Nhân Tạo & Bản Đồ Tham Chiếu
* **Bản đồ phân loại chính xác tuyệt đối**: Ứng dụng tự động đọc file tham chiếu `TachPP_BL mẫu.xlsx` để xây dựng từ điển phân loại vật tư cho hơn 100+ mã vật tư SCL đã được duyệt tay, đảm bảo độ chính xác tuyệt đối.
* **Bộ lọc nhận diện thông minh (Fallback Heuristics)**: Đối với các vật tư mới phát sinh, hệ thống sử dụng thuật toán nhận diện tiền tố mã vật tư (ví dụ: `3.53.60`, `3.53.65`, `3.56` là Trung áp; `3.53.05`, `3.53.08` là Hạ áp) kết hợp phân tích chuỗi ký tự tên vật tư (bắt các từ khóa như `24KV`, `22KV`, `600V`, `HẠ THẾ`...) để phân loại chính xác.

### 2. Định Dạng Đa Sheet Theo Tháng (Multi-Sheet Workbook)
* Tự động tách dữ liệu phát sinh theo từng tháng (ví dụ: Sheet `Tháng 1`, Sheet `Tháng 2`...) tạo thành các bảng báo cáo riêng biệt nằm chung trong một file Excel.
* Tự động xóa các dòng mẫu cũ và chèn dữ liệu mới động mà không làm ảnh hưởng đến các khối tiêu đề (Header) hay chân trang (Footer).

### 3. Công Thức Excel Động (Live Formulas)
* Điền trực tiếp các công thức Excel vào ô thay vì giá trị tĩnh:
  * Công thức tính khâu Phân Phối: `=E{dòng}*80.79%`
  * Công thức tính khâu Bán Lẻ: `=E{dòng}-F{dòng}`
  * Công thức tính Tổng cộng ở chân trang: `=SUM(E12:E{dòng_cuối})`
  * Công thức liên kết tổng thanh toán ở phần đầu: `=E{dòng_tổng}`
* Giúp người dùng có thể tự chỉnh sửa dữ liệu trên Excel sau này và các số liệu tổng hợp sẽ tự động nhảy theo.

### 4. Xử Lý Trùng Chập & Gom Ô Thông Minh (Dynamic Cell Merging)
* Hệ thống tự động phân tích số lượng dự án SCL phát sinh trong tháng để tạo ra các nhóm dòng tương ứng.
* Tự động unmerge các ô mẫu cũ và merge động các ô STT (`A12:A13`), Tên công trình (`B12:B13`) cho từng dự án, và ô "Tổng cộng" (`A14:D14`) ở dòng cuối cùng của bảng dữ liệu thực tế, loại bỏ hoàn toàn các lỗi cảnh báo chồng chéo hoặc đè merged cells trong Excel.

### 5. Giám Sát Tự Động Song Song (Double Auto-Watcher)
* Hệ thống liên tục theo dõi sự thay đổi của các file gốc trên ổ đĩa. Khi bạn lưu đè hoặc cập nhật file Nhập/Xuất mới, **cả 2 file báo cáo** (`Xuat_Nhap_TongHop.xlsx` và `Tach_ChiPhi_PP_BL.xlsx`) sẽ lập tức được tự động sinh mới song song chỉ trong chưa đầy 1 giây.

---

## 🛠️ Hướng Dẫn Cài Đặt & Sử Dụng

### 1. Cài đặt môi trường
Đảm bảo bạn đã cài đặt Python 3.10+ trên hệ thống. Di chuyển vào thư mục dự án và cài đặt các thư viện cần thiết thông qua file `requirements.txt`:

```bash
pip install -r requirements.txt
```

### 2. Khởi chạy ứng dụng
Chạy lệnh Streamlit để mở giao diện web trên trình duyệt:

```bash
streamlit run app.py
```

* **Địa chỉ cục bộ (Local URL)**: [http://localhost:8501](http://localhost:8501)
* **Địa chỉ mạng nội bộ (Network URL)**: [http://10.191.74.68:8501](http://10.191.74.68:8501) (dành cho các máy tính khác trong LAN truy cập)

---

## 📊 Kết Quả Đối Khớp Dữ Liệu SCL (Tháng 3)

Số liệu phân tách Trung/Hạ áp của tháng 3 năm 2026 hoàn toàn khớp chính xác tuyệt đối với số liệu đối soát mẫu:

* **Trung thế (Trung áp) - Dự án VTAD2606001**: `320,378,418 VNĐ` (Khớp 100% ✅)
* **Hạ thế (Hạ áp) - Dự án VTAD2606001**: `31,030,578 VNĐ` (Khớp 100% ✅)
* **Khâu Phân Phối (80.79%)**: Tính bằng công thức Excel trực quan.
* **Khâu Bán Lẻ (19.21%)**: Tính bằng công thức Excel trực quan.

---

## 📁 Cấu Trúc Thư Mục Dự Án

* `app.py`: Giao diện Web Streamlit 2 phân hệ (Tab 1: Báo cáo gộp, Tab 2: Tách chi phí Trung/Hạ áp), hiển thị KPI và biểu đồ trực quan.
* `data_processor.py`: Lõi xử lý dữ liệu chính (parses Master-Detail, phân loại cấp điện áp tự động kết hợp bảng tham chiếu `TachPP_BL mẫu.xlsx`, ghi công thức động và merge ô Excel thông minh).
* `requirements.txt`: Các thư viện phụ thuộc (`pandas`, `openpyxl`, `streamlit`).
* `TachPP_BL mẫu.xlsx`: File tham chiếu chứa danh sách phân loại điện áp mẫu từ kế toán.
* `Tách PP-BL.xlsx`: File mẫu định dạng tách khâu phân phối - bán lẻ.
* `Xuat_Nhap(mau).xlsx`: File mẫu định dạng báo cáo gộp.
* `Xuat_Nhap_TongHop.xlsx`: Báo cáo gộp đầu ra (tự động cập nhật).
* `Tach_ChiPhi_PP_BL.xlsx`: Báo cáo tách chi phí đầu ra (tự động cập nhật).
