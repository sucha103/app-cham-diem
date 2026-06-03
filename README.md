# 📝 Mini App Chấm Bài 28 Ngày — Rèn Tư Duy Qua Viết

Ứng dụng web tĩnh (Static Web App) hỗ trợ giảng viên chấm bài viết của học viên bằng mô hình AI Gemma và đồng bộ tự động dữ liệu lên Google Sheets.

## ✨ Tính năng nổi bật
* **Đọc file .docx trực tiếp:** Trích xuất thông tin học viên (Mã HV, Họ tên, Ngày học) và nội dung bài làm ngay trên trình duyệt mà không cần tải lên server.
* **Chấm bài bằng AI (Gemma Skills):** Tuỳ chọn bật/tắt các tiêu chí chấm điểm chi tiết (Tư duy phản tư, Logic, Hành văn, Đồng hành & Khích lệ).
* **Vòng lặp góp ý (AI Rewrite Loop):** Giảng viên có thể ghi phản hồi, AI sẽ viết lại nhận xét nháp và tự động đồng bộ đè lên Google Sheet.
* **Giao diện hiện đại:** Thiết kế kính mờ (Glassmorphism), hỗ trợ đầy đủ chế độ Sáng/Tối (Light/Dark Mode) và co giãn tối ưu cho thiết bị di động (Responsive).
* **Đồng bộ ngoại tuyến:** Lưu trữ hàng đợi lỗi cục bộ để đồng bộ lại khi mạng internet ổn định.

## 🛠️ Cài đặt & Triển khai
1. Xem nội dung file `Code.gs` và dán vào dự án Google Apps Script của bạn. Triển khai dưới dạng **Ứng dụng web (Web App)** với quyền truy cập **Mọi người (Anyone)**.
2. Mở file `index.html` trực tiếp hoặc truy cập link GitHub Pages của bạn.
3. Nhấp nút **Cấu hình** ⚙️ ở góc trên bên phải giao diện để điền khóa **Google AI Studio API Key** và **URL Apps Script Web App** của bạn.
