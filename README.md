# 🧾 InvoicesExtract AI — Gemini API Automation Tool

[![Language](https://img.shields.io/badge/Language-Python%203.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)](#)
[![AI Model](https://img.shields.io/badge/AI-Gemini%202.5%20Flash-F9AB00?style=for-the-badge&logo=google&logoColor=white)](#)
[![GUI](https://img.shields.io/badge/GUI-Tkinter-lightgrey?style=for-the-badge)](#)

> **InvoicesExtract AI** là công cụ tự động hóa trích xuất dữ liệu từ các hóa đơn, biên lai (PDF, Hình ảnh, XML) bằng sức mạnh của **Google Gemini API**. Hệ thống giúp số hóa dữ liệu hóa đơn hàng loạt, tự động phân loại tiền tệ và xuất ra file Excel chuẩn chỉnh, giúp kế toán viên tiết kiệm hàng giờ nhập liệu thủ công và hạn chế tối đa sai sót.

---

**Download latest version of InvoicesExtract(.exe):** [Download](https://github.com/cronpostps/InvoicesExtract/releases/latest/download/InvoicesExtract.zip)

---

> 💡 **Tip cho Coder:** Hãy copy-paste toàn bộ nội dung tài liệu này cho một AI (ChatGPT/Gemini/Claude...) để nó có ngữ cảnh và giúp bạn gỡ lỗi, nâng cấp thêm tính năng (như thêm progress bar, tối ưu prompt) nhanh chóng hơn!

---

## ✨ 1. Giới thiệu & Các tính năng cốt lõi

Dự án này là giải pháp thay thế cho việc gõ lại từng dòng hóa đơn bằng tay. Nó được xây dựng dưới dạng ứng dụng Desktop (Windows) nhẹ nhàng với giao diện trực quan.

**🚀 Ưu điểm vượt trội**

* **Siêu tốc độ & Thông minh:** Sử dụng model `gemini-2.5-flash` mới nhất từ Google, kết hợp với Pydantic Schema để ép AI trả về dữ liệu cấu trúc chuẩn JSON 100%.
* **Hỗ trợ Đa định dạng nguyên bản:** Đọc mượt mà hóa đơn PDF, ảnh chụp (JPG, PNG) và cả file XML hóa đơn điện tử gốc.
* **Xử lý Đa tiền tệ linh hoạt:** Tự động nhận diện. Nếu là VNĐ -> ép kiểu số nguyên (int). Nếu là Ngoại tệ (USD, EUR...) -> giữ nguyên số thập phân (float).
* **Bảo vệ tiến trình (Smart Retry & Thread-Safe):** Xử lý đa luồng giúp giao diện không bao giờ bị treo (crash). Tự động tính toán thời gian chờ (Exponential Backoff) nếu API bị quá tải (lỗi 429).
* **Dừng Khẩn cấp (STOP):** Nút dừng an toàn, tự động hoàn thiện file đang quét dở, lưu data vào Excel rồi mới thoát tiến trình, đảm bảo không bao giờ hỏng dữ liệu.

---

## 🛠️ 2. Môi trường chuẩn bị

Để tự triển khai (deploy) hoặc biên dịch mã nguồn dự án này, bạn cần chuẩn bị sẵn:

1. **Tài khoản Google AI Studio:** Đăng ký tại `aistudio.google.com` để lấy API Key miễn phí (hoặc trả phí tùy nhu cầu).
2. **Tài khoản GitHub:** Dùng để lưu trữ mã nguồn cá nhân.
3. **Môi trường máy tính:** Đã cài sẵn Python (Phiên bản từ 3.8 trở lên) và pip.

---

## 🚀 3. Hướng dẫn biên dịch (Build) từ Source - Dành cho Lập trình viên

Nếu bạn không muốn tải file `.exe` có sẵn mà muốn tự chạy hoặc đóng gói phần mềm trên máy tính của mình, hãy thực hiện tuần tự các bước sau:

**Bước 3.1: Tải mã nguồn về máy tính**

Mở Terminal (CMD / PowerShell / VS Code) và chạy lệnh:
```bash
git clone https://github.com/cronpostps/InvoicesExtract.git
cd InvoicesExtract
```

**Bước 3.2: Cài đặt các thư viện phụ thuộc**
Chạy lệnh sau để cài đặt Google GenAI SDK mới nhất và các thư viện xử lý Excel, dữ liệu:
```bash
pip install google-genai openpyxl pydantic pyinstaller
```

**Bước 3.3: Chạy thử (Debug Mode)**
Để kiểm tra phần mềm hoạt động tốt trên máy:
```bash
python InvoicesExtract.py
```

**Bước 3.4: Đóng gói thành file chạy (.exe)**
Nếu bạn muốn tạo file `.exe` độc lập (ẩn cửa sổ console đen) để gửi cho phòng kế toán dùng:
```bash
python -m PyInstaller --noconfirm --onefile --windowed .\InvoicesExtract.py
```
*(Sau khi chạy xong, file `InvoicesExtract.exe` sẽ nằm gọn gàng trong thư mục `dist`).*

---

## 🔑 4. Hướng dẫn sử dụng nhanh

1. Mở phần mềm `InvoicesExtract.exe`.
2. Dán **Gemini API Key** của bạn vào ô tương ứng (Hệ thống sẽ tự động ghi nhớ cho lần sau).
3. Bấm **Chọn Thư Mục** và trỏ đến folder đang chứa các file hóa đơn (PDF/JPG/XML).
4. Bấm **BẮT ĐẦU XỬ LÝ** và ngồi nhâm nhi cà phê! ☕ 
5. Toàn bộ dữ liệu trích xuất thành công sẽ được tự động ghi nối tiếp vào file `invoice_extract.xlsx` nằm ngay bên trong thư mục hóa đơn mà bạn vừa chọn.

---

## ☎️ 5. Đóng góp & Mã nguồn mở

Mọi đóng góp (Pull Requests), tối ưu code, báo lỗi (Issues) đều được hoan nghênh tại kho lưu trữ chính thức!

**GitHub Repository:** [github.com/cronpostps/InvoicesExtract](https://github.com/cronpostps/InvoicesExtract)

**File thực thi (.exe):** [Download](https://github.com/cronpostps/InvoicesExtract/releases/latest/download/InvoicesExtract.zip)

## License
Dự án này được cấp phép theo các điều khoản của [MIT License](LICENSE). Mọi người được phép tự do sử dụng, chỉnh sửa và phân phối.

Mã nguồn gốc © [InvoicesExtract AI by anhnn](https://github.com/cronpostps/InvoicesExtract)

Xem thêm các dự án do [anhnn](https://t.me/anhnn83) dev tại [anhnn.cronpost.com](https://anhnn.cronpost.com)