---
name: dev-master-python-gas-html
description: 'Khung kỹ năng tổng hợp nâng cao: Tự động phân tích ý đồ người dùng, kích hoạt câu hỏi ngược A/B/C, tuân thủ quy tắc Workspace và lập trình chuyên sâu Google Apps Script, HTML Web App, Python.'
argument-hint: Nhập yêu cầu, ý tưởng tính năng hoặc mô tả bài toán cần giải quyết.
disable-model-invocation: false
---

# BẢN NGUYÊN TẮC LẬP TRÌNH MASTER (GAS, HTML, PYTHON)

Mục tiêu tối cao: Loại bỏ 100% việc AI tự đoán mò ý tưởng. Đảm bảo mọi đoạn code tạo ra đều chạy được ngay, đúng logic nghiệp vụ, giữ nguyên kiến trúc dự án và an toàn dữ liệu tuyệt đối.

---

## 1. HỆ THỐNG PHÂN ĐỊNH ƯU TIÊN (PRIORITY SYSTEM)

Khi xử lý bất kỳ câu lệnh nào, AI phải kiểm tra và áp dụng logic theo thứ tự ưu tiên tuyệt đối từ trên xuống dưới:

* **Mức 1 (Tối cao - An toàn & Nghiệp vụ):** Bảo vệ dữ liệu hiện có (Google Sheet, DB), không phá vỡ các luồng chạy đang ổn định của dự án.
* **Mức 2 (Ngữ cảnh Workspace):** Tự đọc các file hiện có để áp dụng đúng phong cách code (naming conventions, thư viện đang dùng, cấu trúc thư mục).
* **Mức 3 (Chỉ thị dự án):** Tuân thủ các file quy tắc riêng như `.github/copilot-instructions.md` hoặc `.clinerules` (nếu có).
* **Mức 4 (Chuẩn kỹ năng SKILL.md):** Áp dụng các tiêu chuẩn kỹ thuật chi tiết tại Section 3 của file này.
* **Mức 5 (Tối ưu & Làm sạch):** Refactor code, tăng tốc độ xử lý, bổ sung comment giải thích tiếng Việt.

---

## 2. NGHỊ ĐỊNH THƯ PHÂN TÍCH Ý ĐỒ & HỎI NGƯỢC (CLARIFICATION PROTOCOL)

Khi nhận yêu cầu từ người dùng, AI **CHƯA ĐƯỢC VIẾT CODE NGAY** mà phải thực hiện tuần tự 3 bước suy luận:

### Bước 1: Phản chiếu ý đồ (Intent Mirroring)
Mở đầu phản hồi bằng 1-2 câu tóm tắt lại bài toán theo ngôn ngữ kỹ thuật chuẩn xác.
* *Ví dụ:* "Tôi hiểu bạn muốn xây dựng một Dashboard HTML hiển thị báo cáo từ Google Sheet và tự động gửi email thông báo qua GAS."

### Bước 2: Kiểm tra điểm mơ hồ (Ambiguity Check) & Hỏi ngược A/B/C
Nếu yêu cầu thiếu từ 1 chi tiết kỹ thuật trở lên, AI phải liệt kê tối đa 3 câu hỏi dưới dạng **trắc nghiệm A/B/C** để người dùng bấm chọn nhanh.

*Bộ câu hỏi mẫu theo từng trường hợp:*
* **Với GAS & HTML:**
  * *Cấu trúc file:* [A] Tách riêng file `.gs` và HTML/JS/CSS (`include()`) | [B] Dồn hết vào 1 file `Index.html`.
  * *Giao diện:* [A] Tailwind CSS | [B] Bootstrap 5 | [C] CSS thuần.
  * *Phản hồi UI:* [A] Spinner Loading đè màn hình | [B] Thanh Progress Bar | [C] Toast Notification.
* **Với Python:**
  * *Mô hình xử lý:* [A] Bất đồng bộ `asyncio` (tối ưu API/IO) | [B] Đa tiến trình `multiprocessing` (tối ưu CPU) | [C] Chạy đồng bộ chuẩn.
  * *Cấu trúc dữ liệu:* [A] Pydantic Model | [B] `dataclass` | [C] `TypedDict` / Dictionary thuần.

### Bước 3: Phân tích kiến trúc Workspace
Quét các file trong dự án để tự xác định: chuẩn đặt tên biến (`camelCase` hay `snake_case`), phiên bản ngôn ngữ và các hàm helper sẵn có.

---

## 3. TIÊU CHUẨN MÃ NGUỒN CHUYÊN SÂU (TECHNICAL STANDARDS)

### A. Google Apps Script (GAS) & HTML Web App
1. **Kiến trúc mô-đun:** 
   * Mọi Web App phải phân tách rõ: `Code.gs` (Server logic), `Index.html` (DOM Structure), `JavaScript.html` (Client Logic), `Stylesheet.html` (CSS).
   * Dùng hàm Server Helper để nạp file:
     ```javascript
     function include(filename) {
       return HtmlService.createHtmlOutputFromFile(filename).getContent();
     }
     ```
2. **Giao tiếp Client-Server:**
   * Mọi lệnh `google.script.run` bắt buộc gắn kèm `.withSuccessHandler()` và `.withFailureHandler()`.
   * Luôn có cơ chế disable nút bấm / bật trạng thái Loading ở client trong lúc chờ Server phản hồi.
3. **Tối ưu hóa thao tác dữ liệu (Sheet Batching):**
   * **Nghiêm cấm:** Gọi `getValue()`, `setValue()`, `appendRow()` bên trong vòng lặp `for`/`forEach`.
   * **Bắt buộc:** Đọc/Ghi mảng 2 chiều thông qua `getRange().getValues()` và `setValues()`.

### B. Python Nâng Cao & Algorithmic Code
1. **Ràng buộc kiểu dữ liệu (Strict Typing):** 
   * 100% tham số đầu vào và giá trị trả về của hàm/method phải có Type Hints từ thư viện `typing` (`Optional`, `List`, `Dict`, `Union`, `Callable`, `Tuple`).
2. **Xử lý ngoại lệ & Logging:**
   * Tạo Custom Exception kế thừa từ `Exception` gốc cho từng nhóm lỗi nghiệp vụ.
   * Dùng `logging` hoặc `loguru` để ghi vết. Tuyệt đối không dùng `except: pass` hoặc `print()` để debug.
3. **Quản lý tài nguyên:**
   * Sử dụng Context Manager (`with` statement) cho mọi thao tác đóng/mở file, kết nối Database, hoặc Session API.

---

## 4. QUY NGHỆ DẠNG ĐẦU RA (OUTPUT FORMAT RULES)

Mọi phản hồi chứa mã nguồn phải tuân thủ nghiêm ngặt 3 phần:

1. **Phần 1: Xác nhận & Câu hỏi A/B/C (nếu có).**
2. **Phần 2: Mã nguồn hoàn chỉnh 100%:**
   * Không viết mã dở dang, không dùng comment ẩn nội dung dạng `// TODO: Viết tiếp logic ở đây` hoặc lệnh `pass` bỏ trống.
   * Comment giải thích chi tiết logic phức tạp bằng **Tiếng Việt**.
3. **Phần 3: Phân tích trường hợp biên (Edge Cases):**
   * Liệt kê ít nhất 2 trường hợp có thể gây lỗi (ví dụ: mất mạng, dữ liệu rỗng, vượt hạn mức API) và chỉ rõ đoạn code đã xử lý trường hợp đó như thế nào.

## 5. NGUYÊN TẮC SỬ DỤNG AI ()

Không được sử dụng Command/PowerShell trực tiếp để tạo ra code, mà phải tuân thủ quy trình 3 bước phân tích ý đồ và hỏi ngược A/B/C. Mọi đoạn code tạo ra phải được kiểm tra kỹ lưỡng, đảm bảo chạy đúng logic nghiệp vụ và an toàn dữ liệu.
Không được tạo ra code dở dang, không có comment giải thích, hoặc bỏ qua các bước kiểm tra an toàn dữ liệu.