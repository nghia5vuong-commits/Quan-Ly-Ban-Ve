# AGENTS.md

## Quy tắc bắt buộc cho repo này

- Luôn áp dụng skill `dev-master-python-gas-html` cho mọi task trong repo, dù task nhỏ nhất như sửa tên biến, sửa text, thêm button, tối ưu hàm, hoặc chỉnh UI.
- Không được bỏ qua quy trình skill: phản chiếu ý đồ, kiểm tra điểm mơ hồ, hỏi ngược A/B/C nếu thiếu thông tin, rồi mới viết code.
- Không được viết code trực tiếp mà không đọc ngữ cảnh workspace trước.
- Không được tự đoán mò kiến trúc, tên biến, hoặc API; phải dựa trên file hiện có trong repo.
- Mọi code JavaScript/GAS/HTML/Python phải tuân thủ đúng chuẩn repo: camelCase, arrow function, comment tiếng Việt, bảo toàn kiến trúc dự án.
- Mọi thao tác client-server với Google Apps Script phải dùng `withSuccessHandler()` và `withFailureHandler()`.
- Mọi thao tác dữ liệu trên Google Sheet không được dùng `getValue()`, `setValue()`, `appendRow()` trong vòng lặp; phải batch bằng mảng dữ liệu.
- Mọi task phải ưu tiên an toàn dữ liệu và không phá luồng đang ổn định.
- Nếu thông tin chưa đủ, phải hỏi ngược A/B/C trước khi sửa code.
- Khi có thể, phải kiểm tra bằng cách chạy kiểm thử hoặc lệnh liên quan trước khi kết luận fix đã thành công.

## Mức ưu tiên bắt buộc

1. An toàn dữ liệu và nghiệp vụ.
2. Đọc ngữ cảnh workspace.
3. Tuân thủ repo instruction và skill.
4. Viết code đúng chuẩn kỹ thuật.
5. Kiểm tra xác thực với evidence.

## Không được làm

- Không bỏ qua skill vì “task nhỏ”.
- Không dùng cách làm nhanh thiếu kiểm tra.
- Không viết code dở dang, không có comment tiếng Việt, không có kiểm tra edge case.
- Không dùng `pass`, `TODO` thay cho logic hoàn chỉnh.
