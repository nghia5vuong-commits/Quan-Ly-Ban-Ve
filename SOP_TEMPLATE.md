# SOP - QUẢN LÝ BẢN VẼ & SẢN XUẤT

**Standard Operating Procedure**

---

## Trang Bìa (Cover)

**LIXIL LOGO**

**Phiên bản:** 1.0  
**Ngày phê duyệt:** _______________  
**Người phê duyệt:** _______________  
**Bộ phận:** Kỹ Thuật & Sản Xuất

---

## 📋 MỤC LỤC

1. **Giới Thiệu Hệ Thống** - Trang 3
2. **Đăng Nhập & Giao Diện Chính** - Trang 4
3. **Thêm Bản Vẽ (Tab 2)** - Trang 5-6
4. **Quản Lý Bản Vẽ (Tab 3)** - Trang 7
5. **Upload Thiết Kế (Upload Bản Vẽ)** - Trang 8
6. **Tìm Kiếm Bản Vẽ (AI Search)** - Trang 9
7. **Phê Duyệt Bản Vẽ (Approval)** - Trang 10-11
8. **Quản Lý Phát Hành (Release)** - Trang 12
9. **Thống Kê Tần Suất TO (History)** - Trang 13
10. **Xử Lý Sự Cố Thường Gặp** - Trang 14

---

# 1️⃣ GIỚI THIỆU HỆ THỐNG

## Mục Đích Hệ Thống

Hệ thống Quản Lý Bản Vẽ & Sản Xuất được phát triển để:

- Tiếp nhận, lưu trữ và quản lý tất cả bản vẽ kỹ thuật
- Tự động phân loại bản vẽ theo loại, TO, khách hàng
- Quản lý quy trình phê duyệt (QA → QC → Manager → Phát hành)
- Thống kê tần suất sử dụng của mỗi TO (Technical Outline)
- Cung cấp các báo cáo phân tích nhanh chóng

## Thành Phần Chính Của Hệ Thống

| Module | Chức Năng | Người Dùng |
|--------|---------|-----------|
| **Thêm Bản Vẽ** | Tiếp nhận bản vẽ mới, upload hình ảnh bản vẽ | Design, Checker, QA |
| **Quản Lý Bản Vẽ** | Xem, chỉnh sửa thông tin, lịch sử bản vẽ | Tất cả người dùng |
| **Upload Thiết Kế** | Upload file bản vẽ, CAD files | Design, Checker |
| **Phê Duyệt Bản Vẽ** | QA, QC, Manager ký duyệt bản vẽ | QA, QC, Manager |
| **Quản Lý Phát Hành** | Ghi nhận phát hành, thống kê SO | QA, Manager |

📷 **Dán hình ảnh Dashboard chính tại đây**

---

# 2️⃣ ĐĂNG NHẬP & GIAO DIỆN CHÍNH

## Bước 1: Mở Ứng Dụng

> **1** Sử dụng tài khoản Google của bạn để truy cập URL ứng dụng

## Bước 2: Đăng Nhập

> **2** Nhập email Google và mật khẩu (nếu có dialog yêu cầu)

## Bước 3: Cho Phép Quyền Truy Cập

> **3** Nhấn "Cho phép" để ứng dụng truy cập Google Drive & Sheet

📷 **Dán hình ảnh màn hình đăng nhập**

## Giao Diện Chính (Sau Khi Đăng Nhập)

- **Sidebar Trái:** Menu điều hướng (Trang chủ, Thêm Bản Vẽ, Quản Lý BV, v.v.)
- **Topbar Trên:** Thông tin người dùng, tên bộ phận, thông báo
- **Vùng Chính:** Hiển thị nội dung theo menu được chọn
- **Thanh Cuộn:** Cuộn để xem thêm nội dung nếu cần

📷 **Dán hình ảnh giao diện chính (dashboard)**

---

# 3️⃣ THÊM BẢN VẼ (Tab 2)

## Khi Nào Sử Dụng Tab Này?

Tab "Thêm Bản Vẽ" được sử dụng khi bạn muốn **tiếp nhận bản vẽ mới hoặc thêm thông tin bản vẽ mà chưa có hình ảnh**.

## Bước 1: Chuyển Sang Tab "Thêm Bản Vẽ"

> **1** Nhấp vào menu "Quản Lí Bản Vẽ" (Sidebar Trái)

> **2** Chọn tab thứ 2 có nhãn "Thêm Bản Vẽ"

📷 **Dán hình ảnh tab "Thêm Bản Vẽ"**

## Bước 2: Điền Thông Tin Bản Vẽ

| Trường | Mô Tả | Yêu Cầu |
|--------|-------|---------|
| **Loại Bản Vẽ** | Chọn loại: Bản vẽ kỹ thuật, CAD, 3D, v.v. | Bắt buộc |
| **TO (Technical Outline)** | Mã TO của bản vẽ (VD: 7Y090A) | Bắt buộc |
| **Khách Hàng** | Tên khách hàng yêu cầu bản vẽ | Bắt buộc |
| **Mã Bản Vẽ** | Mã số bản vẽ (VD: DW-2024-001) | Tuỳ chọn |
| **Mô Tả** | Mô tả ngắn về bản vẽ | Tuỳ chọn |

## Bước 3: Nhấn Nút "Thêm Bản Vẽ"

> **3** Sau khi điền xong thông tin, nhấn nút **"Thêm Bản Vẽ"** để lưu

⚠️ **Ghi Chú:** Nếu TO + Khách Hàng đã tồn tại trong hệ thống, bản vẽ sẽ được cập nhật vào sheet "SO" (Số SO), không thêm mới vào sheet "Data".

📷 **Dán hình ảnh form "Thêm Bản Vẽ"**

---

# 3️⃣ THÊM BẢN VẼ (Tiếp Theo)

## Bước 4: Xác Nhận Thông Tin

Sau khi nhấn "Thêm Bản Vẽ", hệ thống sẽ:

- ✅ Kiểm tra TO + Khách Hàng có tồn tại chưa
- ✅ Nếu mới: Lưu vào sheet "Data" và "SO" cùng lúc
- ✅ Nếu đã tồn tại: Chỉ cập nhật tần suất trong sheet "SO"
- ✅ Hiển thị thông báo kết quả (Thành công / Lỗi)

## Bước 5: Upload Hình Ảnh Bản Vẽ

Sau khi thêm bản vẽ thành công, bạn có thể upload hình ảnh:

> **1** Nhấp vào nút **"Chọn Hình"** hoặc kéo thả hình vào vùng

> **2** Chọn tệp hình ảnh từ máy tính (JPG, PNG, PDF)

> **3** Nhập chiều rộng (Width - W) và chiều cao (Height - H) của hình

> **4** Nhấn **"Lưu Hình Ảnh"** để tải lên

📷 **Dán hình ảnh upload hình ảnh bản vẽ**

## Bước 6: Kiểm Tra Kết Quả

Sau khi upload thành công, bạn có thể:

- Xem lại hình ảnh vừa upload
- Chỉnh sửa thông tin nếu cần (W, H)
- Hình ảnh sẽ được lưu vào sheet "Data" column AG

✨ **Kết Quả:** Bản vẽ sẽ được chuyển sang Tab 3 "Quản Lí Bản Vẽ" khi có hình ảnh + W + H hoàn chỉnh.

---

# 4️⃣ QUẢN LÝ BẢN VẼ (Tab 3)

## Khi Nào Sử Dụng Tab Này?

Tab "Quản Lí Bản Vẽ" hiển thị **tất cả bản vẽ đã hoàn chỉnh** (có hình ảnh + W + H). Tại đây bạn có thể xem, chỉnh sửa, và quản lý bản vẽ.

## Bước 1: Xem Danh Sách Bản Vẽ

> **1** Truy cập Tab 3 "Quản Lí Bản Vẽ"

> **2** Xem danh sách tất cả bản vẽ trong bảng

> **3** Tìm kiếm bản vẽ bằng cách nhập TO, Khách Hàng, hoặc Mã Bản Vẽ

📷 **Dán hình ảnh tab "Quản Lí Bản Vẽ"**

## Bước 2: Chỉnh Sửa Bản Vẽ

- Nhấp vào nút **"Sửa"** trong hàng bản vẽ
- Cửa sổ popup sẽ mở ra với thông tin chi tiết
- Chỉnh sửa các trường theo nhu cầu
- Nhấn **"Lưu"** để cập nhật thay đổi

## Bước 3: Xóa Bản Vẽ

⚠️ **Cảnh Báo:** Để xóa bản vẽ, bạn cần quyền Admin. Hãy liên hệ quản trị viên hệ thống.

📷 **Dán hình ảnh form sửa bản vẽ**

---

# 5️⃣ UPLOAD THIẾT KẾ (Upload Bản Vẽ)

## Khi Nào Sử Dụng?

Sử dụng menu "Upload Bản Vẽ" để upload các file thiết kế (CAD, PDF, DWG) của bản vẽ đã được tiếp nhận.

## Bước 1: Truy Cập Menu Upload

> **1** Từ Sidebar Trái, nhấp vào **"Upload Bản Vẽ"**

> **2** Giao diện trang upload sẽ hiển thị

📷 **Dán hình ảnh giao diện upload bản vẽ**

## Bước 2: Chọn Bản Vẽ & Upload File

| Bước | Hành Động |
|------|----------|
| 1 | Chọn loại bản vẽ từ dropdown |
| 2 | Chọn TO từ danh sách (hoặc nhập nếu là TO mới) |
| 3 | Chọn Khách Hàng |
| 4 | Kéo thả hoặc nhấp để chọn file thiết kế |
| 5 | Nhấn **"Upload"** để gửi file |

## Bước 3: Xác Nhận Upload

- ✅ Hệ thống sẽ hiển thị thông báo "Upload thành công"
- ✅ File sẽ được lưu vào Google Drive của dự án
- ✅ Link file sẽ được ghi vào sheet "Data" column tương ứng

📷 **Dán hình ảnh kết quả upload thành công**

---

# 6️⃣ TÌM KIẾM BẢN VẼ (AI Search)

## Khi Nào Sử Dụng?

Menu "Tìm Kiếm Bản Vẽ" cho phép bạn **tìm kiếm bản vẽ nhanh chóng theo TO, Khách Hàng, hoặc các tiêu chí khác**.

## Bước 1: Mở Giao Diện Tìm Kiếm

> **1** Từ Sidebar, nhấp vào **"Tìm Kiếm Bản Vẽ"**

> **2** Giao diện tìm kiếm sẽ hiển thị với các bộ lọc

## Bước 2: Nhập Tiêu Chí Tìm Kiếm

- **TO:** Nhập mã TO (VD: 7Y090A)
- **Khách Hàng:** Chọn hoặc nhập tên khách
- **Loại Bản Vẽ:** Chọn loại (Kỹ Thuật, CAD, 3D)
- **Trạng Thái:** Chọn trạng thái (Mới, Đã Phê Duyệt, v.v.)

## Bước 3: Thực Hiện Tìm Kiếm

> **1** Nhấn nút **"Tìm Kiếm"** hoặc **"Áp Dụng Lọc"**

> **2** Kết quả sẽ hiển thị trong bảng dưới

> **3** Nhấp vào bản vẽ để xem chi tiết

📷 **Dán hình ảnh giao diện tìm kiếm bản vẽ**

## Bước 4: Xuất Kết Quả (Tuỳ Chọn)

- Nhấp vào nút **"Xuất Excel"** để tải kết quả về máy
- Hoặc nhấp **"In"** để in danh sách tìm kiếm

---

# 7️⃣ PHÊ DUYỆT BẢN VẼ (Approval)

## Quy Trình Phê Duyệt 4 Cấp

Mỗi bản vẽ phải qua 4 bước phê duyệt trước khi được phát hành:

- **Cấp 1 (QA):** Kỹ sư QA kiểm tra bản vẽ
- **Cấp 2 (QC):** Nhân viên QC xác nhận chất lượng
- **Cấp 3 (Manager):** Quản lý phê duyệt chính thức
- **Cấp 4 (Release):** Phát hành và ghi nhận SO

## Bước 1: Truy Cập Trang Phê Duyệt

> **1** Từ Sidebar, nhấp vào **"Phê Duyệt Bản Vẽ"**

> **2** Danh sách bản vẽ chờ phê duyệt sẽ hiển thị

## Bước 2: Chọn Bản Vẽ & Phê Duyệt

> **1** Nhấp vào nút **"Ký Duyệt"** trong hàng bản vẽ

> **2** Cửa sổ popup sẽ hiển thị chi tiết bản vẽ

> **3** Xem lại hình ảnh, thông tin TO, Khách Hàng

> **4** Chọn cấp phê duyệt của bạn từ dropdown

> **5** Nhập ghi chú (nếu cần) và nhấn **"Ký Duyệt"**

📷 **Dán hình ảnh giao diện phê duyệt bản vẽ**

## Bước 3: Phê Duyệt Hàng Loạt (Bulk Approve)

Để phê duyệt nhiều bản vẽ cùng lúc:

> **1** Nhấp vào checkbox bên trái để chọn từng bản vẽ

> **2** Hoặc nhấp **"Chọn Tất Cả"** để chọn toàn bộ

> **3** Nhấn nút **"Ký Hàng Loạt"** (sẽ hiện khi có checkbox được chọn)

> **4** Chọn cấp phê duyệt và xác nhận

---

# 7️⃣ PHÊ DUYỆT BẢN VẼ (Tiếp Theo)

## Quyền Hạn Phê Duyệt Theo Chức Vụ

| Chức Vụ | Cấp Phê Duyệt | Quyền |
|---------|--------------|------|
| QA | Cấp 1 | Ký duyệt cấp QA, có thể phê duyệt hàng loạt |
| QC | Cấp 2 | Ký duyệt cấp QC, có thể phê duyệt hàng loạt |
| Manager | Cấp 3 | Ký duyệt cấp Manager, có thể phê duyệt hàng loạt |
| Charge | N/A | Chỉ có thể ký từng bản vẽ một, không được ký hàng loạt |

⚠️ **Ghi Chú Charge User:** Nếu bạn là Charge User, bạn sẽ không thấy nút "Ký Hàng Loạt" nhưng vẫn có thể chọn checkbox để ký từng bản vẽ một.

## Bước 4: Xác Nhận & Hoàn Tất

- ✅ Hệ thống sẽ hiển thị thông báo "Ký duyệt thành công"
- ✅ Bản vẽ sẽ được đánh dấu là "Đã ký [Cấp X]"
- ✅ Email thông báo sẽ được gửi cho người liên quan
- ✅ Bản vẽ sẽ tự động ẩn khỏi danh sách phê duyệt sau khi ký

## Bước 5: Kiểm Tra Lịch Sử Phê Duyệt

Để xem ai đã phê duyệt bản vẽ:

> **1** Truy cập Tab 3 "Quản Lí Bản Vẽ"

> **2** Tìm bản vẽ muốn xem lịch sử

> **3** Nhấp vào nút **"Sửa"** để xem chi tiết (column: Checker, Approver)

📷 **Dán hình ảnh chi tiết bản vẽ sau phê duyệt**

---

# 8️⃣ QUẢN LÝ PHÁT HÀNH (Release)

## Mục Đích Của Trang Release

Trang "Quản Lý Phát Hành" giúp bạn:

- 📊 Xem thống kê tất cả bản vẽ đã phát hành
- 📈 Theo dõi số lượng phát hành theo TO, Khách Hàng
- ✏️ Ghi nhận SO (Sales Order) khi phát hành
- 📝 Cập nhật trạng thái phát hành của bản vẽ

## Bước 1: Truy Cập Trang Release

> **1** Từ Sidebar, nhấp vào **"Quản Lý Phát Hành"**

> **2** Giao diện release dashboard sẽ hiển thị

## Bước 2: Xem Thống Kê KPI

Phía trên sẽ hiển thị các số liệu:

- **Tổng Bản Vẽ:** Tổng số bản vẽ trong hệ thống
- **Đã Phát Hành:** Số bản vẽ đã được phát hành
- **Chưa Phát Hành:** Số bản vẽ chưa phát hành
- **TO Hoạt Động:** Số TO đang được sử dụng
- **Khách Hàng Hoạt Động:** Số khách hàng đang làm việc

📷 **Dán hình ảnh dashboard phát hành**

## Bước 3: Ghi Nhận Phát Hành

> **1** Tìm bản vẽ trong bảng phát hành

> **2** Nhập Số SO (Sales Order) vào cột tương ứng

> **3** Nhấn **"Lưu"** để ghi nhận phát hành

## Bước 4: Xuất Báo Cáo Phát Hành

- Nhấp vào nút **"Xuất Excel"** để tải báo cáo phát hành
- Hoặc nhấp **"In"** để in báo cáo

---

# 9️⃣ THỐNG KÊ TẦN SUẤT TO (History & Analytics)

## Mục Đích Của Trang History

Trang "Lịch Sử & Thống Kê" cung cấp các báo cáo chi tiết:

- 📊 Tần suất sử dụng của mỗi TO
- 🎨 Số lần xuất hiện của mỗi màu cho TO
- 👥 Số khách hàng sử dụng TO
- 📈 So sánh dữ liệu SO giữa các phiên bản bản vẽ

## Bước 1: Truy Cập Trang History

> **1** Từ Sidebar, nhấp vào **"Lịch Sử & Thống Kê"**

> **2** Chọn Tab 1 **"Tần Suất TO"**

## Bước 2: Xem Thống Kê Tần Suất TO (Tab 1)

Phía trên sẽ hiển thị:

- **Tổng TO:** Số lượng TO độc nhất
- **TO Cao Nhất:** TO được sử dụng nhiều nhất
- **Tần Suất Cao Nhất:** Lần sử dụng tối đa
- **Trung Bình Tần Suất:** Bình quân lần sử dụng

## Bước 3: Xem Bảng Tần Suất Chi Tiết

| Cột | Ý Nghĩa |
|-----|---------|
| **TO** | Mã TO (VD: 7Y090A) |
| **Tần Suất** | Lần được sử dụng cao nhất |
| **Số Lần Xuất Hiện** | Tổng lần TO xuất hiện |
| **Số Màu** | Số loại màu khác nhau của TO |
| **Số Khách** | Số khách hàng khác nhau sử dụng TO |
| **Các Màu** | Danh sách các màu sắc (VD: FJ, AB) |
| **Khách Hàng** | Danh sách khách hàng sử dụng TO |

📷 **Dán hình ảnh bảng tần suất TO**

## Bước 4: Xem Tab 2 "Đối Chiếu SO"

> **1** Chọn Tab 2 **"Đối Chiếu SO"**

> **2** Xem so sánh bản vẽ mới vs bản vẽ cũ

> **3** Sử dụng bộ lọc để tìm SO cụ thể

---

# 🔟 XỬ LÝ SỰ CỐ THƯỜNG GẶP

## Sự Cố 1: Bản Vẽ Không Hiển Thị Trong Tab 3 "Quản Lí Bản Vẽ"

**Nguyên Nhân:** Bản vẽ chưa có hình ảnh hoặc thiếu chiều W/H

**Giải Pháp:** Quay lại Tab 2, tìm bản vẽ và upload hình ảnh kèm W, H

## Sự Cố 2: Upload Hình Ảnh Thất Bại

**Nguyên Nhân:** Kích thước file quá lớn (>5MB) hoặc định dạng không hỗ trợ

**Giải Pháp:** Nén hình ảnh hoặc chuyển định dạng sang JPG/PNG (khuyến nghị <3MB)

## Sự Cố 3: Phê Duyệt Bản Vẽ Không Thành Công

**Nguyên Nhân:** Bạn không có quyền phê duyệt hoặc bản vẽ không ở trạng thái đúng

**Giải Pháp:** Kiểm tra chức vụ và quyền của mình, hoặc liên hệ Admin

## Sự Cố 4: Không Thấy Bản Vẽ Cần Tìm

**Nguyên Nhân:** Bản vẽ chưa được thêm vào hệ thống hoặc tên nhập sai

**Giải Pháp:** Kiểm tra lại tên TO/Khách Hàng hoặc thêm bản vẽ mới (Tab 2)

## Sự Cố 5: Không Thể Chỉnh Sửa Bản Vẽ

**Nguyên Nhân:** Bản vẽ đã bị khóa hoặc bạn không có quyền chỉnh sửa

**Giải Pháp:** Liên hệ Admin để mở khóa hoặc cấp quyền

## Liên Hệ Hỗ Trợ

📞 **Nếu gặp sự cố khác:** Liên hệ bộ phận Kỹ Thuật IT tại:

- Email: it-support@lixil.com
- Điện thoại: 0123-456-789
- Nội bộ: Phòng IT, Tầng 3

---

**Tài liệu này được tạo ngày:** 2026-09-14  
**Phiên bản:** 1.0  
**Bộ phận:** Kỹ Thuật & Sản Xuất
