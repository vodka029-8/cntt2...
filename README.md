# cntt2...
Dưới đây là lời giải chi tiết và hướng dẫn các bước thực hiện cho **Bài 3: Chuẩn hóa quy trình giao vận và tra cứu giá cước** thuộc phần Bài tập về nhà (BTVN) Buổi 12 [Nhập Môn Công Nghệ Thông Tin](https://lms-student.rikkei.edu.vn/courses/6a88fb77a493e63f11a75432) trên hệ thống LMS:

---

1. **Phần 1: Phân tích (Mapping VLOOKUP với Nhân viên kho):**
Trong mô hình nghiệp vụ của công ty **FastMove**, hành động tra cứu giá cước thủ công gây chậm trễ 16 tiếng mỗi ngày được tự động hóa thông qua hàm `VLOOKUP` với 4 tham số tương đương với thao tác thực tế của một nhân viên kho:

* **Lookup_value (Giá trị cần dò tìm):** Tương ứng với **Mã vận đơn** trên kiện hàng mà nhân viên kho đang cầm trên tay.
* **Table_array (Vảng bảng dò tìm):** Tương ứng với **Quyển sổ tra cứu giá cước** lưu trữ toàn bộ danh mục mã vận đơn và mức phí tương ứng.
* **Col_index_num (Số thứ tự cột cần lấy dữ liệu):** Tương ứng với việc nhân viên kho dò ngang qua dòng của mã vận đơn để chỉ đọc đúng **cột "Giá cước"** cần lấy.
* **Range_lookup (Kiểu khớp dữ liệu - chọn 0 hoặc FALSE):** Tương ứng với quy tắc tìm kiếm **chính xác tuyệt đối** (phải trùng khớp hoàn toàn mã vận đơn, tuyệt đối không được ước lượng hay lấy gần đúng).


2. **Phần 2: Triển khai (Giải pháp công nghệ cho FastMove):**
* **Microsoft Word (Quy trình vận hành kho):** Soạn thảo tài liệu chuẩn, áp dụng hệ thống Heading kết hợp Styles đồng bộ theo đúng màu sắc thương hiệu của FastMove. Thiết lập mục lục tự động (Auto TOC) để dễ dàng kiểm soát nội dung tài liệu hướng dẫn shipper.
* **Microsoft Excel (Tự động hóa đối soát giá cước):**
* Lập công thức `VLOOKUP` để nối bảng "Mã vận đơn" với "Giá cước" từ 2 Sheet khác nhau.
* Sử dụng ký hiệu Dollar sign (`$`) để **cố định tuyệt đối vùng bảng tra cứu** (ví dụ: `\$Sheet2!\$A\$2:\$B\$500`), giúp khi kéo công thức xuống 500 dòng mà vùng tham chiếu không bị lệch.
* Đảm bảo tuân thủ 5 Quy tắc Vàng: **Tuyệt đối không gộp ô (Merge Cell)** trong vùng dữ liệu và không để trống hàng giữa các bản ghi.


* **Microsoft PowerPoint (Slide hướng dẫn shipper):**
* Áp dụng **Miller's Law** (giới hạn tối đa 3-4 bullet points mỗi slide) để tối ưu hóa lượng thông tin tiếp thu cho shipper.
* Định dạng tiêu đề slide theo dạng kết luận (Assertion Title) thay vì dùng tên chủ đề chung chung.




3. **Phần 3: Kiểm thử và Xử lý lỗi #N/A (Dùng hàm IFERROR):**
* **Nguyên nhân:** Khi thực hiện hàm `VLOOKUP` đối với các mã vận đơn có chứa ký tự lạ, khoảng trắng ẩn hoặc mã không tồn tại trong bảng tra cứu, Excel sẽ trả về lỗi `#N/A`.
* **Cách khắc phục bằng hàm IFERROR:**
1. Bọc hàm `VLOOKUP` bên trong hàm `IFERROR` để hệ thống tự động bẫy lỗi và trả về giá trị thay thế thân thiện hơn thay vì hiển thị lỗi `#N/A` làm hỏng giao diện báo cáo.
2. Cấu trúc công thức chuẩn: `=IFERROR(VLOOKUP(A2, Sheet2!$A$2:$B$500, 2, FALSE), "Không tìm thấy mã cước")`.
3. Kiểm tra và làm sạch dữ liệu đầu vào bằng hàm `TRIM` nếu phát hiện lỗi do khoảng trắng thừa ở mã vận đơn.
