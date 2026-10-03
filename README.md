# OrderBridge DN2

Ứng dụng cục bộ tự động chuyển báo cáo bán hàng Sopoka thành workbook import KiotViet và hoàn tất từng đơn trên chi nhánh **Đồng Nai 2 - ĐN2**. Mỗi workbook tối đa 19 sheet/đơn hàng.

## Dùng hằng ngày

1. Mở `start_automation.cmd`, sau đó vào `http://127.0.0.1:8765`.
2. Đăng nhập Sopoka và KiotViet một lần trong cửa sổ Microsoft Edge do tool mở nếu phiên đăng nhập đã hết hạn.
3. Trên dashboard, bấm **Tải report & chạy tự động**.

Tool sẽ thực hiện tuần tự:

- Vào **Báo cáo → Báo cáo bán hàng → Tổng hợp bán hàng**, chọn **Hôm nay**, tìm kiếm và xuất Excel. Tool ưu tiên tải file report đã xác thực trực tiếp từ Office Viewer bằng phiên Edge đang đăng nhập; nếu Edge hiển thị được nút native thì **Download file** vẫn là phương án dự phòng.
- Chỉ nhận báo cáo một ngày. Ngày hiện tại dùng mặc định; nếu chọn ngày cũ trên dashboard, tool giữ đúng ngày đã chọn qua 00:00 để có thể chạy lại an toàn. Report nhiều ngày vẫn bị chặn.
- Khi chạy thủ công, ô đường dẫn nhận file Excel `.xls`, `.xlsx`, `.xlsm`, `.xltx` và `.xltm`; file OpenXML có lỗi định dạng màu vẫn được đọc theo dữ liệu ô. `.xlsb` chưa được hỗ trợ.
- File nguồn nhập hàng Đồng Nai 1 cũng có thể là các định dạng Excel trên; nguồn chỉ đọc, dòng có số lượng giao hàng trống/dấu `-` vẫn được bỏ qua và giá nhập được chia 1,08 trước khi tạo workbook.
- Đối chiếu SQLite theo `ngày + Mã đặt đơn hàng liên quan + dữ liệu đơn`. Mã đã hoàn tất trong ngày được cảnh báo và không tạo lại nếu không có quyết định tiếp tục của người dùng.
- Tạo file import tối đa 19 sheet mỗi lô; lô sau chỉ được upload sau khi lô trước đã được xác minh và ghi SQLite thành công.
- Mở đúng **Đồng Nai 2 - ĐN2**. Nếu KiotViet đang ở tab **Hóa đơn**, tool chuyển sang **Đặt hàng**; nếu không xác minh được thì dừng an toàn.
- Upload từng file, điền **Ghi chú đơn hàng** bằng đúng mã tương ứng với từng sheet/tab, kiểm tra lại ghi chú và tab trước khi tự bấm **ĐẶT HÀNG**.
- Khi cửa sổ in PDF của Edge xuất hiện, tool bấm **Cancel** rồi chuyển sang đơn tiếp theo.

Dashboard hiển thị ngày nghiệp vụ theo Asia/Bangkok, tiến độ theo report, lô và từng đơn. Khi có nhiều file import, thanh tổng không reset; UI nêu rõ `Lô X/Y` và thanh riêng của từng lô. Bạn không cần cài hoặc bật extension Chrome/Edge.

## Khi tool dừng an toàn

Nếu Edge không xác minh được chi nhánh, tab, ghi chú, hoặc kết quả của một lần bấm **ĐẶT HÀNG**, lô được chuyển sang `automation_attention`. Tool không tự bấm lại mù quáng và không upload lô kế tiếp, nhằm tránh tạo đơn trùng. Dọn các tab dở trên KiotViet, kiểm tra lại, rồi dùng nút **Đã dọn tab, thử lại** trên dashboard.

## Dữ liệu và đối soát

- SQLite: `data/automation.sqlite3`
- Backup trước lần tự bấm đơn: `data/backups/`
- Workbook đã tạo: `outputs/<ngày>/<run>/`
- Log Edge: `logs/edge-automation.log`

SQLite chỉ ghi `automation_verified_submission` sau khi Edge đã xác minh đủ toàn bộ tab/mã của lô. Sang ngày mới, cùng một mã có thể được xử lý lại theo quy tắc ngày kinh doanh mới.

## Điều kiện máy

- Microsoft Excel phải được cài để tạo workbook theo mẫu KiotViet.
- `msedgedriver.exe` phải nằm tại `C:\Users\binni\OneDrive\Desktop\msedgedriver.exe`.
- Nếu Sopoka hoặc KiotViet yêu cầu đăng nhập lại, đăng nhập ngay trong cửa sổ Edge do tool điều khiển, rồi bấm lại thao tác trên dashboard.

Tool không tự đăng nhập hoặc đọc Gmail. Nếu cần dùng file gửi qua email, hãy tải file đính kèm về máy rồi dán đường dẫn vào ô report/nhập hàng; việc này giữ thông tin tài khoản ngoài phạm vi của ứng dụng.
