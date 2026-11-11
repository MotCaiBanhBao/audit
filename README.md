# Device Audit Tool — bản cập nhật Windows nhập tay & lý do cần xác minh

## MỚI: Windows không lấy được bản quyền thì hỏi người khảo sát

Sau khi thử WMI và `slmgr /dli`, nếu trạng thái Windows vẫn chưa xác định thì màn hình khảo sát **tự động hỏi**:

- `[1] DA_KICH_HOAT`: chỉ chọn khi đã xem trực tiếp Activation trong Windows hoặc `slmgr /xpr`.
- `[2] CHUA_KICH_HOAT`: chỉ chọn khi đã kiểm tra trực tiếp và thấy chưa kích hoạt.
- `[3] GIU CAN_XAC_MINH`: Enter để bỏ qua nếu chưa kiểm tra được.

Màn hình sẽ hiện **mã trạng thái tự động, nguồn truy vấn, và thông báo lỗi thực tế** để biết vì sao bị thiếu kết quả. **Không hỏi lại** nếu kiểm tra tự động đã xác định chắc chắn `DA_KICH_HOAT` hoặc `CHUA_KICH_HOAT`.

Nguồn được ghi `XAC_NHAN_THU_CONG` nếu người khảo sát đã nhập tay; bản ghi gốc vẫn lưu các cột `Windows Auto State (trước nhập tay)`, `Windows Auto Method (trước nhập tay)` và `Windows Auto Note (trước nhập tay)`, cùng giờ nhập tay. **Nhập tay chỉ là ghi nhận trạng thái kích hoạt do người khảo sát kiểm tra; không phải chứng minh tính hợp pháp của giấy phép Windows.**

## MỚI: Giải thích các trạng thái cần xác minh

`Inventory.xlsx` có thêm cột **Lý do cần xác minh** ở ngay sau **Kích hoạt Windows**. CSV gốc có trường **Giải thích mục cần xác minh**. Nội dung ghi nhận giá trị thực tế và lý do cần đối chiếu cho Windows/Office, WMI, antivirus, Kaspersky, TPM, BitLocker, SMBv1, NTP, Firewall, HotFix, USB, Domain, Screen Lock, v.v.

Các trạng thái không đồng nghĩa:
- `CAN_XAC_MINH`: truy vấn chưa đủ căn cứ; chưa kết luận trạng thái thật.
- `TIMEOUT`: quá thời gian chờ; chưa đồng nghĩa thiết bị chưa kích hoạt hay thiếu phần mềm.
- `OFF` / `GUEST_ENABLED`: trạng thái hệ thống thực tế, có thể cần đánh giá chính sách; không phải lỗi thu thập.
- `XAC_NHAN_THU_CONG`: thông tin do người khảo sát nhập sau khi đối chiếu; không giả mạo kết quả WMI.


## Chạy khảo sát

1. Giải nén **nguyên thư mục `audit-main`** xuống máy/USB. Không chạy `.ps1` riêng lẻ.
2. Nhấp phải **`1_Run_Survey.bat` → Run as administrator**.
3. Nhập tầng, khoa/phòng, loại máy và người phụ trách như trước.
4. Công cụ tiếp tục tạo báo cáo `KhaoSat\Inventory.xlsx`, `Inventory.csv`, `Inventory_Detail.csv`, `DuLieuGoc\Device_*.csv`, danh sách ứng dụng và log.

## MỚI: Toàn bộ dữ liệu lệnh trả về (không chỉ dữ liệu lọc)

Trong **`KhaoSat\LenhGoc\Raw_<Tên máy>_<thời gian>_<mã lần chạy>\`** có:

- **`LenhGoc_TatCaLenh.txt`**: **một file TXT duy nhất** tổng hợp kết quả các lệnh đã được ghi nhận, tên lệnh, thời gian và lỗi/timeout (nếu có). Đọc nhanh bằng Notepad++/VS Code.
- `Main\`: các câu lệnh WMI phần cứng/Windows/mạng và Registry chạy trong phiên khảo sát chính.
- `WINDOWS\`, `OFFICE\`, `SECURITY\`, `POLICIES\`: các kết quả lệnh đã lấy trong những worker, kể cả dữ liệu đã ghi được trước khi worker bị timeout.
- Mỗi lệnh có file `.meta.txt` (lệnh, thời gian, trạng thái), `.txt` (đọc bằng mắt), `.clixml` (**đối tượng/thuộc tính đầy đủ** của PowerShell) hoặc `.stdout.txt` / `.stderr.txt` (đầu ra công cụ dòng lệnh). Tập tin được tạo từ kết quả trước khi lọc để điền phiếu.
- `WINDOWS.xml`, `OFFICE.xml`, `SECURITY.xml`, `POLICIES.xml` (nếu worker chạy xong) là bản kết quả tuần tự riêng, dùng để đối chiếu.

**Giới hạn thực tế:** không có cách lấy được dữ liệu chưa được hệ điều hành trả về. Những lệnh bị lỗi/timeout có nhật ký truy vấn nhưng có thể không có dữ liệu; đây không phải kết quả “trống”. CLIXML là dữ liệu đầy đủ *trong phạm vi thuộc tính mà lệnh đã trả về*, không phải bản sao toàn bộ ổ đĩa/Registry. Một số API hệ thống trả về kết quả suy diễn qua COM/method thay vì dữ liệu WMI trực tiếp; kết quả kiểm tra vẫn xuất trong các file worker và log.

Lưu ý: `2_Rebuild_Excel.bat` chỉ dựng lại báo cáo từ các CSV đã có; **không thể tạo lại output lệnh gốc từ một lần quét cũ**. Muốn có bằng chứng thô phải khảo sát lại bằng bản mới.

## MỚI: sửa timeout Windows

- Tách `WINDOWS` và `OFFICE` thành **hai worker riêng**. Windows hoàn tất là có `WINDOWS.xml` riêng, Office không thể làm mất kết quả Windows nữa.
- Windows ưu tiên `SoftwareLicensingProduct` (WMI) và đối chiếu `EditionID` từ Registry. Nếu WMI lỗi/timeout hoặc chưa đủ bằng chứng, thử **`cscript.exe //Nologo slmgr.vbs /dli`** (chỉ đọc, có giới hạn 12 giây).
- Nếu trạng thái `slmgr` được nhận diện một cách chắc chắn thì sử dụng kết quả đó. Nếu khác ngôn ngữ/không đủ bằng chứng, giữ **CẦN XÁC MINH**, không kết luận nhầm “Chưa kích hoạt”.
- Office có giới hạn 40 giây **riêng**, nếu quá hạn sẽ giữ kết quả Windows và chỉ xử lý Office thủ công.
- Có thể sửa trong `1_Run_Survey.bat`: `AUDIT_WINDOWS_TIMEOUT_SEC=30` (8–90 giây), `AUDIT_LICENSE_TIMEOUT_SEC=40` (10–120 giây). Không nên giảm quá thấp khi khảo sát máy cũ.
- Đánh giá **trạng thái kích hoạt** không có nghĩa xác nhận **tính hợp pháp của giấy phép**.

## Về an toàn dữ liệu

Chế độ ghi lệnh gốc có thể chứa **tên tài khoản, địa chỉ IP/MAC, số serial, mã định danh giấy phép, danh sách ứng dụng và cấu hình bảo mật**. Chỉ lưu trữ trên USB/thiết bị được quản lý, phân quyền cho người được phép; không gửi nguyên bộ lên nơi công khai. Không dùng các lệnh thay đổi giấy phép (`slmgr /ato`, `/ipk`, `/rearm`); script chỉ chạy truy vấn.

## Kiểm tra phiên bản

- Hệ thống mục tiêu: Windows PowerShell 2.0–5.1, Windows 7–11; các cmdlet chỉ tồn tại ở Windows mới đã được kiểm tra trước khi gọi.
- **Chưa thể kiểm thử trực tiếp trên Windows tại môi trường xuất bản.** Nên kiểm thử trước trên một máy Windows 10/11 đã biết trạng thái kích hoạt, một máy chưa kích hoạt (nếu có) và một máy dùng Office. So sánh `WINDOWS.xml`, `Windows kích hoạt` trong `Inventory.csv` và `LenhGoc_TatCaLenh.txt` với `slmgr /dli` chạy trên máy.

## Phân biệt các trạng thái

- `DA_KICH_HOAT`: xác nhận kích hoạt từ WMI hoặc dòng trạng thái slmgr khớp.
- `CHUA_KICH_HOAT`: có dữ liệu khớp sản phẩm Windows hiện tại và trạng thái chưa kích hoạt.
- `CAN_XAC_MINH_*`: WMI lỗi, lệnh timeout, trạng thái grace hoặc thông tin không đủ rõ. Không đánh đồng với máy chưa có bản quyền.

**Tự kiểm tra cú pháp trước khi khảo sát:** chạy `3_Kiem_Tra_Cu_Phap.bat`. Lệnh này chỉ chạy bộ phân tích cú pháp PowerShell, **không** thay thế kiểm thử thực tế Windows/WMI/Office. Trên PowerShell 2.0, công cụ kiểm tra token có giới hạn so với parser PS3+.
