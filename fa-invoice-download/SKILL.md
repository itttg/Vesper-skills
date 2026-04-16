---
name: "FA Invoice Download"
description: "Đăng nhập cổng thông tin hóa đơn điện tử từ credential.txt, lọc theo thời gian user yêu cầu, tải toàn bộ XML vào thư mục Files, rồi parse XML để cập nhật InvoiceList.xlsx với bộ cột control-ready."
alwaysAllow:
  - Bash
  - Write
---

# FA Invoice Download

Skill này hỗ trợ workflow tải hóa đơn XML từ cổng thông tin thuế và cập nhật file Excel review-friendly cho kế toán.

Helper script của skill được đặt cùng cấp với `SKILL.md`, không dùng thư mục con `scripts/`.

## Mục tiêu

1. Đọc input gồm:
   - `fromDate`
   - `toDate`
   - `folderPath`
2. Kiểm tra cấu trúc thư mục làm việc.
3. Đọc `credential.txt` hoặc `Credentials.txt` trong folder user cung cấp.
4. Đăng nhập cổng thông tin thuế bằng browser automation.
5. Tìm đến chức năng xuất / tải XML hóa đơn.
6. Lọc theo khoảng thời gian user yêu cầu.
7. Download toàn bộ hóa đơn XML trong kỳ.
8. Lưu vào thư mục `Files` với tên file:
   - `{invoiceSeries}_{invoiceNumber}.xml`
9. Parse toàn bộ XML và cập nhật `InvoiceList.xlsx`.
10. Kết thúc bằng reconciliation summary.

## Input bắt buộc

User phải cung cấp đủ:

- `fromDate`
- `toDate`
- `folderPath`

Ví dụ prompt:

- `Dùng fa-invoice-download. fromDate=01/03/2026, toDate=31/03/2026, folderPath=C:/Data/Build Demand`

## Cấu trúc thư mục kỳ vọng

Trong `folderPath` phải có:

- `credential.txt` hoặc `Credentials.txt`
- `Files/`
- `InvoiceList.xlsx`

## SOP bắt buộc trước khi chạy

Trước khi automation, luôn đi theo checklist này:

1. Xác nhận `fromDate`, `toDate`, `folderPath`.
2. Kiểm tra folder tồn tại.
3. Kiểm tra file credentials tồn tại.
4. Kiểm tra thư mục `Files` tồn tại; nếu chưa có thì tạo.
5. Kiểm tra `InvoiceList.xlsx` tồn tại.
6. Kiểm tra workbook có sheet `Invoice_Tax_Lines`.
7. Xác nhận portal không chặn bởi captcha / OTP / MFA.

Nếu một điều kiện không đạt, dừng và báo lỗi rõ ràng.

## Workflow thực thi

### Bước 1 - Inspect local files trước

Dùng file tools để:

- đọc credentials
- kiểm tra `Files`
- kiểm tra `InvoiceList.xlsx`

Không vào browser trước khi chắc chắn input folder hợp lệ.

### Bước 2 - Dùng browser automation để login

Khi user nói mở / vào / truy cập website, luôn dùng `browser_tool`.

Flow khuyến nghị:

1. open browser
2. navigate tới URL từ `Link`
3. snapshot
4. tìm form login
5. fill user/password
6. submit
7. snapshot lại sau mỗi lần chuyển màn hình

### Bước 3 - Xử lý login

- Dùng `User` và `Pass` từ credential file
- Không in raw password ra log hoặc final answer
- Nếu portal yêu cầu captcha / OTP / human verification thì dừng và yêu cầu user hỗ trợ
- Nếu login xong nhưng UI portal lỗi, ưu tiên API fallback dựa trên phiên đã xác thực

### Bước 4 - Tìm chức năng tải XML

Sau khi login:

- tìm khu vực tra cứu / danh sách hóa đơn / xuất XML
- ưu tiên quan sát bằng `snapshot` hoặc `find`
- xác nhận loại ngày đang lọc là gì:
  - ngày lập
  - ngày ký
  - ngày phát hành

Nếu UI lỗi hoặc route lookup không hoạt động ổn định, cho phép dùng authenticated API fallback để:
- tìm danh sách hóa đơn
- tải XML từng hóa đơn

### Bước 5 - Lọc theo thời gian user yêu cầu

Áp dụng `fromDate` và `toDate` đúng format mà portal/API yêu cầu.

### Bước 6 - Download toàn bộ XML

Nguyên tắc:

- ưu tiên XML
- nếu có bulk export hợp lệ thì dùng
- nếu chỉ tải từng dòng thì loop từng hóa đơn
- đặt tên file:
  - `{invoiceSeries}_{invoiceNumber}.xml`
- lưu tại:
  - `{folderPath}/Files`

Nếu file đã tồn tại:

- mặc định skip-and-log
- chỉ overwrite khi user yêu cầu refresh

### Bước 7 - Parse XML và cập nhật Excel

Sau khi download xong, chạy script local:

```bash
python "$VESPER_SKILL_DIR/update_invoice_list.py" "C:/Path/To/Folder"
```

Script sẽ:

- đọc tất cả file XML trong `Files`
- parse invoice data theo hướng heuristic + fallback
- upgrade header của `InvoiceList.xlsx` nếu còn thiếu cột
- append dữ liệu mới vào sheet `Invoice_Tax_Lines`
- nếu hóa đơn đã tồn tại theo duplicate key thì **không import lại vào Excel**
- tạo file summary JSON và error CSV nếu cần

### Bước 8 - Validation / reconciliation

Luôn kết thúc bằng đối chiếu:

- số invoice nhìn thấy trên portal/API
- số file XML download được
- số file XML parse thành công
- số row append mới vào Excel
- số duplicate bị skip vì hóa đơn đã tồn tại
- số file lỗi parse / download

## Mapping output kỳ vọng

Workbook `InvoiceList.xlsx`, sheet `Invoice_Tax_Lines` cần tối thiểu các cột sau:

### Existing review columns
- `InvoiceDate`
- `InvoiceNo`
- `SellerCompany`
- `SellerTaxCode`
- `BuyerName`
- `BuyerTaxCode`
- `Description`
- `TaxRate`
- `AmountBeforeTax`
- `TaxAmount`
- `AmountAfterTax`
- `DetailsText`
- `Status`
- `Notes`

### Control-ready columns
- `InvoiceFormNo`
- `InvoiceSeries`
- `InvoiceNumber`
- `TaxAuthorityCode`
- `SigningDate`
- `InvoiceType`
- `InvoiceNature`
- `PaymentMethod`
- `SellerAddress`
- `BuyerAddress`
- `Currency`
- `ExchangeRate`
- `AmountInWords`
- `SourceFile`

## Rule nghiệp vụ

1. Một invoice có thể sinh nhiều dòng output.
2. Nếu hóa đơn có nhiều mức thuế VAT khác nhau, phải lưu **mỗi mức thuế thành 1 line riêng**.
3. Các thông tin chung của cùng một hóa đơn phải được lặp lại giống nhau trên các line đó.
4. Ưu tiên group theo VAT rate summary trong XML; nếu không có thì group từ invoice line/item line.
5. Nếu XML chỉ có summary totals, tạo một dòng summary và ghi limitation vào `Notes`.
6. Không đoán mạnh khi schema không rõ; đưa vào exception / review.
7. Không tự động post bút toán.
8. Không bỏ qua parse failure một cách im lặng.
9. Nếu rerun, hóa đơn đã tồn tại thì skip, không import lại vào Excel để tránh user nhầm.

## Cách dùng script local

### Chạy mặc định

```bash
python "$VESPER_SKILL_DIR/update_invoice_list.py" "C:/Users/.../Build Demand"
```

### Chỉ định workbook khác

```bash
python "$VESPER_SKILL_DIR/update_invoice_list.py" "C:/Users/.../Build Demand" --workbook "InvoiceList.xlsx"
```

### Bỏ qua chống duplicate

```bash
python "$VESPER_SKILL_DIR/update_invoice_list.py" "C:/Users/.../Build Demand" --disable-dedupe
```

## Exception handling

Dừng và hỏi user khi gặp:

- captcha / OTP / MFA
- portal đổi UI lớn và API fallback cũng thất bại
- XML schema không parse được ở nhiều file
- workbook thiếu sheet hoặc lỗi format nghiêm trọng
- không xác định được ký hiệu / số hóa đơn để đặt tên file

## Final response format

Sau khi chạy xong, trả lời ngắn gọn theo mẫu:

1. Kỳ thời gian đã dùng
2. Folder xử lý
3. Số hóa đơn tìm thấy
4. Số XML tải được
5. Số XML parse được
6. Số dòng ghi mới vào Excel
7. Số duplicate bị skip
8. Danh sách lỗi / exception chính
9. Ghi chú reconcile
