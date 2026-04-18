---
name: "fa-tnt-ebill-import"
description: "Chuyển file hóa đơn điện Excel thành bộ file import SAP JE theo template JE mới, dùng cấu trúc skill phẳng và tăng cường kiểm tra đầu vào/đầu ra."
alwaysAllow:
  - Bash
  - Write
---

# Skill: fa-tnt-ebill-import

Skill này dùng để tạo bộ file import SAP JE từ file hóa đơn điện Excel.

## SOP thực hiện
1. Xác định file đầu vào là Excel hóa đơn điện hoặc thư mục chứa file đó.
2. Kiểm tra workbook, sheet dữ liệu và header thực tế của file nguồn.
3. Dò `header row` theo `header-based mapping + alias + normalization`, không phụ thuộc cứng vào vị trí cột.
4. Áp dụng rule chuyển đổi theo template JE mới đã được người dùng xác nhận.
5. Sinh bộ output review-ready trong thư mục riêng `output_<MM>_<YYYY>`, không tự động post vào SAP.
6. Đối chiếu tổng tiền đầu vào với tổng Debit và Credit đầu ra.
7. Sinh summary JSON kèm profile input detect được, warnings, control points và exception nếu có.

## Cải tiến chính so với skill gốc
- Không dùng folder bên trong skill; template, script, requirements và icon đặt cùng cấp với `SKILL.md`.
- Dò sheet dữ liệu linh hoạt hơn: quét nhiều sheet và chọn sheet phù hợp nhất.
- Mapping header robust hơn nhờ normalize dấu tiếng Việt, khoảng trắng và alias mở rộng.
- Parse số tiền an toàn hơn với định dạng có dấu phẩy / dấu chấm / khoảng trắng.
- Sinh `warnings` và `skipped_rows_detail` rõ hơn để kế toán review.
- Kiểm tra duplicate invoice key (`invoice_no + invoice_series + issue_date`).
- Summary output đầy đủ hơn để tiện audit và handoff.

## Input mapping hiện tại
Skill hỗ trợ đọc dữ liệu theo tên cột và alias, ưu tiên các cột logic sau:
- `row_no`
- `gross_amount`
- `invoice_no`
- `invoice_series`
- `issue_date`

### Các alias chính đang hỗ trợ
- `row_no`
  - `STT`
- `gross_amount`
  - `TONG_NO`
  - `Tổng nợ`
  - `Tổng tiền`
  - `Gross Amount`
  - `Thành tiền`
- `invoice_no`
  - `số HD`
  - `Số HĐ`
  - `SO HD`
  - `Sery HĐ`
  - `Số hóa đơn`
  - `Invoice No`
- `invoice_series`
  - `Ký hiệu`
  - `Seri HĐ`
  - `Mã kí hiệu`
  - `Mã ký hiệu`
  - `Invoice Series`
- `issue_date`
  - `NGÀY PHÁT HÀNH`
  - `Ngày PH`
  - `Ngày phát hành`
  - `Issue Date`
  - `Document Date`

### Profile input hiện có
- `legacy_layout`
  - ví dụ: `số HD`, `Ký hiệu`, `NGÀY PHÁT HÀNH`
- `evn_layout_202603`
  - ví dụ: `Sery HĐ`, `Mã kí hiệu`, `Ngày PH`
- `generic_header_mapping`
  - fallback khi detect được cột bắt buộc nhưng không match profile signature cụ thể

## Rule chuyển đổi hiện tại
Các rule dưới đây bám theo template JE mới đã confirm:
- Toàn bộ file đầu vào chỉ sinh **1 journal / 1 dòng header** trên sheet `JE-Header` cho tất cả hóa đơn.
- Mỗi hóa đơn vẫn sinh **3 dòng line** trên sheet `JE-Line` và cùng thuộc về journal duy nhất đó:
  - 1 dòng chi phí `62721001` (Debit)
  - 1 dòng VAT `13311001` (Debit)
  - 1 dòng công nợ `33111001` (Credit)
- VAT mặc định: `8%`
- Base amount = `ROUND(TONG_NO / 1.08)`
- VAT amount = `TONG_NO - Base amount`
- Trên dòng VAT, cột `BaseSum` phải lấy **giá trị trước thuế** (`Base amount`), không lấy tổng tiền thanh toán.
- Vendor mặc định:
  - `Mã đối tác`: `V00000162`
  - `Tên đối tác`: `CHI NHÁNH TỔNG CÔNG TY ĐIỆN LỰC TPHCM TNHH-CÔNG TY ĐIỆN LỰC SÀI GÒN`
  - `MST`: `0300951119-001`
- Project: `M02`
- Voucher Type: `7012`
- Branch / Note for Import: `7`
- Costing split dòng chi phí tách từ:
  - `12090310;M999994;M02;PMO;M0100000`
- Tax Group dòng VAT: `PVN5`
- Diễn giải mặc định suy ra từ ngày phát hành, mặc định lấy kỳ tiêu thụ là **tháng trước của ngày phát hành**.
- Sheet `JE-Header` cột `D` (`Memo`) dùng format rút gọn:
  - `Điện tiêu thụ tháng <m> năm <yyyy>`
- Với journal header duy nhất:
  - `ReferenceDate`, `TaxDate` lấy theo **ngày tạo file import**.
  - `Memo` lấy theo **ngày phát hành sau cùng** trong danh sách hóa đơn.
- Sheet `JE-Line`:
  - cột `I` (`DueDate`) lấy theo `ReferenceDate` trên sheet `JE-Header`.
  - cột `L` (`LineMemo`) lấy đúng theo giá trị cột `D` của sheet `JE-Header`.
  - cột `M` (`ReferenceDate1`) lấy theo `ReferenceDate` trên sheet `JE-Header`.
  - cột `TaxDate` lấy theo **ngày thực tế của hóa đơn**.
  - cột `U_InvSeri`: nếu ký hiệu hóa đơn bắt đầu bằng ký tự `1` thì bỏ ký tự `1` đầu tiên trước khi ghi ra output SAP.

## Template nội bộ
Skill sinh theo template đã được đóng gói sẵn ngay trong skill:
- Template workbook nội bộ: `fa-tnt-ebill-import-template.xlsx`
- Workbook output: `JE-Header`, `JE-Line`
- Text export:
  - `Header.txt`
  - `Line.txt`

## Validation hiện tại
### Validation đầu vào
- Quét tối đa `20` dòng đầu trên mỗi sheet để tìm `header row`.
- Chỉ chạy tiếp nếu detect được đủ các cột logic bắt buộc:
  - `row_no`
  - `gross_amount`
  - `invoice_no`
  - `invoice_series`
  - `issue_date`
- Chặn hoặc skip có log rõ cho các dòng:
  - thiếu `gross_amount`
  - thiếu `invoice_no`
  - thiếu `invoice_series`
  - thiếu `issue_date`
  - `gross_amount <= 0`
- Cảnh báo duplicate invoice key.

### Validation đầu ra
- Tổng đầu vào phải bằng tổng Debit đầu ra.
- Tổng đầu vào phải bằng tổng Credit đầu ra.
- Sheet `JE-Header` chỉ có `1` dòng journal cho toàn bộ hóa đơn.
- Số dòng output trên `JE-Line` phải bằng `3 x số hóa đơn hợp lệ`.
- `LineNum` trên `JE-Line` phải chạy tuần tự trong journal duy nhất.
- Kiểm tra file output chính, `Header.txt`, `Line.txt`, `.summary.json` đều được sinh ra.

### Summary output
Sau khi chạy, skill sinh thêm file `.summary.json` để phục vụ review với các thông tin như:
- `invoice_count`
- `header_row_count`
- `line_row_count`
- `input_total`
- `debit_total`
- `credit_total`
- `sheet_name`
- `header_row`
- `input_profile`
- `detected_headers`
- `skipped_rows`
- `skipped_rows_detail`
- `warning_count`
- `warnings`
- `duplicate_invoice_keys`
- `output_folder`
- `period_month`
- `period_year`

## Cách chạy
### Truyền trực tiếp file đầu vào
```bash
python "$VESPER_SKILL_DIR/build_fa_tnt_ebill_import.py" "C:/path/input.xlsx"
```

### Truyền thư mục chứa file đầu vào
```bash
python "$VESPER_SKILL_DIR/build_fa_tnt_ebill_import.py" "C:/path/folder"
```

### Chỉ định file output
```bash
python "$VESPER_SKILL_DIR/build_fa_tnt_ebill_import.py" "C:/path/input.xlsx" "C:/path/output/target.xlsx"
```

## Output
Mặc định skill tạo thư mục con theo kỳ tiêu thụ:
- `output_<MM>_<YYYY>/`

Trong đó sinh ra:
- `SAP_Import by JE_<MM>_<YYYY>.xlsx`
- `<output>.summary.json`
- `Header.txt`
- `Line.txt`

## Control points cần kế toán confirm
- Xác nhận lại VAT, tài khoản, vendor và các field fix trước khi dùng rộng.
- Với format EVN mới, cần confirm mapping nghiệp vụ:
  - `Sery HĐ` có đúng là **Số HĐ** hay không
  - `Mã kí hiệu` / `Mã ký hiệu` có đúng là **Seri/Ký hiệu HĐ** hay không
- Confirm `TaxGroup`, `VoucherType`, `AP account`, costing split là đúng với template SAP đang dùng.
- Confirm `ReferenceDate`, `TaxDate`, `DueDate` theo rule hiện tại là phù hợp.

## Assumptions / Risks
- Rule hiện tại vẫn chuyên biệt cho bộ hóa đơn điện mẫu đã cung cấp.
- VAT, account, vendor, project, costing split, voucher type đang hard-coded theo xác nhận hiện tại.
- Nếu file nguồn xuất hiện header mới chưa có trong alias list thì cần bổ sung alias hoặc profile input tương ứng.
- Output chỉ là file review-ready, không tự động post SAP.
