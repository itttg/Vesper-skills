---
name: "fa-grab-import-sap"
description: >
  Chuyển báo cáo Grab For Business (GFB Billing Calculation Report) sang
  template SAP Journal Entry dạng 2 sheet JE-Header / JE-Line theo mẫu
  "SAP_Template import JE bằng WB". Dùng rules tương tự skill gốc nhưng
  không dùng folder con bên trong skill.
alwaysAllow:
  - Bash
  - Write
---

# Skill: fa-grab-import-sap

Skill này dùng để chuyển file Grab For Business sang template SAP dạng:
- `JE-Header`
- `JE-Line`

## Mục tiêu
1. Đọc file Grab billing `.xlsx`
2. Đọc template `SAP_Template import JE bằng WB.xlsx` đặt cùng cấp với `SKILL.md`
3. Tạo **1 JE cho mỗi phòng / department** dựa trên `GROUP_NAME`
4. Trong `JE-Line`, vẫn giữ chi tiết **theo từng hóa đơn thực tế**
5. Mỗi hóa đơn sinh tối đa 3 line:
   - line chi phí
   - line VAT
   - line công nợ nhà cung cấp Grab
6. Xuất file Excel mới theo đúng cấu trúc 2 sheet của template
7. Xuất thêm 2 file text tab-delimited theo mẫu import:
   - `Header.txt`
   - `Line.txt`
8. Xuất summary JSON để kế toán review

## Quy tắc bắt buộc
- Luôn bắt đầu bằng kiểm tra đầu vào và template trước khi chạy script.
- Chỉ tạo file import review-ready, không tự động post vào SAP.
- Nếu hướng dẫn line 1 của template mâu thuẫn với dòng mẫu, ưu tiên:
  1. báo rõ mâu thuẫn cho user
  2. dùng giá trị mẫu đang có trong template làm mặc định kỹ thuật
- Nếu file template `.xls` không đọc được bằng `openpyxl`, ưu tiên dùng bản `.xlsx` tương đương.
- Nếu dữ liệu hóa đơn thiếu `INVOICE_NUMBER`, vẫn cho phép tạo JE nhưng phải ghi warning vào summary.
- Nếu `GROUP_NAME` trống:
  1. map vào `General`
  2. ghi warning vào summary
  3. tô vàng các dòng liên quan trên file Excel output để kế toán review
- Giữ cơ chế tạo folder con `Output` trong thư mục làm việc của tháng và lưu toàn bộ file do skill tạo ra vào đó.
- Nếu user chỉ định `--output` là file `.xlsx`, skill vẫn sẽ tự tạo folder con `Output` trong thư mục cha của path đó và lưu file vào trong folder `Output`.

## Mapping đang dùng mặc định
### Header
- `JdtNum` = số JE tăng dần theo department
- `ReferenceDate` = posting date format `YYYYMMDD`
- `Memo` = diễn giải theo tháng + **mã phòng** (không dùng số hóa đơn ở header)
- `ProjectCode` = lấy từ dòng mẫu template
- `TaxDate` = `ReferenceDate`
- `U_VoucherTypeID` = lấy từ dòng mẫu template
- `U_Branch` = lấy từ dòng mẫu template

### Line chi phí
- `AccountCode` = `64281001`
- `Debit` = `PRE_VAT_DELIVERY_FEE + PRE_VAT_SERVICE_FEE`
- `CostingCode..CostingCode5` = tách từ chuỗi default `17020101;M999998;M02;ADM;M0100000`
- `ProjectCode` = lấy từ dòng mẫu template
- `BPLID` = lấy từ dòng mẫu template
- `TaxDate` = `VAT_INVOICE_DATE` format `YYYYMMDD`
- `LineMemo` / `U_RemarksJE` = diễn giải theo **hóa đơn thực tế**

### Line VAT
- `AccountCode` = `13311001`
- `Debit` = `VAT_VALUE_DELIVERY_FEE + VAT_VALUE_SERVICE_FEE`
- `BaseSum` = số tiền pre-VAT
- `TaxDate` = `VAT_INVOICE_DATE` format `YYYYMMDD`
- `TaxGroup` = `PVN5`
- `U_InvNo` = `INVOICE_NUMBER`
- `U_Invdate` = `VAT_INVOICE_DATE`
- `U_InvSeri` = `VAT_INVOICE_SERIAL` bỏ ký tự `1` đầu nếu có
- `U_isVat` = `Y`
- `U_BPcode` = `V00000070`
- `U_BPname` = `CÔNG TY TNHH GRAB`
- `U_TaxCode` = `312650437`

### Line công nợ
- `AccountCode` = `33111001`
- `Credit` = tổng debit của từng hóa đơn
- `ShortName` = `V00000070`
- `TaxDate` = `VAT_INVOICE_DATE` format `YYYYMMDD`

## Quy tắc đánh số line
- `ParentKey` = `JdtNum` của department
- `LineNum` phải chạy **liên tục trong từng JE**
- Không reset `LineNum = 0,1,2` cho từng hóa đơn nữa

## Thông tin cố định hiện dùng
| Trường | Giá trị |
|---|---|
| Vendor code | V00000070 |
| Vendor name | CÔNG TY TNHH GRAB |
| Vendor tax code | 312650437 |
| Expense account | 64281001 |
| VAT account | 13311001 |
| AP control account | 33111001 |
| Tax Group | PVN5 |
| Default costing string | 17020101;M999998;M02;ADM;M0100000 |

## Yêu cầu môi trường tối thiểu
- Python 3
- Package Python: `openpyxl`
- Có thể cài bằng:
```bash
pip install -r "$VESPER_SKILL_DIR/requirements.txt"
```

## Input hỗ trợ
### Cách 1: truyền file GFB trực tiếp
```bash
python "$VESPER_SKILL_DIR/convert_gfb_to_je_template.py" \
  "C:/path/GFB Billing Calculation Report_xxx.xlsx"
```

### Cách 2: truyền folder tháng
```bash
python "$VESPER_SKILL_DIR/convert_gfb_to_je_template.py" \
  "C:/path/202603"
```

### Chỉ định output riêng
```bash
python "$VESPER_SKILL_DIR/convert_gfb_to_je_template.py" \
  "C:/path/202603" \
  --output "C:/path/202603/SAP JE Import_Grab_03.2026.xlsx"
```

### Chỉ định template riêng
```bash
python "$VESPER_SKILL_DIR/convert_gfb_to_je_template.py" \
  "C:/path/202603" \
  --template "C:/path/SAP_Template import JE bằng WB.xlsx"
```

## Template đi kèm skill
- Template mặc định đặt cùng cấp với `SKILL.md` tại: `SAP_Template import JE bằng WB.xlsx`
- Khi copy skill cho user khác, chỉ cần giữ nguyên folder skill này là script có thể tự tìm template mặc định.
- Vẫn có thể override bằng `--template` nếu cần dùng mẫu khác.

## Output
Mặc định tạo trong folder con `Output`:
- `SAP JE Import_Grab_<MM>.<YYYY>.xlsx`
- `Header.txt`
- `Line.txt`
- `SAP JE Import_Grab_<MM>.<YYYY> - Summary.json`

## Quy tắc export TXT
- TXT được xuất từ 2 sheet `JE-Header` và `JE-Line`.
- Bỏ dòng hướng dẫn ở row 1, giữ từ row 2 trở đi để khớp format import mẫu.
- Phân tách bằng tab.
- Encoding dùng `UTF-16` để tương thích cách export mẫu đang dùng.
- `TaxDate` trên TXT xuất theo format `YYYYMMDD`.
- `U_Invdate` trên TXT xuất theo format `dd/mm/yyyy`.
- Các cột số tiền như `Debit`, `Credit`, `FCDebit`, `FCCredit`, `BaseSum` xuất với 2 chữ số thập phân.

## Validation cần báo cho user
Sau khi chạy xong, luôn báo:
1. Số dòng GFB hợp lệ
2. Số JE header đã tạo
3. Số JE line đã tạo
4. Số department đã tạo JE
5. Tổng debit / credit
6. Các warning:
   - thiếu số hóa đơn
   - thiếu ngày hóa đơn
   - VAT = 0
   - `GROUP_NAME` trống được map vào `General`
   - template không tìm thấy / fallback template

## Hạn chế hiện tại
- Mapping costing đang dùng default cố định, chưa phân bổ theo department.
- Template `.xls` cũ có thể không đọc trực tiếp được bằng `openpyxl`; nên dùng bản `.xlsx` tương đương.
- Nếu template `JE-Line` cũ chưa có cột `TaxDate`, script sẽ tự chèn cột này vào output workbook theo vị trí sau `ReferenceDate1` và trước `Reference1`.
- Chưa tự suy luận voucher type / branch ngoài giá trị dòng mẫu template.
