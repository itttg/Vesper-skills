---
name: gfb-to-sap
description: >
  Chuyển đổi báo cáo cước Grab For Business (GFB Billing Calculation Report)
  thành file Excel import SAP B1. Mỗi dòng GFB tạo 2 dòng SAP (chi phí TK 64281001
  + thuế GTGT TK 13311001) và 1 dòng Credit cuối cho nhà cung cấp Grab.
  File input là GFB Excel (.xlsx), output là SAP Import Excel (.xlsx).
  Kích hoạt khi nghe thấy: "GFB", "Grab For Business", "hạch toán Grab",
  "chi phí Grab", "cước Grab", "import SAP Grab", "bút toán Grab",
  "chuyển đổi GFB", "chuyển hóa đơn Grab", hoặc bất kỳ yêu cầu nào liên
  quan đến xử lý hóa đơn Grab và nhập vào SAP. Luôn sử dụng skill này
  khi người dùng đề cập đến Grab kết hợp với SAP, kế toán, hoặc hạch toán.
---

# Skill: GFB Billing → SAP Import (Per Invoice)

## Mục đích
Chuyển đổi file báo cáo cước Grab For Business (GFB Billing Calculation Report)
thành file Excel import SAP B1. Mỗi dòng GFB tạo **2 dòng SAP** (chi phí + thuế),
cộng **1 dòng Credit** cuối cùng cho nhà cung cấp.

## Thông tin cố định

| Thông số | Giá trị |
|---|---|
| TK Chi phí (Debit) | 64281001 - Chi phí bằng tiền khác |
| TK Thuế GTGT (Debit) | 13311001 - Thuế GTGT được khấu trừ |
| Mã nhà cung cấp | V00000070 |
| Tên đối tác | CÔNG TY TNHH GRAB |
| MST | 312650437 |
| Địa chỉ | 268 Tô Hiến Thành, TP.HCM, Quận 10 |
| Distr. Rule | 17020101;M999998;M02;ADM;M0100000 |
| Tax Group | PVN5 |
| Project | M02 |
| Branch | LEGACY |
| Payment Block | N |
| TK Credit (Control) | 33111001 |
| Offset Account (Credit line) | 64281001 |

## Logic chuyển đổi

Mỗi dòng GFB tạo 2 dòng SAP:

**Dòng 1 – Chi phí (Debit TK 64281001):**
- Amount = PRE_VAT_DELIVERY_FEE + PRE_VAT_SERVICE_FEE
- Có Distr. Rule, không có Tax Group

**Dòng 2 – Thuế GTGT (Debit TK 13311001):**
- Amount = VAT_VALUE_DELIVERY_FEE + VAT_VALUE_SERVICE_FEE
- Base Amount = amount dòng chi phí
- Có Tax Group = PVN5, có Seri HĐ
- Không có Distr. Rule

**Dòng cuối – Credit cho NCC:**
- Credit = tổng tất cả Debit (chi phí + thuế)
- G/L Code = V00000070, Control = 33111001
- Offset Account = 24121001
- Document Date = Posting Date (không phải ngày hóa đơn)

## Các trường động

| Trường SAP | Nguồn GFB |
|---|---|
| Số HĐ (Y) | INVOICE_NUMBER |
| Seri HĐ (Z) | VAT_INVOICE_SERIAL bỏ ký tự "1" đầu (1C26MGA → C26MGA) |
| Document Date (O) | VAT_INVOICE_DATE format dd.mm.yy |
| Posting Date (N) | Cuối tháng kế tiếp của tháng dữ liệu |
| Due Date (M) | = Posting Date |
| Remarks (J) | "Chi phí Grab tháng MM.YYYY_Grab" |
| Diễn giải (AK) | = Remarks |
| RemarksJE (AL) | = Remarks |

## Sắp xếp
Theo VAT_INVOICE_DATE tăng dần, rồi INVOICE_NUMBER tăng dần.

## Lọc dữ liệu
- Chỉ lấy dòng có AMOUNT > 0 VÀ COMPANY_NAME không rỗng (loại dòng tổng cuối file GFB)

## Posting Date
- Tự động: cuối tháng kế tiếp (dữ liệu T02 → 31/03, T03 → 30/04...)
- Có thể override bằng tham số `--posting-date dd/mm/yyyy`

## Quy trình thực hiện

### Bước 1: Xác định file đầu vào
Tìm file GFB trong thư mục người dùng chỉ định. File thường có tên:
`GFB Billing Calculation Report_*.xlsx`

Nếu người dùng chỉ định folder mà không chỉ rõ file, tìm file GFB trong folder đó.
Nếu không tìm thấy, hỏi người dùng.

### Bước 2: Chạy script chuyển đổi

```bash
python3 <skill-path>/scripts/convert_gfb_to_sap.py \
  "<đường dẫn file GFB>" \
  "<đường dẫn file output>" \
  [--posting-date dd/mm/yyyy]
```

File output nên lưu vào **cùng folder** với file input.
Tên file mặc định: `SAP Import_Grab_T{MM}.{YYYY}.xlsx`

### Bước 3: Báo cáo kết quả
Sau khi chạy xong, trình bày cho người dùng:
1. Số chuyến GFB và số dòng SAP đã tạo
2. Bảng thống kê theo phòng ban
3. Tổng Debit và kết quả validation
4. Posting Date được sử dụng
5. Link file output

### Bước 4: Trình bày file
Dùng `present_files` hoặc link `computer://` đến file output.

## Xử lý đặc biệt

**Admin Fee**: Dòng cuối GFB có BILLING_TYPE = "Admin Fee" (không có GROUP_NAME,
không có VERTICAL) vẫn là giao dịch hợp lệ. Script tự xử lý.

**Dòng tổng cuối file GFB**: File GFB có thể chứa dòng tổng hợp ở cuối
(không có COMPANY_NAME). Script tự loại bỏ nhờ filter COMPANY_NAME.

**Tổng tiền không khớp**: Kiểm tra NON_VAT_VALUE. Thông báo chênh lệch nếu có.

## Validation
Script tự kiểm tra:
- Tổng Debit SAP = Tổng AMOUNT trong GFB
- Thông báo nếu có chênh lệch
