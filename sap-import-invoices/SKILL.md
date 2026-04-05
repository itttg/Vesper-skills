---
name: "sap-import-invoices"
description: "Phân tích hóa đơn PDF/HTML, học rule từ sổ nhật ký chung và sinh bộ import SAP gồm Header.txt, Line.txt, Excel review."
alwaysAllow:
  - Bash
  - Write
---

# SAP Import Invoices

Skill này dùng để chuyển một thư mục hóa đơn đầu vào thành bộ file import SAP theo mẫu công ty.

## Bộ dữ liệu tham chiếu đi kèm skill
- `data/rules/Main account.xlsx`
- `data/rules/SAP Vendor.xlsx`
- `data/rules/SAP_Headcode 3.xlsx`
- `data/rules/Sổ nhật ký chung 1.xlsx`
- `data/rules/Header.txt`
- `data/rules/Line.txt`
- `data/rules/SAP_Template import JE bằng WB.xlsx`

## Rule đã chốt
- Học rule ưu tiên từ **Sổ nhật ký chung**
- Được suy luận theo lịch sử gần nhất cùng vendor / cùng loại chi phí
- VAT account chuẩn: `13331001`
- Tax group:
  - `10% -> PVN1`
  - `5% -> PVN2`
  - `0% -> PVN3`
  - `Không chịu thuế -> PVN4`
  - `8% -> PVN5`
- `U_VoucherTypeID = 1704`
- `U_Branch = theo template`
- `ProjectCode line = theo template`
- Vendor chưa chắc / vendor mới: để trống các BP fields trong TXT và tô vàng trong Excel review

## Output
Skill luôn sinh tại thư mục output đã chỉ định:
- `Header.txt` (UTF-16 LE BOM)
- `Line.txt` (UTF-16 LE BOM)
- `SAP_Review.xlsx`
- `run_summary.json`

## Cách chạy

```bash
python "$VESPER_SKILL_DIR/scripts/build_sap_import.py" \
  "C:/Data/SAP Import/3. Invoice/Thang 3" \
  "C:/Data/SAP Import/4. Output/Thang 3"
```

## Cách xử lý mặc định
1. Đọc toàn bộ PDF/HTML trong thư mục invoice
2. Trích xuất thông tin hóa đơn
3. Match vendor theo MST trước, sau đó fuzzy theo tên
4. Học mapping account/headcode/bộ phận từ sổ nhật ký chung
5. Sinh nhiều JE line theo số mức thuế thực tế của hóa đơn
6. Tô vàng trong Excel review cho trường mới / chưa chắc / suy luận yếu

## Lưu ý
- File PDF scan/xấu có thể trích xuất text không hoàn hảo; khi đó Excel review là nơi user bổ sung.
- Rule hiện tại ưu tiên độ an toàn hơn độ “đoán bừa”. Giá trị không chắc sẽ để trống + tô vàng.
