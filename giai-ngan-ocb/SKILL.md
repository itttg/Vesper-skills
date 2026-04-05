---
name: giai-ngan-ocb
description: >
  Tự động điền form Đề nghị Giải Ngân kiêm Khế ước Nhận nợ (KUNN) gửi
  Ngân hàng OCB cho Công ty CP Trung Thủy - Đà Nẵng (hợp đồng tín dụng
  số 0239/2024/HĐTD-OCB-DN). Skill nhận thông tin từ người dùng (số tiền,
  lãi suất, ngày trả lãi đầu tiên) và tạo file Word đã điền sẵn, sẵn sàng
  ký và gửi ngân hàng. Kích hoạt khi nghe: "giải ngân", "KUNN", "khế ước
  nhận nợ", "hồ sơ vay OCB", "đề nghị giải ngân", "thanh toán bằng giải
  ngân", "lập hồ sơ giải ngân", hoặc bất kỳ yêu cầu nào liên quan đến
  việc chuẩn bị hồ sơ vay/giải ngân ngân hàng OCB.
---

# Skill: Đề nghị Giải Ngân kiêm Khế ước Nhận nợ (KUNN) - OCB

## Mục đích
Tự động điền form KUNN gửi OCB từ thông tin người dùng cung cấp.
Hợp đồng tín dụng: **0239/2024/HĐTD-OCB-DN** - Công ty CP Trung Thủy - Đà Nẵng.

## 2 loại form
| Loại | Template | Khi nào dùng |
|---|---|---|
| **VND** | `templates/KUNN_template_VND.docx` | Chuyển thẳng cho bên thụ hưởng VND |
| **Ngoại tệ** | `templates/KUNN_template_NGOAITE.docx` | Giải ngân vào TK công ty → mua ngoại tệ |

## Thông tin cố định (không cần hỏi)
- **Ngày ký**: Lấy ngày hiện tại tự động
- **Họ tên + Chức vụ**: Để trống (người dùng tự ký)
- **Năm hồ sơ**: Năm hiện tại

## Quy trình thực hiện

### Bước 1: Xác định loại giải ngân
Hỏi người dùng nếu chưa rõ:
- "Giải ngân VND trực tiếp cho bên thụ hưởng?" → dùng `KUNN_template_VND.docx`
- "Thanh toán ngoại tệ (mua ngoại tệ)?" → dùng `KUNN_template_NGOAITE.docx`

### Bước 2: Thu thập thông tin bắt buộc
Hỏi người dùng những thông tin sau (nếu chưa cung cấp):

| Thông tin | Ví dụ | Ghi chú |
|---|---|---|
| **Số tiền giải ngân** | 217,857,000 VND | Script tự chuyển sang chữ |
| **Lãi suất** | 9.5%/năm | Lãi suất hiện tại tại thời điểm giải ngân |
| **Ngày trả lãi đầu tiên** | 25/06/2026 | Định dạng dd/mm/yyyy |

### Bước 3: Tạo bộ hồ sơ 3 file
#### 3.1. Điền KUNN
```bash
python3 scripts/fill_kunn_form.py \
  "templates/1. KUNN_template_VND.docx" \
  "<output_kunn_path>" \
  --so-tien <số_tiền> \
  --lai-suat <lãi_suất> \
  --ngay-tra-lai <dd/mm/yyyy>
```

#### 3.2. Điền Lịch trả nợ
- Dùng template: `templates/2. LICH TRA NO KUNN 0239.docx`
- Input cần có: **tổng số tiền phải trả**
- Nguyên tắc chia nợ gốc: **12 kỳ**, chia đều, **phần dư để kỳ cuối**

```bash
python3 scripts/fill_repayment_schedule.py \
  "templates/2. LICH TRA NO KUNN 0239.docx" \
  "<output_schedule_path>" \
  --tong-so-tien <số_tiền> \
  --ngay-kunn <dd/mm/yyyy>
```

#### 3.3. Điền Bảng kê chứng từ
- Dùng template: `templates/3. Bang ke chung tu KUNN_0239.docx`
- Tự đọc các file PDF trong bộ hồ sơ thanh toán để liệt kê chứng từ
- Mỗi bộ hồ sơ có thể có **nhiều chứng từ**, phải tạo **đủ số dòng** theo số chứng từ đọc được
- Nếu PDF scan kém hoặc thiếu text, fallback tối thiểu là **tên file**, đồng thời nêu rõ điểm cần người dùng rà soát

```bash
python3 scripts/fill_supporting_documents.py \
  "templates/3. Bang ke chung tu KUNN_0239.docx" \
  "<output_docs_path>" \
  --input-folder "<thư_mục_hồ_sơ_thanh_toán>" \
  --tong-so-tien <số_tiền> \
  --ngay-kunn <dd/mm/yyyy>
```

### Quy tắc đặt tên file output
- `KUNN_0239_<loại>_<ngày>.docx`
- `LICH_TRA_NO_KUNN_0239_<ngày>.docx`
- `BANG_KE_CHUNG_TU_KUNN_0239_<ngày>.docx`

Lưu toàn bộ output vào cùng thư mục với tờ trình (nếu người dùng chỉ định),
hoặc vào thư mục làm việc hiện tại.

### Bước 4: Xác nhận với người dùng
Sau khi tạo xong, trình bày tóm tắt:
- Số tiền (số + chữ)
- Ngày ký
- Lãi suất
- Ngày trả lãi đầu tiên
- Số kỳ trả nợ và số tiền mỗi kỳ
- Số dòng chứng từ đã điền
- Link mở từng file

## Trường hợp có tờ trình PDF
Nếu người dùng cung cấp file tờ trình (PDF), đọc và trích xuất:
- **Số tiền**: Tìm dòng có "số tiền", "payment amount", "transfer amount" kèm con số
- **Loại thanh toán**: Tìm "USD/EUR/ngoại tệ" → ngoại tệ; ngược lại → VND

Xác nhận lại với người dùng trước khi điền:
*"Tôi đọc được từ tờ trình: số tiền X, loại Y. Lãi suất và ngày trả lãi đầu tiên là bao nhiêu?"*

## Các trường để trống (không điền)
- Điện thoại, Fax, Email
- Họ tên + Chức vụ người ký đại diện
- Văn bản ủy quyền số
- Ân hạn lãi: từ ngày…đến ngày…
- Phần dành cho ngân hàng OCB

## Ví dụ yêu cầu người dùng
> "Lập hồ sơ giải ngân 217 triệu, lãi suất 9.5%, ngày trả lãi 25/6"
> "Làm KUNN cho khoản thanh toán TK Studio 147 triệu, lãi 9.5%, trả lãi 25/6"
> "Giải ngân ngoại tệ, số tiền quy đổi 852,527,200 VND, lãi suất hiện tại 9.2%"
