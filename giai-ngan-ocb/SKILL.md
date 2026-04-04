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

### Bước 3: Chạy script điền form
```bash
python3 scripts/fill_kunn_form.py \
  "templates/KUNN_template_VND.docx" \
  "<output_path>" \
  --so-tien <số_tiền> \
  --lai-suat <lãi_suất> \
  --ngay-tra-lai <dd/mm/yyyy>
```

Đặt tên file output theo quy tắc: `KUNN_0239_<loại>_<ngày>.docx`
Ví dụ: `KUNN_0239_VND_01042026.docx` hoặc `KUNN_0239_NGOAITE_01042026.docx`

Lưu file output vào cùng thư mục với tờ trình (nếu người dùng chỉ định),
hoặc vào thư mục làm việc hiện tại.

### Bước 4: Xác nhận với người dùng
Sau khi tạo xong, trình bày tóm tắt:
- Số tiền (số + chữ)
- Ngày ký
- Lãi suất
- Ngày trả lãi đầu tiên
- Link mở file

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
