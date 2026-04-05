# Skill - SAP Import từ hóa đơn đầu vào

Cập nhật ngày: 2026-04-04 16:21 GMT+7

## Mục tiêu
Chuẩn hóa quy tắc tạo file template import SAP từ hóa đơn đầu vào.

## Nguồn tham chiếu
- File template import SAP
- Sổ nhật ký chung để tham chiếu cách hạch toán cũ
- Bảng headcode công ty
- Bảng main account công ty
- Thông tư 99/2025/BTC

## Nguyên tắc xử lý
1. Tạo **1 file import tổng** cho toàn bộ hóa đơn trong thư mục xử lý.
2. Nghiên cứu lịch sử hạch toán trước khi map account/headcode.
3. Chỉ thay đổi **headcode** trong `Distr. Rule`, các thành phần khác giữ cố định theo template nếu chưa có hướng dẫn khác.
4. Tô màu vàng các giá trị chưa chắc chắn.

## Quy tắc tiền tệ
- Giá trị tiền **không thêm hậu tố `VND`**.
- Chỉ dùng số có dấu phân tách hàng nghìn nếu template đang theo format đó.

## Quy tắc tax group
- 10% -> `PVN1`
- 5% -> `PVN2`
- 0% -> `PVN3`
- 8% -> `PVN5`
- Không chịu thuế -> `PVN4`

## Quy tắc account cấp cao
- **Tài sản cố định** -> account nhóm `211...`
- **Công cụ dụng cụ** -> account nhóm `242...`
- **Chi phí** -> account nhóm `642...`
- **VAT đầu vào** ưu tiên map theo tính chất thực tế và lịch sử bút toán

## Quy tắc nghiệp vụ đã chốt trong case này
- Viettel -> headcode `17040600` (Điện thoại, fax)
- TM Grow -> headcode `12080250` (HT xử lý nước hồ bơi)
- Diệu Phúc -> tính chất `CCDC`
- BYD / Harmony -> tính chất `TSCĐ`
- IDC -> tự map theo lịch sử hạch toán phù hợp nhất

## Checklist trước khi import SAP
- Kiểm tra BP code / mã đối tác
- Kiểm tra tax group
- Kiểm tra account chính
- Kiểm tra headcode / distribution
- Kiểm tra số hóa đơn, ký hiệu, ngày chứng từ
- Kiểm tra tổng trước thuế, VAT, tổng thanh toán
- Kiểm tra các ô tô vàng
