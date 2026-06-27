# Push Bot — PENDING BILL CHECK

Bot Telegram lắng nghe keyword **`push`**. Khi nhận `push`, bot sẽ:

1. Gọi API `export-detail` (tương đương lệnh curl) để tải file Excel chi tiết đơn.
2. Dựng bảng pivot **PENDING BILL CHECK** giống mẫu (lọc đơn *tồn kết nối Pending*, gom theo **Zone › Current Post Office › Order ID**, cột theo ngày, có Grand Total).
3. Gửi lại file Excel kết quả vào chat.

## Cấu trúc

| File | Vai trò |
|------|---------|
| `bot.py` | Bot Telegram, lắng nghe keyword `push` / lệnh `/push` |
| `downloader.py` | Gọi API export-detail, lưu file Excel (giống curl) |
| `pivot.py` | Đọc Excel nguồn → dựng & xuất pivot giống hình |
| `config.json` | Toàn bộ cấu hình: token, mã trạng thái, bảng Zone |
| `requirements.txt` | Thư viện cần cài |

## Cài đặt

```bash
pip install -r requirements.txt
```

## Cấu hình (`config.json`)

Bắt buộc điền trước khi chạy:

- `telegram.bot_token` — token bot Telegram (lấy từ @BotFather).
- `api.bearer_token` — Bearer token cho API (phần `token_config` trong curl).

Các mục quan trọng khác:

- `pivot.pending_status_codes` — **danh sách mã CURRENT STATUS coi là "tồn kết nối Pending"**. Mặc định `["110"]`. Sửa lại đúng (các) mã nghiệp vụ của anh/chị, ví dụ `["110","210"]`.
- `pivot.exclude_test` — `true` để loại đơn test (so khớp `test_keywords` trong tên người gửi/nhận).
- `zone_mapping.by_post_office` — **bảng ánh xạ Post Office → Zone** của anh/chị. Có thể dùng mã đầy đủ (`"KANP001":"Zone 1"`) hoặc `by_prefix` (`"KAN":"Zone 1"`).
- `api.branch_code` — mặc định `PRE,PNP,SVA,KAN`.
- `api.date_range_days` — khoảng ngày tính ngược từ hôm nay (mặc định 7).

## Chạy bot

```bash
python bot.py
```

Sau đó vào chat Telegram gõ `push` (hoặc `/push`). Bot tải dữ liệu, dựng pivot và gửi lại file.

## Test nhanh phần pivot (không cần Telegram)

```bash
python pivot.py duong_dan_file_nguon.xlsx ket_qua.xlsx config.json
```

## Xử lý lỗi `telegram.error.TimedOut` (không kết nối được Telegram)

Lỗi này nghĩa là máy không vào được `api.telegram.org` (mạng chậm hoặc bị chặn). Thử theo thứ tự:

1. **Kiểm tra mạng tới Telegram:** mở trình duyệt vào
   `https://api.telegram.org/bot<BOT_TOKEN>/getMe`. Nếu không mở được → mạng đang chặn Telegram.
2. **Đã tăng timeout sẵn** (30s) trong `bot.py`. Chạy lại `python bot.py`; nếu chỉ chậm tạm thời thì sẽ vào được.
3. **Nếu bị chặn → dùng proxy/VPN.** Điền `telegram.proxy_url` trong `config.json`, ví dụ:
   - `"socks5://127.0.0.1:1080"` (cần `pip install "python-telegram-bot[socks]==20.7"`)
   - `"http://127.0.0.1:8080"`
   Hoặc bật VPN trên máy rồi chạy lại bot.

## Lưu ý về cột nguồn

`pivot.py` đọc theo vị trí cột của file `Detail_order status report_v1.xlsx`:
`CREATED DATE` (cột B), `ORDER ID` (C), `CURRENT POST OFFICE` (P), `CURRENT STATUS` (X).
Nếu API đổi thứ tự cột, chỉnh các hằng `COL_*` ở đầu `pivot.py`.
