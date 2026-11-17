# spellcheck-vi — Kiểm tra lỗi chính tả và từ vô nghĩa tiếng Việt

Tool dòng lệnh (CLI) viết bằng Python thuần, kiểm tra lỗi chính tả và các từ vô nghĩa
(không có trong từ điển) trong văn bản tiếng Việt. Không cần cài thêm thư viện nào.

## Cách hoạt động

- Tách văn bản thành các **âm tiết** (từ), bỏ qua số, mã, ngày tháng, địa chỉ, URL.
- Đối chiếu từng âm tiết với **từ điển tiếng Việt** (`vi_words.txt`, ~6.700 âm tiết,
  lấy từ bộ từ điển Hunspell chính thống dùng cho Firefox/LibreOffice, gộp cả kiểu
  dấu mới và kiểu dấu cũ).
- Âm tiết **không có trong từ điển** được đánh dấu `[!]` — đó là ứng viên **sai chính
  tả** hoặc **từ vô nghĩa**.
- Tuỳ chọn `--suggest` đưa ra **gợi ý sửa** (cùng cách đọc khi bỏ dấu, hoặc gần giống
  theo khoảng cách ký tự).

> Lưu ý: tiếng Việt viết rời từng âm tiết nên tool kiểm tra ở mức âm tiết. Các từ tiếng
> nước ngoài (Anh, Pháp…) và tên riêng không có trong từ điển tiếng Việt cũng sẽ được
> liệt kê để bạn tự cân nhắc.

## Yêu cầu

- Python 3.8 trở lên (chỉ dùng thư viện chuẩn).

## Các file

| File | Vai trò |
|---|---|
| `spellcheck_vi.py` | Script chính kiểm tra chính tả |
| `vi_words.txt` | Từ điển âm tiết tiếng Việt (dữ liệu) |
| `download_dict.py` | Tải lại/cập nhật từ điển từ nguồn Hunspell |
| `sample.txt` | Văn bản mẫu để thử |

## Cách dùng

### 1. Kiểm tra văn bản nhập trực tiếp

```bash
python spellcheck_vi.py "Hôm nay trời rất đẹp"
python spellcheck_vi.py -s "Tôi muâ mua rau"
```

### 2. Kiểm tra từ file

```bash
python spellcheck_vi.py -f vanban.txt
python spellcheck_vi.py -f vanban.txt --suggest
```

### 3. Đọc từ stdin (pipe)

```bash
type vanban.txt | python spellcheck_vi.py
echo "Nội dung cần kiểm" | python spellcheck_vi.py -s
```

### Các tuỳ chọn

| Tuỳ chọn | Ý nghĩa |
|---|---|
| `-f, --file FILE` | Kiểm tra nội dung file văn bản (UTF-8) |
| `-s, --suggest` | Bật gợi ý từ thay thế gần giống |
| `--dict FILE` | Dùng file từ điển khác (mặc định `vi_words.txt`) |
| `-q, --quiet` | Chỉ in các từ nghi vấn, mỗi từ một dòng (tiện cho script khác) |
| `--no-color` | Tắt màu đầu ra |

## Ví dụ kết quả

```text
[!] dòng 3, cột 34: "chinhs"  -> gợi ý: chính, chinh
[!] dòng 3, cột 46: "muâ"     -> gợi ý: mua, mùa, múa, mưa, mứa
[!] dòng 4, cột 30: "lkjhgf"  -> gợi ý: không có gợi ý

Kiểm tra 53 từ; phát hiện 4 từ nghi vấn (sai chính tả hoặc vô nghĩa).
```

Mã trả về: `0` khi chạy thành công (kể cả có phát hiện lỗi), `2` khi có lỗi về file/từ điển.

## Cập nhật từ điển

```bash
python download_dict.py
```

Script sẽ tải lại từ điển mới nhất từ kho
[hunspell-vi](https://github.com/1ec5/hunspell-vi) (cần kết nối mạng) và ghi đè
`vi_words.txt`.

---

# nd30_formatter — Căn chỉnh lề & thể thức Word theo Nghị định 30/2020/NĐ-CP

Tool xử lý file Word (`.docx`) tự động căn chỉnh khổ giấy A4, lề trang, phông chữ `Times New Roman`, giãn dòng, thụt đầu dòng, đánh số trang và tạo file mẫu chuẩn theo **Phụ lục I – Nghị định 30/2020/NĐ-CP**.

## Cài đặt

```bash
pip install -r requirements.txt
```

## Cách dùng

```bash
# 1. Căn chỉnh 1 file Word (tạo file <tên>_nd30.docx)
python nd30_formatter.py vanban.docx --page-numbers

# 2. Căn chỉnh và chỉ định file đầu ra + tùy chỉnh lề (mm)
python nd30_formatter.py vanban.docx -o vanban_chuan.docx --top 20 --bottom 20 --left 30 --right 15

# 3. Căn chỉnh hàng loạt tất cả file .docx trong một thư mục
python nd30_formatter.py ./thu_muc_van_ban -o ./da_chuan_hoa --page-numbers

# 4. Kiểm tra file có đạt chuẩn Nghị định 30/2020 hay không
python nd30_formatter.py vanban.docx --check

# 5. Tạo file Word mẫu chuẩn thể thức NĐ 30/2020
python nd30_formatter.py --create-template mau_cong_van.docx --doc-type "CÔNG VĂN"
```

