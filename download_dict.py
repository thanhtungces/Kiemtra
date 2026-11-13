"""Tải từ điển tiếng Việt (Hunspell) về dạng danh sách từ thuần UTF-8.

Gộp cả hai bảng chính tả: DauMoi (kiểu mới, chuẩn hiện hành) và DauCu (kiểu cũ)
để tool chấp nhận cả hai cách đặt dấu.
"""
import os
import sys
import urllib.request

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE = "https://raw.githubusercontent.com/1ec5/hunspell-vi/main/dictionaries/"
SOURCES = [
    BASE + "vi-DauMoi.dic",
    BASE + "vi-DauCu.dic",
]
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vi_words.txt")


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "spellcheck-vi/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


words = set()
for url in SOURCES:
    try:
        data = fetch(url)
    except Exception as e:
        print(f"[skip] {url}: {e}", file=sys.stderr)
        continue
    text = data.decode("utf-8", errors="replace")
    lines = text.splitlines()
    # Hunspell .dic: dòng đầu là số từ, mỗi dòng "từ/cờ"; bỏ phần /cờ
    if lines and lines[0].strip().isdigit():
        lines = lines[1:]
    n = 0
    for line in lines:
        w = line.split("/", 1)[0].strip()
        if w:
            words.add(w.lower())
            n += 1
    print(f"OK {url}: +{n} từ")

with open(OUT, "w", encoding="utf-8") as f:
    for w in sorted(words):
        f.write(w + "\n")

print(f"Tổng cộng {len(words)} từ -> {OUT}")
