#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
spellcheck_vi.py — Kiểm tra lỗi chính tả và từ vô nghĩa trong văn bản tiếng Việt.

Cách hoạt động: tách văn bản thành các âm tiết (từ), đối chiếu từng âm tiết với
từ điển tiếng Việt (Hunspell, ~6,7 nghìn âm tiết). Âm tiết không có trong từ điển
bị đánh dấu là "nghi vấn" — có thể là sai chính tả hoặc từ vô nghĩa.

Chỉ dùng thư viện chuẩn Python (3.8+), không cần cài gì thêm.

Ví dụ:
    python spellcheck_vi.py "Hôm nay trời rất đẹp"
    python spellcheck_vi.py -f vanban.txt
    python spellcheck_vi.py -f vanban.txt --suggest
    type vanban.txt | python spellcheck_vi.py
"""

import argparse
import difflib
import os
import re
import sys

# ---------------------------------------------------------------------------
# Bảng rút gọn dấu tiếng Việt -> chữ cái La-tinh cơ bản (để gợi ý từ gần giống)
# ---------------------------------------------------------------------------
_TONE = {
    "á": "a", "à": "a", "ả": "a", "ã": "a", "ạ": "a",
    "ă": "a", "ắ": "a", "ằ": "a", "ẳ": "a", "ẵ": "a", "ặ": "a",
    "â": "a", "ấ": "a", "ầ": "a", "ẩ": "a", "ẫ": "a", "ậ": "a",
    "é": "e", "è": "e", "ẻ": "e", "ẽ": "e", "ẹ": "e",
    "ê": "e", "ế": "e", "ề": "e", "ể": "e", "ễ": "e", "ệ": "e",
    "í": "i", "ì": "i", "ỉ": "i", "ĩ": "i", "ị": "i",
    "ó": "o", "ò": "o", "ỏ": "o", "õ": "o", "ọ": "o",
    "ô": "o", "ố": "o", "ồ": "o", "ổ": "o", "ỗ": "o", "ộ": "o",
    "ơ": "o", "ớ": "o", "ờ": "o", "ở": "o", "ỡ": "o", "ợ": "o",
    "ú": "u", "ù": "u", "ủ": "u", "ũ": "u", "ụ": "u",
    "ư": "u", "ứ": "u", "ừ": "u", "ử": "u", "ữ": "u", "ự": "u",
    "ý": "y", "ỳ": "y", "ỷ": "y", "ỹ": "y", "ỵ": "y",
    "đ": "d",
}


def fold_word(word: str) -> str:
    """Rút gọn dấu: 'mưa' -> 'mua', 'múa' -> 'mua', 'đ' -> 'd'."""
    return "".join(_TONE.get(c, c) for c in word.lower())


# Mẫu nhận diện một "từ": chữ cái/số liên tiếp, cho phép gạch ngang/ngoặc giữa từ.
_TOKEN_RE = re.compile(r"\w+(?:[-']\w+)*", re.UNICODE)


def is_checkable(tok: str) -> bool:
    """Chỉ kiểm tra token có chữ cái và không lẫn chữ số (bỏ qua mã/URL/ngày...)."""
    if not tok:
        return False
    if tok.isdigit():
        return False
    if any(ch.isdigit() for ch in tok):
        return False
    if not any(ch.isalpha() for ch in tok):
        return False
    return True


class SpellChecker:
    def __init__(self, word_file: str):
        self.words = set()
        with open(word_file, "r", encoding="utf-8-sig") as f:
            for line in f:
                w = line.strip().lower()
                if w:
                    self.words.add(w)
        self._sorted = sorted(self.words)
        self._fold_index = {}
        for w in self._sorted:
            self._fold_index.setdefault(fold_word(w), []).append(w)

    @property
    def size(self) -> int:
        return len(self.words)

    def check(self, token: str) -> bool:
        return token.lower() in self.words

    def suggest(self, token: str, limit: int = 5) -> list:
        t = token.lower()
        cands = []
        # 1) Cùng cách đọc sau khi bỏ dấu (lỗi đặt dấu sai).
        cands.extend(self._fold_index.get(fold_word(t), []))
        # 2) Gần giống theo khoảng cách ký tự (lỗi gõ sai chữ).
        cands.extend(difflib.get_close_matches(t, self._sorted, n=limit * 2, cutoff=0.82))
        out, seen = [], set()
        for c in cands:
            if c == t or c in seen:
                continue
            seen.add(c)
            out.append(c)
            if len(out) >= limit:
                break
        return out


def find_suspects(text: str, checker: SpellChecker):
    """Duyệt từng dòng, trả danh sách (số dòng, cột, token)."""
    results = []
    for line_no, line in enumerate(text.splitlines(), start=1):
        for m in _TOKEN_RE.finditer(line):
            tok = m.group(0)
            if not is_checkable(tok) or checker.check(tok):
                continue
            results.append((line_no, m.start() + 1, tok))
    return results


def count_words(text: str) -> int:
    return sum(
        1
        for line in text.splitlines()
        for m in _TOKEN_RE.finditer(line)
        if is_checkable(m.group(0))
    )


def _print_report(text: str, checker: SpellChecker, suggest_flag: bool, color: bool):
    suspects = find_suspects(text, checker)
    total = count_words(text)

    red, green, dim, reset = "", "", "", ""
    if color:
        red, green, dim, reset = "\033[31m", "\033[32m", "\033[2m", "\033[0m"

    for line_no, col, tok in suspects:
        mark = f"{red}[!]{reset}" if color else "[!]"
        extra = ""
        if suggest_flag:
            tip = ", ".join(checker.suggest(tok)) if checker.suggest(tok) else "không có gợi ý"
            extra = f"  {dim}→ gợi ý:{reset} {tip}" if color else f"  -> gợi ý: {tip}"
        print(f'{mark} dòng {line_no}, cột {col}: "{tok}"{extra}')

    print()
    if suspects:
        n = len(suspects)
        summary = (
            f"Kiểm tra {total} từ; phát hiện {n} từ nghi vấn "
            f"(sai chính tả hoặc vô nghĩa)."
        )
        print(f"{red}{summary}{reset}" if color else summary)
    else:
        ok = f"Kiểm tra {total} từ; không phát hiện lỗi chính tả."
        print(f"{green}{ok}{reset}" if color else ok)


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="spellcheck_vi.py",
        description="Kiểm tra lỗi chính tả và từ vô nghĩa trong văn bản tiếng Việt.",
        epilog="Nếu không truyền --file, các tham số còn lại được ghép thành văn bản. "
               "Nếu cũng không có, đọc từ stdin.",
    )
    ap.add_argument("text", nargs="*", help="Văn bản cần kiểm tra (ghép các tham số lại).")
    ap.add_argument("-f", "--file", help="Đường dẫn file văn bản (UTF-8) cần kiểm tra.")
    ap.add_argument("--dict", default=None,
                    help="File từ điển (mặc định: vi_words.txt nằm cạnh script).")
    ap.add_argument("-s", "--suggest", action="store_true",
                    help="Bật gợi ý từ thay thế gần giống.")
    ap.add_argument("--no-color", action="store_true", help="Tắt màu đầu ra.")
    ap.add_argument("-q", "--quiet", action="store_true",
                    help="Chỉ in các từ nghi vấn, mỗi từ một dòng.")
    args = ap.parse_args(argv)

    dict_path = args.dict or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "vi_words.txt"
    )
    if not os.path.exists(dict_path):
        print(f"LỖI: không tìm thấy từ điển '{dict_path}'.", file=sys.stderr)
        print("Chạy lại download_dict.py để tải từ điển.", file=sys.stderr)
        return 2

    checker = SpellChecker(dict_path)

    # Thu thập văn bản đầu vào
    if args.file:
        try:
            with open(args.file, "r", encoding="utf-8-sig") as f:
                text = f.read()
        except OSError as e:
            print(f"LỖI: không đọc được file '{args.file}': {e}", file=sys.stderr)
            return 2
    elif args.text:
        text = " ".join(args.text)
    else:
        try:
            # Đọc dạng nhị phân để tránh lệch mã hoá khi pipe từ PowerShell/cmd.
            data = getattr(sys.stdin, "buffer", sys.stdin).read()
            text = data.decode("utf-8", errors="replace") if isinstance(data, bytes) else data
        except OSError as e:
            print(f"LỖI: không đọc được stdin: {e}", file=sys.stderr)
            return 2

    if not text.strip():
        print("Không có văn bản để kiểm tra.", file=sys.stderr)
        return 0

    suspects = find_suspects(text, checker)

    if args.quiet:
        for _, _, tok in suspects:
            print(tok)
        return 0

    color = not args.no_color and sys.stdout.isatty()
    _print_report(text, checker, args.suggest, color)
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    raise SystemExit(main())
