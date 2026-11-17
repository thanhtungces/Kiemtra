#!/usr/bin/env python3
"""
Tool căn chỉnh lề và chuẩn hóa thể thức văn bản Word (.docx)
theo Nghị định 30/2020/NĐ-CP ngày 05/03/2020 của Chính phủ về công tác văn thư.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

# Hằng số chuẩn theo Phụ lục I - Nghị định 30/2020/NĐ-CP
A4_WIDTH_MM = 210.0
A4_HEIGHT_MM = 297.0

MARGIN_RANGES_MM = {
    "top": (20.0, 25.0),
    "bottom": (20.0, 25.0),
    "left": (30.0, 35.0),
    "right": (15.0, 20.0),
}

BODY_FONT_SIZE_RANGE_PT = (13.0, 14.0)
LINE_SPACING_RANGE = (1.0, 1.5)
FIRST_LINE_INDENT_RANGE_CM = (1.0, 1.27)
MIN_SPACE_AFTER_PT = 6.0

FONT_NAME = "Times New Roman"
BLACK_RGB = RGBColor(0, 0, 0)

# Danh sách tên loại văn bản hành chính phổ biến theo Điều 7 NĐ 30/2020/NĐ-CP
ADMIN_DOC_TYPES = {
    "NGHỊ QUYẾT",
    "QUYẾT ĐỊNH",
    "CHỈ THỊ",
    "QUY CHẾ",
    "QUY ĐỊNH",
    "THÔNG CÁO",
    "THÔNG BÁO",
    "HƯỚNG DẪN",
    "CHƯƠNG TRÌNH",
    "KẾ HOẠCH",
    "PHƯƠNG ÁN",
    "ĐỀ ÁN",
    "DỰ ÁN",
    "BÁO CÁO",
    "BIÊN BẢN",
    "TỜ TRÌNH",
    "HỢP ĐỒNG",
    "CÔNG VĂN",
    "CÔNG ĐIỆN",
    "BẢN GHI NHỚ",
    "BẢN THỎA THUẬN",
    "GIẤY ỦY QUYỀN",
    "GIẤY MỜI",
    "GIẤY GIỚI THIỆU",
    "GIẤY NGHỈ PHÉP",
    "PHIẾU GỬI",
    "PHIẾU CHUYỂN",
    "PHIẾU BÁO",
    "THƯ CÔNG",
}


@dataclass
class ND30Config:
    """Cấu hình định dạng theo Nghị định 30/2020/NĐ-CP."""

    top_mm: float = 20.0
    bottom_mm: float = 20.0
    left_mm: float = 30.0
    right_mm: float = 15.0
    header_distance_mm: float = 12.0
    footer_distance_mm: float = 12.0

    font_name: str = FONT_NAME
    body_font_size_pt: float = 14.0
    line_spacing: float = 1.15
    space_after_pt: float = 6.0
    first_line_indent_cm: float = 1.27

    add_page_numbers: bool = False
    page_number_size_pt: float = 13.0
    keep_landscape: bool = True
    format_body: bool = True

    def validate(self) -> List[str]:
        """Kiểm tra các thông số cấu hình có nằm trong ngưỡng NĐ 30/2020 hay không."""
        warnings: List[str] = []
        for side, value in [
            ("top", self.top_mm),
            ("bottom", self.bottom_mm),
            ("left", self.left_mm),
            ("right", self.right_mm),
        ]:
            lo, hi = MARGIN_RANGES_MM[side]
            if not (lo <= value <= hi):
                warnings.append(
                    f"Lề {side} ({value} mm) nằm ngoài khoảng chuẩn NĐ 30 ({lo}–{hi} mm)."
                )

        if not (BODY_FONT_SIZE_RANGE_PT[0] <= self.body_font_size_pt <= BODY_FONT_SIZE_RANGE_PT[1]):
            warnings.append(
                f"Cỡ chữ nội dung ({self.body_font_size_pt} pt) nằm ngoài khoảng chuẩn "
                f"({BODY_FONT_SIZE_RANGE_PT[0]}–{BODY_FONT_SIZE_RANGE_PT[1]} pt)."
            )

        if not (LINE_SPACING_RANGE[0] <= self.line_spacing <= LINE_SPACING_RANGE[1]):
            warnings.append(
                f"Giãn dòng ({self.line_spacing}) nằm ngoài khoảng chuẩn "
                f"({LINE_SPACING_RANGE[0]}–{LINE_SPACING_RANGE[1]})."
            )

        if not (FIRST_LINE_INDENT_RANGE_CM[0] <= self.first_line_indent_cm <= FIRST_LINE_INDENT_RANGE_CM[1]):
            warnings.append(
                f"Thụt đầu dòng ({self.first_line_indent_cm} cm) nằm ngoài khoảng chuẩn "
                f"({FIRST_LINE_INDENT_RANGE_CM[0]}–{FIRST_LINE_INDENT_RANGE_CM[1]} cm)."
            )

        if self.space_after_pt < MIN_SPACE_AFTER_PT:
            warnings.append(
                f"Khoảng cách đoạn ({self.space_after_pt} pt) nhỏ hơn mức tối thiểu {MIN_SPACE_AFTER_PT} pt."
            )

        return warnings


@dataclass
class CheckResult:
    """Kết quả kiểm tra thể thức một file Word."""

    file_path: str
    passed: bool
    issues: List[str] = field(default_factory=list)
    info: List[str] = field(default_factory=list)

    def format_report(self) -> str:
        status = "ĐẠT CHUẨN NĐ 30/2020" if self.passed else "CHƯA ĐẠT CHUẨN NĐ 30/2020"
        lines = [f"=== Kiểm tra: {self.file_path} [{status}] ==="]
        if self.info:
            lines.append("Thông tin:")
            lines.extend(f"  - {item}" for item in self.info)
        if self.issues:
            lines.append(f"Các lỗi / điểm chưa phù hợp ({len(self.issues)}):")
            lines.extend(f"  [!] {issue}" for issue in self.issues)
        else:
            lines.append("Không phát hiện lỗi định dạng nào.")
        return "\n".join(lines)


def _emu_to_mm(emu: Optional[int]) -> float:
    if emu is None:
        return 0.0
    return round(emu / 36000.0, 2)


def _set_run_font(run, font_name: str = FONT_NAME, size_pt: Optional[float] = None,
                  bold: Optional[bool] = None, italic: Optional[bool] = None) -> None:
    """Thiết lập font đầy đủ (bao gồm w:rFonts cho Unicode tiếng Việt) và màu đen."""
    run.font.name = font_name
    run.font.color.rgb = BLACK_RGB
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)


def _set_style_font(style, font_name: str = FONT_NAME, size_pt: Optional[float] = None) -> None:
    """Thiết lập font mặc định cho một Style trong Word."""
    style.font.name = font_name
    style.font.color.rgb = BLACK_RGB
    if size_pt is not None:
        style.font.size = Pt(size_pt)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)


def _add_page_number_to_header(section, font_name: str = FONT_NAME, size_pt: float = 13.0) -> None:
    """
    Đánh số trang ở giữa phía trên đầu trang, không hiển thị ở trang đầu tiên
    (theo Mục 6 Phần I Phụ lục I Nghị định 30/2020/NĐ-CP).
    """
    section.different_first_page_header_footer = True
    header = section.header
    header.is_linked_to_previous = False

    if header.paragraphs:
        p = header.paragraphs[0]
        p.text = ""
    else:
        p = header.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Cm(0)

    run_begin = p.add_run()
    _set_run_font(run_begin, font_name=font_name, size_pt=size_pt, bold=False, italic=False)
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w"))
    run_begin._r.append(fldChar1)

    run_instr = p.add_run()
    _set_run_font(run_instr, font_name=font_name, size_pt=size_pt, bold=False, italic=False)
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls("w"))
    run_instr._r.append(instrText)

    run_sep = p.add_run()
    _set_run_font(run_sep, font_name=font_name, size_pt=size_pt, bold=False, italic=False)
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w"))
    run_sep._r.append(fldChar2)

    run_val = p.add_run("1")
    _set_run_font(run_val, font_name=font_name, size_pt=size_pt, bold=False, italic=False)

    run_end = p.add_run()
    _set_run_font(run_end, font_name=font_name, size_pt=size_pt, bold=False, italic=False)
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w"))
    run_end._r.append(fldChar3)


def _is_list_paragraph(paragraph) -> bool:
    """Kiểm tra đoạn văn có phải là danh sách gạch đầu dòng / đánh số tự động không."""
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is not None and pPr.find(qn("w:numPr")) is not None:
        return True
    text = paragraph.text.strip()
    if re.match(r"^([-•+*]|(\d+[\.\)])|([a-zA-ZđĐ][\.\)]))\s+", text):
        return True
    return False


def _classify_paragraph(paragraph) -> str:
    """Nhận diện sơ bộ loại thành phần trong văn bản hành chính."""
    text = paragraph.text.strip()
    if not text:
        return "empty"

    upper = text.upper()
    if "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in upper or "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM" in upper:
        return "quoc_hieu"
    if text.lower().startswith("độc lập - tự do - hạnh phúc"):
        return "tieu_ngu"
    if upper in ADMIN_DOC_TYPES:
        return "ten_loai"
    if text.lower().startswith("nơi nhận:"):
        return "noi_nhan_title"
    if text.lower().startswith("kính gửi:"):
        return "kinh_gui"
    if _is_list_paragraph(paragraph):
        return "list"

    # Giữ nguyên các tiêu đề đã căn giữa hoặc căn phải (vd: chữ ký, ngày tháng)
    if paragraph.alignment in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
        return "aligned_special"

    return "body"


def _set_paragraph_bottom_border(paragraph, sz: str = "6") -> None:
    """Thêm đường kẻ ngang phía dưới đoạn văn (dùng cho Tiêu ngữ)."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _remove_table_borders(table) -> None:
    """Ẩn toàn bộ viền bảng (dùng cho bảng bố cục Quốc hiệu / Chữ ký)."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        "</w:tblBorders>"
    )
    tblPr.append(borders)


def format_document(
    input_path: str | Path,
    output_path: Optional[str | Path] = None,
    config: Optional[ND30Config] = None,
) -> Path:
    """
    Căn chỉnh lề và chuẩn hóa thể thức file Word (.docx) theo Nghị định 30/2020/NĐ-CP.
    Trả về đường dẫn file đầu ra.
    """
    cfg = config or ND30Config()
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {in_path}")

    out_path = Path(output_path) if output_path else in_path.with_stem(f"{in_path.stem}_nd30")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(str(in_path))

    # 1. Thiết lập khổ giấy A4 và lề trang cho tất cả sections
    for section in doc.sections:
        is_landscape = section.orientation == WD_ORIENT.LANDSCAPE and cfg.keep_landscape
        if is_landscape:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Mm(A4_HEIGHT_MM)
            section.page_height = Mm(A4_WIDTH_MM)
        else:
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Mm(A4_WIDTH_MM)
            section.page_height = Mm(A4_HEIGHT_MM)

        section.top_margin = Mm(cfg.top_mm)
        section.bottom_margin = Mm(cfg.bottom_mm)
        section.left_margin = Mm(cfg.left_mm)
        section.right_margin = Mm(cfg.right_mm)
        section.header_distance = Mm(cfg.header_distance_mm)
        section.footer_distance = Mm(cfg.footer_distance_mm)

        if cfg.add_page_numbers:
            _add_page_number_to_header(
                section,
                font_name=cfg.font_name,
                size_pt=cfg.page_number_size_pt,
            )

    # 2. Chuẩn hóa Style Normal
    if "Normal" in doc.styles:
        normal_style = doc.styles["Normal"]
        _set_style_font(normal_style, font_name=cfg.font_name, size_pt=cfg.body_font_size_pt)

    # 3. Chuẩn hóa các đoạn văn bản chính
    for p in doc.paragraphs:
        role = _classify_paragraph(p)

        if role == "empty":
            for run in p.runs:
                _set_run_font(run, font_name=cfg.font_name)
            continue

        if role == "quoc_hieu":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            for run in p.runs:
                _set_run_font(run, font_name=cfg.font_name, size_pt=12.0, bold=True, italic=False)
            continue

        if role == "tieu_ngu":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            for run in p.runs:
                _set_run_font(run, font_name=cfg.font_name, size_pt=13.0, bold=True, italic=False)
            continue

        if role == "ten_loai":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(4)
            for run in p.runs:
                _set_run_font(run, font_name=cfg.font_name, size_pt=cfg.body_font_size_pt, bold=True, italic=False)
            continue

        if role == "noi_nhan_title":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_after = Pt(3)
            for run in p.runs:
                _set_run_font(run, font_name=cfg.font_name, size_pt=12.0, bold=True, italic=True)
            continue

        if role in ("aligned_special", "kinh_gui"):
            for run in p.runs:
                _set_run_font(run, font_name=cfg.font_name, size_pt=cfg.body_font_size_pt)
            continue

        if role == "list":
            if cfg.format_body:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.line_spacing = cfg.line_spacing
                p.paragraph_format.space_after = Pt(cfg.space_after_pt)
            for run in p.runs:
                _set_run_font(run, font_name=cfg.font_name, size_pt=cfg.body_font_size_pt)
            continue

        # role == "body"
        if cfg.format_body:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(cfg.first_line_indent_cm)
            p.paragraph_format.line_spacing = cfg.line_spacing
            p.paragraph_format.space_after = Pt(cfg.space_after_pt)
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.right_indent = Cm(0)

        for run in p.runs:
            _set_run_font(run, font_name=cfg.font_name, size_pt=cfg.body_font_size_pt)

    # 4. Chuẩn hóa font trong bảng biểu (không áp dụng thụt đầu dòng trong ô)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        _set_run_font(run, font_name=cfg.font_name)

    doc.save(str(out_path))
    return out_path


def check_document(docx_path: str | Path) -> CheckResult:
    """Kiểm tra mức độ tuân thủ Nghị định 30/2020/NĐ-CP của một file Word (.docx)."""
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {path}")

    doc = Document(str(path))
    issues: List[str] = []
    info: List[str] = []

    for idx, section in enumerate(doc.sections, start=1):
        w_mm = _emu_to_mm(section.page_width)
        h_mm = _emu_to_mm(section.page_height)
        top_mm = _emu_to_mm(section.top_margin)
        bottom_mm = _emu_to_mm(section.bottom_margin)
        left_mm = _emu_to_mm(section.left_margin)
        right_mm = _emu_to_mm(section.right_margin)

        info.append(
            f"Section {idx}: Khổ {w_mm}x{h_mm} mm | "
            f"Lề (Trên/Dưới/Trái/Phải): {top_mm}/{bottom_mm}/{left_mm}/{right_mm} mm"
        )

        # Kiểm tra khổ giấy A4 (dung sai ±1.5 mm)
        is_a4_portrait = abs(w_mm - A4_WIDTH_MM) <= 1.5 and abs(h_mm - A4_HEIGHT_MM) <= 1.5
        is_a4_landscape = abs(w_mm - A4_HEIGHT_MM) <= 1.5 and abs(h_mm - A4_WIDTH_MM) <= 1.5
        if not (is_a4_portrait or is_a4_landscape):
            issues.append(
                f"Section {idx}: Khổ giấy {w_mm}x{h_mm} mm không phải chuẩn A4 (210x297 mm)."
            )

        # Kiểm tra lề trang (dung sai ±0.5 mm)
        tol = 0.5
        for side, val in [("top", top_mm), ("bottom", bottom_mm), ("left", left_mm), ("right", right_mm)]:
            lo, hi = MARGIN_RANGES_MM[side]
            if not (lo - tol <= val <= hi + tol):
                issues.append(
                    f"Section {idx}: Lề {side} = {val} mm không nằm trong chuẩn {lo}–{hi} mm."
                )

    # Kiểm tra font chữ các run
    non_tnr_fonts = set()
    for p in doc.paragraphs:
        for run in p.runs:
            if run.text.strip() and run.font.name and run.font.name != FONT_NAME:
                non_tnr_fonts.add(run.font.name)

    if non_tnr_fonts:
        issues.append(
            f"Phát hiện font chữ khác {FONT_NAME}: {', '.join(sorted(non_tnr_fonts))}."
        )

    passed = len(issues) == 0
    return CheckResult(file_path=str(path), passed=passed, issues=issues, info=info)


def create_template(
    output_path: str | Path,
    doc_type: str = "CÔNG VĂN",
    config: Optional[ND30Config] = None,
) -> Path:
    """
    Tạo mới một file Word mẫu (.docx) chuẩn thể thức Nghị định 30/2020/NĐ-CP.
    """
    cfg = config or ND30Config(add_page_numbers=True)
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Mm(A4_WIDTH_MM)
    section.page_height = Mm(A4_HEIGHT_MM)
    section.top_margin = Mm(cfg.top_mm)
    section.bottom_margin = Mm(cfg.bottom_mm)
    section.left_margin = Mm(cfg.left_mm)
    section.right_margin = Mm(cfg.right_mm)
    section.header_distance = Mm(cfg.header_distance_mm)
    section.footer_distance = Mm(cfg.footer_distance_mm)

    if cfg.add_page_numbers:
        _add_page_number_to_header(section, font_name=cfg.font_name, size_pt=cfg.page_number_size_pt)

    if "Normal" in doc.styles:
        _set_style_font(doc.styles["Normal"], font_name=cfg.font_name, size_pt=cfg.body_font_size_pt)

    # Bảng phần đầu: Cơ quan ban hành (trái) & Quốc hiệu, Tiêu ngữ (phải)
    header_table = doc.add_table(rows=2, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _remove_table_borders(header_table)

    # Cột trái (khoảng 6.5 cm), cột phải (khoảng 10 cm)
    for row in header_table.rows:
        row.cells[0].width = Cm(6.5)
        row.cells[1].width = Cm(10.0)

    # Ô (0,0): Cơ quan chủ quản & Cơ quan ban hành
    c00 = header_table.cell(0, 0)
    p_cq1 = c00.paragraphs[0]
    p_cq1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cq1.paragraph_format.space_after = Pt(0)
    r = p_cq1.add_run("TÊN CƠ QUAN CHỦ QUẢN\n")
    _set_run_font(r, size_pt=12.0, bold=False)
    r2 = p_cq1.add_run("TÊN CƠ QUAN BAN HÀNH")
    _set_run_font(r2, size_pt=12.0, bold=True)

    # Ô (0,1): Quốc hiệu & Tiêu ngữ
    c01 = header_table.cell(0, 1)
    p_qh = c01.paragraphs[0]
    p_qh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_qh.paragraph_format.space_after = Pt(2)
    r_qh = p_qh.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    _set_run_font(r_qh, size_pt=12.0, bold=True)

    p_tn = c01.add_paragraph()
    p_tn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tn.paragraph_format.space_after = Pt(6)
    r_tn = p_tn.add_run("Độc lập - Tự do - Hạnh phúc")
    _set_run_font(r_tn, size_pt=13.0, bold=True)
    _set_paragraph_bottom_border(p_tn)

    # Ô (1,0): Số, ký hiệu
    c10 = header_table.cell(1, 0)
    p_so = c10.paragraphs[0]
    p_so.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_so.paragraph_format.space_before = Pt(6)
    r_so = p_so.add_run("Số: ....../....-....")
    _set_run_font(r_so, size_pt=13.0, bold=False)

    # Ô (1,1): Địa danh và ngày tháng năm
    c11 = header_table.cell(1, 1)
    p_ngay = c11.paragraphs[0]
    p_ngay.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ngay.paragraph_format.space_before = Pt(6)
    r_ngay = p_ngay.add_run("Hà Nội, ngày ... tháng ... năm 20...")
    _set_run_font(r_ngay, size_pt=13.0, italic=True)

    # Tên loại và trích yếu nội dung
    doc_type_upper = doc_type.strip().upper()
    if doc_type_upper != "CÔNG VĂN":
        p_loai = doc.add_paragraph()
        p_loai.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_loai.paragraph_format.space_before = Pt(12)
        p_loai.paragraph_format.space_after = Pt(2)
        r_loai = p_loai.add_run(doc_type_upper)
        _set_run_font(r_loai, size_pt=cfg.body_font_size_pt, bold=True)

        p_ty = doc.add_paragraph()
        p_ty.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ty.paragraph_format.space_after = Pt(10)
        r_ty = p_ty.add_run("Về việc .......................................................")
        _set_run_font(r_ty, size_pt=cfg.body_font_size_pt, bold=True)
    else:
        p_ty = doc.add_paragraph()
        p_ty.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_ty.paragraph_format.space_before = Pt(10)
        p_ty.paragraph_format.space_after = Pt(6)
        r_ty = p_ty.add_run("V/v .......................................................")
        _set_run_font(r_ty, size_pt=12.0, bold=False)

        p_kg = doc.add_paragraph()
        p_kg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_kg.paragraph_format.space_before = Pt(6)
        p_kg.paragraph_format.space_after = Pt(10)
        r_kg = p_kg.add_run("Kính gửi: .......................................................")
        _set_run_font(r_kg, size_pt=cfg.body_font_size_pt, bold=False)

    # Đoạn nội dung mẫu
    for idx in range(1, 3):
        p_body = doc.add_paragraph()
        p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_body.paragraph_format.first_line_indent = Cm(cfg.first_line_indent_cm)
        p_body.paragraph_format.line_spacing = cfg.line_spacing
        p_body.paragraph_format.space_after = Pt(cfg.space_after_pt)
        r_body = p_body.add_run(
            f"Nội dung đoạn {idx} của văn bản được trình bày bằng phông chữ Times New Roman, "
            f"cỡ chữ {int(cfg.body_font_size_pt)} pt, căn đều hai bên và thụt đầu dòng {cfg.first_line_indent_cm} cm "
            "theo quy định tại Phụ lục I Nghị định 30/2020/NĐ-CP."
        )
        _set_run_font(r_body, size_pt=cfg.body_font_size_pt)

    # Bảng phần cuối: Nơi nhận (trái) & Chữ ký (phải)
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _remove_table_borders(sig_table)
    sig_table.rows[0].cells[0].width = Cm(8.0)
    sig_table.rows[0].cells[1].width = Cm(8.5)

    p_nn = sig_table.cell(0, 0).paragraphs[0]
    p_nn.paragraph_format.space_before = Pt(12)
    r_nn_title = p_nn.add_run("Nơi nhận:\n")
    _set_run_font(r_nn_title, size_pt=12.0, bold=True, italic=True)
    r_nn_list = p_nn.add_run("- Như trên;\n- Lưu: VT, ...")
    _set_run_font(r_nn_list, size_pt=11.0, bold=False, italic=False)

    p_sig = sig_table.cell(0, 1).paragraphs[0]
    p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig.paragraph_format.space_before = Pt(12)
    r_chucvu = p_sig.add_run("QUYỀN HẠN, CHỨC VỤ NGƯỜI KÝ\n\n\n\n")
    _set_run_font(r_chucvu, size_pt=13.0, bold=True)
    r_ten = p_sig.add_run("Họ và tên")
    _set_run_font(r_ten, size_pt=13.0, bold=True)

    doc.save(str(out_path))
    return out_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Tool căn chỉnh lề và chuẩn hóa thể thức Word (.docx) theo Nghị định 30/2020/NĐ-CP."
    )
    parser.add_argument(
        "input",
        nargs="?",
        help="Đường dẫn file .docx hoặc thư mục chứa các file .docx cần xử lý.",
    )
    parser.add_argument(
        "-o", "--output",
        help="Đường dẫn file đầu ra (hoặc thư mục đầu ra khi xử lý hàng loạt).",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Chỉ kiểm tra mức độ tuân thủ NĐ 30/2020 mà không sửa file.",
    )
    parser.add_argument(
        "--create-template",
        metavar="OUTPUT_DOCX",
        help="Tạo file Word mẫu chuẩn Nghị định 30/2020/NĐ-CP.",
    )
    parser.add_argument(
        "--doc-type",
        default="CÔNG VĂN",
        help="Tên loại văn bản khi tạo mẫu (vd: 'CÔNG VĂN', 'QUYẾT ĐỊNH', 'TỜ TRÌNH').",
    )

    # Tùy chỉnh lề (mm)
    parser.add_argument("--top", type=float, default=20.0, help="Lề trên (mm, chuẩn 20–25). Mặc định: 20")
    parser.add_argument("--bottom", type=float, default=20.0, help="Lề dưới (mm, chuẩn 20–25). Mặc định: 20")
    parser.add_argument("--left", type=float, default=30.0, help="Lề trái (mm, chuẩn 30–35). Mặc định: 30")
    parser.add_argument("--right", type=float, default=15.0, help="Lề phải (mm, chuẩn 15–20). Mặc định: 15")

    # Tùy chỉnh đoạn văn & cỡ chữ
    parser.add_argument("--font-size", type=float, default=14.0, help="Cỡ chữ nội dung (pt, chuẩn 13–14). Mặc định: 14")
    parser.add_argument("--line-spacing", type=float, default=1.15, help="Giãn dòng (1.0–1.5). Mặc định: 1.15")
    parser.add_argument("--indent", type=float, default=1.27, help="Thụt đầu dòng (cm, 1.0–1.27). Mặc định: 1.27")
    parser.add_argument("--space-after", type=float, default=6.0, help="Khoảng cách sau đoạn (pt, >=6). Mặc định: 6")
    parser.add_argument(
        "--margins-only",
        action="store_true",
        help="Chỉ căn lề A4 và font chữ, không thay đổi căn đều/thụt dòng nội dung.",
    )
    parser.add_argument(
        "--page-numbers",
        action="store_true",
        help="Thêm đánh số trang chuẩn NĐ 30 (giữa lề trên, ẩn trang 1).",
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    cfg = ND30Config(
        top_mm=args.top,
        bottom_mm=args.bottom,
        left_mm=args.left,
        right_mm=args.right,
        body_font_size_pt=args.font_size,
        line_spacing=args.line_spacing,
        first_line_indent_cm=args.indent,
        space_after_pt=args.space_after,
        add_page_numbers=args.page_numbers,
        format_body=not args.margins_only,
    )

    warnings = cfg.validate()
    for w in warnings:
        print(f"[Cảnh báo cấu hình] {w}", file=sys.stderr)

    if args.create_template:
        out = create_template(args.create_template, doc_type=args.doc_type, config=cfg)
        print(f"Đã tạo file mẫu chuẩn NĐ 30/2020 tại: {out}")
        return 0

    if not args.input:
        parser.print_help()
        return 1

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"Lỗi: Không tìm thấy '{in_path}'", file=sys.stderr)
        return 1

    if in_path.is_dir():
        docx_files = sorted(p for p in in_path.glob("*.docx") if not p.name.startswith("~$"))
        if not docx_files:
            print(f"Không tìm thấy file .docx nào trong '{in_path}'.")
            return 0

        if args.check:
            all_passed = True
            for f in docx_files:
                res = check_document(f)
                print(res.format_report())
                print()
                all_passed = all_passed and res.passed
            return 0 if all_passed else 2

        out_dir = Path(args.output) if args.output else in_path / "formatted_nd30"
        out_dir.mkdir(parents=True, exist_ok=True)
        for f in docx_files:
            out_file = out_dir / f.name
            format_document(f, out_file, cfg)
            print(f"Đã chuẩn hóa: {f.name} -> {out_file}")
        return 0

    # Xử lý 1 file
    if args.check:
        res = check_document(in_path)
        print(res.format_report())
        return 0 if res.passed else 2

    out_file = format_document(in_path, args.output, cfg)
    print(f"Đã chuẩn hóa file theo Nghị định 30/2020: {out_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
