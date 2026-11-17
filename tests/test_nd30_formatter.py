from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Mm, Pt

from nd30_formatter import (
    A4_HEIGHT_MM,
    A4_WIDTH_MM,
    FONT_NAME,
    ND30Config,
    _emu_to_mm,
    check_document,
    create_template,
    format_document,
    main,
)


def test_config_validation_defaults():
    cfg = ND30Config()
    assert cfg.validate() == []


def test_config_validation_out_of_range():
    cfg = ND30Config(top_mm=10.0, left_mm=15.0, body_font_size_pt=11.0, line_spacing=2.0)
    warnings = cfg.validate()
    assert len(warnings) == 4


def test_check_and_format_non_compliant_document(tmp_path: Path):
    src = tmp_path / "bad.docx"
    doc = Document()
    sec = doc.sections[0]
    # Khổ Letter và lề sai chuẩn
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11.0)
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(0.5)
    sec.right_margin = Inches(0.5)

    p_qh = doc.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    p_qh.runs[0].font.name = "Arial"

    p_body = doc.add_paragraph("Đây là nội dung văn bản hành chính cần căn chỉnh theo Nghị định 30.")
    p_body.runs[0].font.name = "Calibri"
    p_body.runs[0].font.size = Pt(11)

    table = doc.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "Cột 1"
    table.cell(0, 1).text = "Cột 2"

    doc.save(str(src))

    # Kiểm tra trước khi sửa -> phải báo lỗi
    before = check_document(src)
    assert not before.passed
    assert any("A4" in issue for issue in before.issues)
    assert any("Arial" in issue or "Calibri" in issue for issue in before.issues)

    # Định dạng lại
    out = tmp_path / "good.docx"
    cfg = ND30Config(add_page_numbers=True)
    format_document(src, out, cfg)

    after = check_document(out)
    assert after.passed, after.format_report()

    formatted_doc = Document(str(out))
    s = formatted_doc.sections[0]
    assert abs(_emu_to_mm(s.page_width) - A4_WIDTH_MM) < 0.5
    assert abs(_emu_to_mm(s.page_height) - A4_HEIGHT_MM) < 0.5
    assert abs(_emu_to_mm(s.top_margin) - 20.0) < 0.5
    assert abs(_emu_to_mm(s.bottom_margin) - 20.0) < 0.5
    assert abs(_emu_to_mm(s.left_margin) - 30.0) < 0.5
    assert abs(_emu_to_mm(s.right_margin) - 15.0) < 0.5
    assert s.different_first_page_header_footer is True

    # Quốc hiệu căn giữa, cỡ 12, đậm
    assert formatted_doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert formatted_doc.paragraphs[0].runs[0].font.name == FONT_NAME
    assert formatted_doc.paragraphs[0].runs[0].bold is True

    # Nội dung căn đều 2 bên, thụt đầu dòng, Times New Roman 14pt
    body_p = formatted_doc.paragraphs[1]
    assert body_p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
    assert body_p.runs[0].font.name == FONT_NAME
    assert body_p.runs[0].font.size == Pt(14)
    assert body_p.paragraph_format.first_line_indent is not None

    # Bảng không bị thụt đầu dòng
    cell_p = formatted_doc.tables[0].cell(0, 0).paragraphs[0]
    assert cell_p.runs[0].font.name == FONT_NAME
    assert cell_p.paragraph_format.first_line_indent is None


def test_create_templates_pass_check(tmp_path: Path):
    cv_path = tmp_path / "mau_cong_van.docx"
    qd_path = tmp_path / "mau_quyet_dinh.docx"

    create_template(cv_path, doc_type="CÔNG VĂN")
    create_template(qd_path, doc_type="QUYẾT ĐỊNH")

    assert check_document(cv_path).passed
    assert check_document(qd_path).passed


def test_cli_batch_and_check(tmp_path: Path):
    tpl = tmp_path / "in" / "doc1.docx"
    tpl.parent.mkdir(parents=True, exist_ok=True)
    create_template(tpl, doc_type="BÁO CÁO")

    # --check trên file chuẩn trả về 0
    assert main([str(tpl), "--check"]) == 0

    # Xử lý thư mục
    out_dir = tmp_path / "out"
    assert main([str(tpl.parent), "-o", str(out_dir)]) == 0
    assert (out_dir / "doc1.docx").exists()
