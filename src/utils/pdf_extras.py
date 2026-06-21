import os
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter
from docx import Document

from src.config import YEAR, TEMP_DIR
from src.utils.fonts import get_cjk_font


def create_cover(year: int = YEAR, output_path: str | Path = None) -> str:
    """生成封面页，返回临时文件路径（若未指定则使用默认临时路径）。"""
    if output_path is None:
        TEMP_DIR.mkdir(parents=True, exist_ok=True)
        output_path = TEMP_DIR / "cover.pdf"
    else:
        output_path = Path(output_path)

    c = canvas.Canvas(str(output_path), pagesize=A4)
    width, height = A4
    title = f"{year}年暑假化学吧吧赛试题（大学组）"
    c.setFont(get_cjk_font(), 28)
    c.drawCentredString(width / 2, height / 2, title)
    c.save()
    return str(output_path)


def create_toc(entries: list[tuple[str, int]], body_offset: int,
               output_path: str | Path = None) -> str:
    """
    生成目录PDF，返回文件路径。
    entries: [(文档标题, 在正文中的起始页码), ...]
    body_offset: 目录之后正文之前的总页数偏移（封面+目录页数）
    """
    if output_path is None:
        TEMP_DIR.mkdir(parents=True, exist_ok=True)
        output_path = TEMP_DIR / "toc.pdf"
    else:
        output_path = Path(output_path)

    c = canvas.Canvas(str(output_path), pagesize=A4)
    width, height = A4
    left_margin = 40 * mm
    right_margin = width - 40 * mm
    y = height - 30 * mm
    line_height = 7 * mm

    c.setFont(get_cjk_font(), 12)
    for title, body_page in entries:
        final_page = body_page + body_offset - 1
        if y < 30 * mm:
            c.showPage()
            c.setFont(get_cjk_font(), 12)
            y = height - 30 * mm
        c.drawString(left_margin, y, title[:80])
        c.drawRightString(right_margin, y, str(final_page))
        y -= line_height
    c.save()
    return str(output_path)


def merge_pdfs(pdf_paths: list[str | Path], output_path: str | Path) -> None:
    """按顺序合并多个PDF并写入 output_path。"""
    writer = PdfWriter()
    for p in pdf_paths:
        reader = PdfReader(str(p))
        for page in reader.pages:
            writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)


def get_doc_title(docx_path: str | Path) -> str:
    """从 docx 第一段提取标题（失败则返回文件名）。"""
    try:
        doc = Document(str(docx_path))
        if doc.paragraphs:
            return doc.paragraphs[0].text.strip()
    except Exception:
        pass
    return Path(docx_path).stem
