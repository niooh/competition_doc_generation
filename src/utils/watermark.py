import tempfile
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import gray

from src.config import YEAR
from src.utils.fonts import get_cjk_font


def _create_watermark_page(year: int, page_width: float, page_height: float) -> str:
    """生成单页灰色倾斜水印，返回临时文件路径。"""
    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    c = canvas.Canvas(tmp.name, pagesize=(page_width, page_height))
    c.setFillColor(gray, alpha=0.15)
    c.setFont(get_cjk_font(), 72)
    c.translate(page_width / 2, page_height / 2)
    c.rotate(45)
    c.drawCentredString(0, 0, f"化吧吧赛 {year}")
    c.save()
    return tmp.name


def add_watermark(input_pdf: str, output_pdf: str, year: int = YEAR) -> None:
    """将水印叠加到 PDF 的所有页面，保存至 output_pdf。"""
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    page0 = reader.pages[0]
    w = float(page0.mediabox.width)
    h = float(page0.mediabox.height)

    wm_path = _create_watermark_page(year, w, h)
    watermark_page = PdfReader(wm_path).pages[0]

    for page in reader.pages:
        page.merge_page(watermark_page, over=True)
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    # 清理临时水印文件
    import os
    try:
        os.unlink(wm_path)
    except OSError:
        pass
