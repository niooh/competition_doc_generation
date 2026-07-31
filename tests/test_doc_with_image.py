import tempfile
from pathlib import Path

import pytest
from pypdf import PdfReader
from PIL import Image

from src.config import PERIODIC_TABLE_IMG
from src.utils.pdf_extras import create_toc

"""
测试 create_toc 在提供图片时的输出效果。
"""

@pytest.fixture
def test_image_path():
    """如果项目中没有现成图片，就动态生成一张简单的测试 PNG。"""
    if PERIODIC_TABLE_IMG.is_file():
        return PERIODIC_TABLE_IMG
    # 降级：生成一张 400x300 的蓝色占位图
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    img = Image.new("RGB", (400, 300), color="blue")
    img.save(tmp.name)
    return Path(tmp.name)


def test_create_toc_with_image(test_image_path, tmp_path):
    """验证目录+图片生成 PDF 且页面正常。"""
    entries = [("元素推断、配合物结构和磁性", 1)]
    out_pdf = tmp_path / "output.pdf"

    result = create_toc(entries, image_path=test_image_path, output_path=out_pdf)

    # 检查返回值是路径
    assert result == str(out_pdf)
    assert out_pdf.is_file()

    # 检查 PDF 至少有 1 页
    reader = PdfReader(str(out_pdf))
    assert len(reader.pages) >= 1

    # 可选：检查第一页尺寸为 A4
    page = reader.pages[0]
    width = float(page.mediabox.width)
    height = float(page.mediabox.height)
    assert abs(width - 595.27) < 1 and abs(height - 841.89) < 1  # A4 in points


def test_create_toc_without_image(tmp_path):
    """验证不提供图片时，目录仍正常生成。"""
    entries = [("测试题", 2)]
    out_pdf = tmp_path / "no_image.pdf"
    result = create_toc(entries, output_path=out_pdf)

    assert out_pdf.is_file()
    reader = PdfReader(str(result))
    assert len(reader.pages) >= 1
