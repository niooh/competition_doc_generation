"""
合并 docx 试题并生成带封面、目录、水印的 PDF。
"""

import os
import shutil
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from src.config import DOC_DIR, DIST_DIR, TEMP_DIR, YEAR
from src.converter.docx_to_pdf import convert_batch
from src.utils.pdf_extras import create_cover, create_toc, merge_pdfs, get_doc_title
from src.utils.watermark import add_watermark


def merge_docx_to_pdf(input_dir: Path, output_pdf: Path) -> None:
    """主流程：转换所有 docx -> 合并 -> 加封面/目录/水印。"""
    input_dir = input_dir.resolve()
    output_pdf = output_pdf.resolve()
    TEMP_DIR.mkdir(parents=True, exist_ok=True)

    # 收集并排序 docx
    docx_files = sorted(input_dir.glob("*.docx"))
    if not docx_files:
        print("No .docx files found.")
        return

    # 批量转换为 PDF
    pdf_files = convert_batch(docx_files, TEMP_DIR)
    if not pdf_files:
        print("No files converted.")
        return

    # 记录各文档标题及正文起始页码
    doc_info = []
    writer = PdfWriter()
    for fpath, docx_path in zip(pdf_files, docx_files):
        start_page = len(writer.pages) + 1
        reader = PdfReader(str(fpath))
        for page in reader.pages:
            writer.add_page(page)
        title = get_doc_title(docx_path)
        doc_info.append((title, start_page))

    # 写入临时无封面目录的正文 PDF
    body_pdf = TEMP_DIR / "body.pdf"
    with open(body_pdf, "wb") as f:
        writer.write(f)

    # 生成封面
    cover_pdf = create_cover(YEAR)

    # 生成目录（先估算，实际页数不符则重生成）
    assumed_toc_pages = 1
    body_offset_guess = 1 + assumed_toc_pages
    toc_temp = TEMP_DIR / "toc_temp.pdf"
    create_toc(doc_info, body_offset_guess, toc_temp)
    toc_reader = PdfReader(str(toc_temp))
    real_toc_pages = len(toc_reader.pages)
    if real_toc_pages != assumed_toc_pages:
        body_offset = 1 + real_toc_pages
        toc_final = TEMP_DIR / "toc.pdf"
        create_toc(doc_info, body_offset, toc_final)
        try:
            os.unlink(toc_temp)
        except OSError:
            pass
    else:
        toc_final = toc_temp

    # 合并封面 + 目录 + 正文
    merged_no_wm = TEMP_DIR / "merged_no_watermark.pdf"
    merge_pdfs([cover_pdf, toc_final, body_pdf], merged_no_wm)

    # 加水印
    add_watermark(str(merged_no_wm), str(output_pdf), YEAR)

    # 清理临时文件
    shutil.rmtree(TEMP_DIR, ignore_errors=True)

    print(f"✓ Output saved to {output_pdf}")


def main() -> None:
    DIST_DIR.mkdir(exist_ok=True)
    merge_docx_to_pdf(DOC_DIR, DIST_DIR / "merged.pdf")


if __name__ == "__main__":
    main()