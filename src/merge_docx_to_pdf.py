import shutil
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from src.config import DOC_DIR, DIST_DIR, TEMP_DIR, YEAR
from src.converter.docx_to_pdf import convert_batch
from src.utils.pdf_extras import (
    create_cover, create_toc, merge_pdfs, get_doc_title, add_page_numbers,
)
from src.utils.watermark import add_watermark


def merge_docx_to_pdf(input_dir: Path, output_pdf: Path) -> None:
    """转换所有 docx -> 合并 -> 加封面/目录/页码/水印。"""
    input_dir = input_dir.resolve()
    output_pdf = output_pdf.resolve()
    TEMP_DIR.mkdir(parents=True, exist_ok=True)

    docx_files = sorted(input_dir.glob("*.docx"))
    if not docx_files:
        print("No .docx files found.")
        return

    pdf_files = convert_batch(docx_files, TEMP_DIR)
    if not pdf_files:
        print("No files converted.")
        return

    # 合并试题 PDF，记录各文档在正文中的起始页码
    doc_info = []
    writer = PdfWriter()
    for fpath, docx_path in zip(pdf_files, docx_files):
        start_page = len(writer.pages) + 1
        reader = PdfReader(str(fpath))
        for page in reader.pages:
            writer.add_page(page)
        doc_info.append((get_doc_title(docx_path), start_page))

    body_pdf = TEMP_DIR / "body.pdf"
    with open(body_pdf, "wb") as f:
        writer.write(f)

    # 正文添加页码（从 1 开始，底部居中）
    body_numbered = TEMP_DIR / "body_numbered.pdf"
    add_page_numbers(str(body_pdf), str(body_numbered))

    cover_pdf = create_cover(YEAR)
    toc_pdf = TEMP_DIR / "toc.pdf"
    toc_pdf_path = create_toc(doc_info, toc_pdf)  # 现在只返回路径

    # 合并封面、目录、带页码的正文，再加水印
    merged_no_wm = TEMP_DIR / "merged_no_watermark.pdf"
    merge_pdfs([cover_pdf, toc_pdf_path, body_numbered], merged_no_wm)
    add_watermark(str(merged_no_wm), str(output_pdf), YEAR)

    # 添加 PDF 书签大纲
    reader = PdfReader(str(output_pdf))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    # 页面索引：0=封面，1~toc_pages=目录，正文从 toc_pages+1 开始
    toc_pages = len(PdfReader(str(toc_pdf_path)).pages)
    body_start_index = 1 + toc_pages  # 封面 + 目录页数

    # 添加一个父书签指向目录页
    toc_bookmark = writer.add_outline_item("目录", page_number=1)  # 目录页在索引 1

    for title, body_page in doc_info:
        target_page = body_start_index + body_page - 1
        writer.add_outline_item(title, target_page, parent=toc_bookmark)

    with open(str(output_pdf), "wb") as f:
        writer.write(f)

    shutil.rmtree(TEMP_DIR, ignore_errors=True)
    print(f"✓ Output saved to {output_pdf}")

def main() -> None:
    DIST_DIR.mkdir(exist_ok=True)
    merge_docx_to_pdf(DOC_DIR, DIST_DIR / "merged.pdf")


if __name__ == "__main__":
    main()
