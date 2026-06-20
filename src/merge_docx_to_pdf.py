import os
import time
import traceback
from pathlib import Path
import pypdf
import pythoncom
import win32com.client as win32

def convert_one(docx_path: str, pdf_path: str) -> bool:
    """
    转换单个 docx -> pdf，每次调用都独立初始化和释放 Word。
    内部自动重试 2 次，应对偶发的 COM 断开。
    """
    max_retries = 2
    for attempt in range(max_retries):
        word = None
        doc = None
        try:
            pythoncom.CoInitialize()
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            # 只读打开
            doc = word.Documents.Open(docx_path, ReadOnly=True, Visible=False)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
            return True
        except Exception as e:
            print(f"  ✗ 尝试 {attempt + 1}/{max_retries} 失败: {e}")
            if attempt == max_retries - 1:
                return False
            time.sleep(1)  # 等待一会儿再重试
        finally:
            try:
                if doc: doc.Close(False)
            except: pass
            try:
                if word: word.Quit()
            except: pass
            pythoncom.CoUninitialize()
    return False

def merge_pdfs(pdf_list, output_path):
    merger = pypdf.PdfWriter()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

def merge_docx_to_pdf(input_dir, output_pdf):
    input_dir = Path(input_dir).resolve()
    output_pdf = Path(output_pdf).resolve()
    temp_dir = output_pdf.parent / "temp_pdf"
    temp_dir.mkdir(exist_ok=True)

    docx_files = sorted(input_dir.glob("*.docx"))
    if not docx_files:
        print("没有找到任何 .docx 文件")
        return

    pdf_files = []
    print(f"开始转换 {len(docx_files)} 个文件…")
    start = time.perf_counter()

    for docx_path in docx_files:
        pdf_path = temp_dir / (docx_path.stem + ".pdf")
        print(f"→ {docx_path.name}")
        if convert_one(str(docx_path), str(pdf_path)):
            pdf_files.append(pdf_path)
            print(f"  ✓ 完成")
        else:
            print(f"  ✗ 最终失败，已跳过")

    elapsed = time.perf_counter() - start
    print(f"\n转换总耗时 {elapsed:.1f} 秒")

    if not pdf_files:
        print("没有成功转换任何文件")
        return

    print(f"合并 {len(pdf_files)} 个 PDF …")
    merge_pdfs([str(p) for p in pdf_files], str(output_pdf))
    print(f"✓ 合并完成: {output_pdf}")

    # 清理临时 PDF
    for p in pdf_files:
        p.unlink()
    temp_dir.rmdir()

def main():
    base = Path(__file__).resolve().parent.parent
    dist_dir = base / "dist"
    dist_dir.mkdir(exist_ok=True)
    merge_docx_to_pdf(base / "doc", dist_dir / "merged.pdf")

if __name__ == "__main__":
    main()