import time
from pathlib import Path
import pythoncom
import win32com.client as win32


def convert_one(docx_path: str, pdf_path: str, max_retries: int = 2) -> bool:
    """单个 docx -> PDF，自动重试，返回成功标志。"""
    for attempt in range(max_retries):
        word = None
        doc = None
        try:
            pythoncom.CoInitialize()
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(docx_path, ReadOnly=True, Visible=False)
            doc.SaveAs(pdf_path, FileFormat=17)   # 17 = wdFormatPDF
            return True
        except Exception as e:
            print(f"  ✗ Attempt {attempt + 1}/{max_retries} failed for {Path(docx_path).name}: {e}")
            if attempt == max_retries - 1:
                return False
            time.sleep(1)
        finally:
            # 确保资源释放
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
    return False


def convert_batch(docx_files: list[Path], output_dir: Path) -> list[Path]:
    """批量转换 docx 到 PDF，返回成功生成的 PDF 文件路径列表。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_files = []
    print(f"Converting {len(docx_files)} files...")
    start = time.perf_counter()
    for docx_path in docx_files:
        pdf_path = output_dir / (docx_path.stem + ".pdf")
        print(f" -> {docx_path.name}")
        if convert_one(str(docx_path), str(pdf_path)):
            pdf_files.append(pdf_path)
            print("  ✓ Done")
        else:
            print("  ✗ Failed")
    elapsed = time.perf_counter() - start
    print(f"Conversion finished in {elapsed:.1f}s")
    return pdf_files
