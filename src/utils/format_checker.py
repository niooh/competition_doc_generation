"""
Check docx structural issues: numbering, scores, green comments.
"""

from argparse import ArgumentParser
from pathlib import Path
import re

from docx import Document
from docx.oxml.ns import qn


# ANSI colors
CYAN = "\033[36m"
YELLOW = "\033[33m"
RED = "\033[31m"
GREEN = "\033[32m"
RESET = "\033[0m"


def get_paragraph_level(para) -> int:
    num_pr = para._p.xpath('.//w:numPr')
    if not num_pr:
        return -1
    ilvl = num_pr[0].find(qn('w:ilvl'))
    if ilvl is None:
        return -1
    try:
        return int(ilvl.get(qn('w:val')))
    except ValueError:
        return -1


def get_paragraph_number(para) -> str | None:
    text = para.text.strip()
    m = re.match(r'^(\d+(?:-\d+)?)[.、）]', text)
    if m:
        return m.group(1)
    return None


def extract_end_score(line: str) -> int | None:
    s = line.rstrip()
    if not s:
        return None
    last_left = s.rfind('（')
    if last_left == -1:
        return None
    right = s.find('）', last_left)
    if right == -1:
        return None
    content = s[last_left + 1: right]
    fen_pos = content.find('分')
    if fen_pos <= 0:
        return None
    num_str = ''
    i = fen_pos - 1
    while i >= 0 and content[i].isdigit():
        num_str = content[i] + num_str
        i -= 1
    if not num_str:
        return None
    if right != len(s) - 1:
        return None
    return int(num_str)


def has_green_text(para) -> bool:
    for run in para.runs:
        color = run.font.color
        if color is not None and color.rgb is not None:
            rgb = (color.rgb[0], color.rgb[1], color.rgb[2])
            if rgb in ((0, 192, 0), (0, 255, 0)):
                return True
    return False


def check_file(docx_path: Path) -> tuple[list[dict], list[dict]]:
    doc = Document(docx_path)
    errors = []
    warnings = []

    expected_next = {}
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        if idx == 0:
            if '分' not in text:
                warnings.append({"para": idx, "msg": "title missing total score hint (e.g., 25分)"})
            if get_paragraph_level(para) != -1:
                warnings.append({"para": idx, "msg": "title should not be numbered"})
            continue

        if has_green_text(para):
            errors.append({"para": idx, "msg": "green comment text found, remove it"})

        level = get_paragraph_level(para)
        if level == -1:
            continue

        num = get_paragraph_number(para)
        if num is None:
            continue

        score = extract_end_score(text)
        if score is None:
            warnings.append({"para": idx, "msg": "missing score at end (e.g., （X分）)"})
        else:
            if score <= 0:
                errors.append({"para": idx, "msg": f"score must be positive, got {score}"})
            if score > 50:
                warnings.append({"para": idx, "msg": f"score unusually high ({score})"})

        parts = num.split('-')
        try:
            if len(parts) == 1:
                cur = int(parts[0])
            elif len(parts) == 2:
                cur = int(parts[-1])
            else:
                continue
        except ValueError:
            continue

        if level not in expected_next:
            if cur != 1:
                warnings.append({"para": idx, "msg": f"first number at this level should be 1, got {cur}"})
            expected_next[level] = 2
        else:
            if cur != expected_next[level]:
                warnings.append({"para": idx, "msg": f"numbering gap: expected {expected_next[level]}, got {cur}"})
            expected_next[level] = cur + 1

        if len(parts) == 2 and int(parts[0]) == 0:
            warnings.append({"para": idx, "msg": f"sub-number has zero parent: {num}"})

    return errors, warnings


def colorize(text: str, color: str) -> str:
    return f"{color}{text}{RESET}"


def print_report(path: Path, errors: list[dict], warnings: list[dict]):
    print(f"\n{colorize(path.name, CYAN)}")
    if not errors and not warnings:
        print(f"  {colorize('OK', GREEN)}")
        return
    for item in errors:
        print(f"  {colorize('✗', RED)} para {item['para']+1}: {item['msg']}")
    for item in warnings:
        print(f"  {colorize('⚠', YELLOW)} para {item['para']+1}: {item['msg']}")


def main() -> None:
    parser = ArgumentParser(description="Check docx structure: numbering and scores")
    parser.add_argument("--dir", type=Path, help="Directory containing docx files")
    parser.add_argument("--file", type=Path, help="Single docx file")
    args = parser.parse_args()

    if args.file:
        files = [args.file]
    elif args.dir:
        files = list(args.dir.glob("*.docx"))
    else:
        root = Path(__file__).resolve().parent.parent.parent
        files = list((root / "doc").glob("*.docx"))

    if not files:
        print("No .docx files found.")
        return

    total_errors = 0
    for p in files:
        err, warn = check_file(p)
        print_report(p, err, warn)
        total_errors += len(err)

    if total_errors:
        print(f"\n{colorize(f'✗ Found {total_errors} error(s). Fix before proceeding.', RED)}")
    else:
        print(f"\n{colorize('✓ All files passed (warnings are advisory).', GREEN)}")

if __name__ == "__main__":
    main()
