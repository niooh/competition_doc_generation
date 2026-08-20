from pathlib import Path

def get_docx_order(input_dir: Path) -> list[Path]:
    """
    从 order.txt 读取顺序；若无该文件则返回字母排序的列表。
    order.txt 每行一个文件名（可带 .docx 也可不带）。
    """
    order_file = input_dir / "order.txt"
    if not order_file.is_file():
        return sorted(input_dir.glob("*.docx"))

    lines = []
    with open(order_file, "r", encoding="utf-8") as f:
        for line in f:
            name = line.strip()
            if name:
                lines.append(name)

    ordered = []
    all_docx = {p.stem: p for p in input_dir.glob("*.docx")}

    for name in lines:
        stem = Path(name).stem if name.lower().endswith(".docx") else name

        matched = all_docx.get(stem)
        if matched is None:
            # 大小写不敏感匹配
            for s, p in all_docx.items():
                if s.lower() == stem.lower():
                    matched = p
                    break

        if matched is None:
            print(f"⚠ order.txt 中指定的文件不存在: {name}")
        else:
            ordered.append(matched)

    if not ordered:
        print("order.txt 有效条目为空，按文件名排序。")
        return sorted(input_dir.glob("*.docx"))

    return ordered
    