import subprocess
from pathlib import Path

tsv_data = """
AB/3280428432	74	45	2
BC/5868954263	73	45	3
CD/787402410	75	45	1
FX/811285566	69	43	6
GI/3169678109	57	43	8
LP/3447077562	55	35	10
EP/5414608897	69	43	6
QS/313912621	73	45	3
LO/5863880932	67	42	9
BI/16500212	11	11	12
PQ/5873186607	32	32	11
ZP/5410807248	69	45	5
"""

def parse_tsv(text: str) -> list[dict[str, str | int]]:
    rows: list[dict[str, str | int]] = []
    for line in text.strip().splitlines():
        if not line.strip():
            continue
        parts = line.split("\t")
        if len(parts) != 4:
            continue

        name_id, total, top3, rank = parts
        username, user_id = name_id.split("/", 1)

        rows.append({
            "username": username.strip(),
            "id": user_id.strip(),
            "total": total.strip(),
            "top3": top3.strip(),
            "rank": int(rank.strip()),
        })

    rows.sort(key=lambda r: r["rank"])  # type: ignore[return-value]
    return rows


def typst_escape(s: str) -> str:
    escapes: dict[str, str] = {
        "\\": r"\\",
        "#": r"\#",
        "[": r"\[",
        "]": r"\]",
        "*": r"\*",
        "_": r"\_",
        "$": r"\$",
        "~": r"\~",
    }
    return "".join(escapes.get(ch, ch) for ch in s)


def generate_typst(rows: list[dict[str, str | int]]) -> str:
    header = """#set page(width: auto, height: auto, margin: 1em)
#set text(font: ("Source Han Sans SC"), size: 11pt, lang: "zh")

#figure(
  table(
    columns: (1.2cm, 3.5cm, 3.5cm, 2cm, 2cm),
    align: center,
    inset: 6pt,
    stroke: none,

    table.hline(stroke: 0.8pt),

    table.header(
      strong[排名], strong[用户名], strong[ID], strong[全卷总分], strong[三题最高],
    ),

    table.hline(stroke: 0.4pt),
"""

    data_lines: list[str] = []
    for r in rows:
        rank: int = r["rank"]
        username: str = r["username"] if rank <= 5 else "-"

        username_esc: str = typst_escape(username)
        id_esc: str = typst_escape(str(r["id"]))
        total_esc: str = typst_escape(str(r["total"]))
        top3_esc: str = typst_escape(str(r["top3"]))

        data_lines.append(
            f"    [{rank}], [{username_esc}], [{id_esc}], [{total_esc}], [{top3_esc}],"
        )

    footer = """
    table.hline(stroke: 0.8pt),
  ),
)
"""

    return header + "\n".join(data_lines) + footer


def main() -> None:
    rows: list[dict[str, str | int]] = parse_tsv(tsv_data)
    typst_code: str = generate_typst(rows)

    typ_path: Path = Path("output.typ")
    png_path: Path = Path("output.png")

    typ_path.write_text(typst_code, encoding="utf-8")

    try:
        subprocess.run(
            ["typst", "compile", str(typ_path), str(png_path), "--ppi", "500"],
            check=True,
            capture_output=True,
            text=True,
        )
        print(f"PNG generated: {png_path}")
        typ_path.unlink(missing_ok=True)   # 成功后删除临时 .typ 文件
    except FileNotFoundError:
        print("Error: typst command not found. Please install Typst and add it to PATH.")
        print("Typst file kept for manual compilation.")
    except subprocess.CalledProcessError as e:
        print("Typst compilation failed:")
        print(e.stderr)
        print("Typst file kept for debugging.")


if __name__ == "__main__":
    main()
