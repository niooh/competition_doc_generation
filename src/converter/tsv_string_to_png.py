import subprocess
from pathlib import Path

# 最后一行是备注列，选择 - 或 放弃本组
tsv_data = """
a/787402410	23	23	9	-
b/5309840844	224	99	1	-
sjfkf/3546457755	29	29	8	-
pado/480312606	73	73	4	-
wldp/5229257112	189	80	3	-
apa/4111264575	87	59	7	-
osw/334465154	86	68	5	-
woro/738997985	154	87	2	-
pqo/5494378240	62	62	6	-
"""


def parse_tsv(text: str) -> list[dict[str, str | int]]:
    rows: list[dict[str, str | int]] = []
    for line in text.strip().splitlines():
        if not line.strip():
            continue
        parts = line.split("\t")
        if len(parts) < 4:
            continue

        name_id, total, top3, rank = parts[:4]
        remark = parts[4].strip() if len(parts) >= 5 else "-"
        username, user_id = name_id.split("/", 1)

        rows.append({
            "username": username.strip(),
            "id": user_id.strip(),
            "total": total.strip(),
            "top3": top3.strip(),
            "rank": int(rank.strip()),
            "remark": remark if remark else "-",
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
    columns: (1.2cm, 3.5cm, 3.5cm, 2cm, 2cm, 2cm),
    align: center + horizon,
    inset: 6pt,
    rows: 2.5em,
    stroke: none,

    fill: (x, y) => {
      if y > 0 and calc.odd(y) { rgb("#f6f8fa") }
      else { white }
    },

    table.hline(stroke: 0.8pt),

    table.header(
      strong[排名], strong[用户名], strong[ID], strong[全卷总分], strong[三题最高], strong[备注]
    ),

    table.hline(stroke: 0.4pt),
"""

    esc = lambda v: typst_escape(str(v))

    data_lines = []
    for r in rows:
        data_lines.append(f"    [{r['rank']}],")
        data_lines.append(f"    [{(r['username'] if r['rank'] <= 5 else '-')}],")
        data_lines.append(f"    [{esc(r['id'])}],")
        data_lines.append(f"    [{esc(r['total'])}],")
        data_lines.append(f"    [{esc(r['top3'])}],")
        data_lines.append(f"    [{esc(r['remark'])}],")

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
        typ_path.unlink(missing_ok=True)
    except FileNotFoundError:
        print("Error: typst command not found. Please install Typst and add it to PATH.")
        print("Typst file kept for manual compilation.")
    except subprocess.CalledProcessError as e:
        print("Typst compilation failed:")
        print(e.stderr)
        print("Typst file kept for debugging.")


if __name__ == "__main__":
    main()
