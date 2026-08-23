import subprocess
from pathlib import Path

tsv_data = """
为啥都有名/5138191992	167	74	1	-
圆周率末/787402410	90	70	4	-
拾序舟/3280428432	76	63	7	-
肖方刚/607294491	81	43	15	-
skyscp049/4319419149	51	51	13	-
说不定是个负担/334465154	159	73	2	放弃本组
怀氚/5494378240	109	59	8	-
ToloeZe/3416485394	53	53	12	-
一只无害小墨鱼/5309728558	104	64	6	-
自在每刻/-	34	34	18	-
冰SHROOM/4139120213	77	53	11	-
没有三纵的雪降/788992882	79	53	10	-
姬急几季/5081728667	51	51	13	-
啥名字都没了捏/544007547	58	58	9	-
MYC号/4348964485	143	69	5	-
Chengbai_jim/5399964677	49	36	17	-
江卓阳宝宝/3169678109	37	37	16	放弃本组
1sahh/738997985	170	72	3	放弃本组
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
    align: center,
    inset: 6pt,
    stroke: none,

    table.hline(stroke: 0.8pt),

    table.header(
      strong[排名], strong[用户名], strong[ID], strong[全卷总分], strong[三题最高], strong[备注]
    ),

    table.hline(stroke: 0.4pt),
"""

    esc = lambda v: typst_escape(str(v))

    data_lines: list[str] = [
        f"    [{r['rank']}], "
        f"[{(r['username'] if r['rank'] <= 5 else '-')}], "
        f"[{esc(r['id'])}], "
        f"[{esc(r['total'])}], "
        f"[{esc(r['top3'])}], "
        f"[{esc(r['remark'])}],"
        for r in rows
    ]

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
