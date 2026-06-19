from pathlib import Path
import random
import re

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from src.get_docx_data import docx_data_to_ls

workbook = openpyxl.Workbook()
default_sheet = workbook.active
workbook.remove(default_sheet)
green_font = Font(color="00C000")
bold_font = Font(bold=True)

"""
- 绿色字体 用于强调：请勿删除这个公式单元格。且它的下方均应通过'下拉填充'产生，而非手填
- Random/123 仅作为 用户名/id数字 示例，可实际调整。
"""

# 从docx自动读取题目信息
# 获取docx中的题目数据
question_data = docx_data_to_ls()

# 处理sheet名：去掉末尾（XX分）+ 替换非法字符
sheet_mapping = []  # 保存 (原docx题目名, 清洗后sheet名)
for raw_title in question_data.keys():
    # 去掉结尾类似（25分）、（30分）的后缀
    clean_title = re.sub(r"（\d+分）$", "", raw_title).strip()
    # 替换在线文档中sheet不允许的字符
    clean_title = re.sub(r"[:：\\/／?？*＊\[\]［］]", "_", clean_title)
    sheet_mapping.append((raw_title, clean_title))

# 清洗后的sheet名列表
sheet_names = [item[1] for item in sheet_mapping]
num_sheets = len(sheet_names)

first_sheet = workbook.create_sheet("答卷基本信息")
first_sheet["A2"] = "用户名/id"
first_sheet["A2"].font = bold_font
first_sheet["A3"] = "满分值"
first_sheet["A4"] = "Random/123"

#! 选题情况标题：跨列居中+加粗
start_col_selection = 2
end_col_selection = start_col_selection + num_sheets - 1
first_sheet.merge_cells(
    start_row=1, start_column=start_col_selection,
    end_row=1, end_column=end_col_selection
)
selection_title = first_sheet.cell(row=1, column=start_col_selection)
selection_title.value = "选题情况"
selection_title.font = bold_font
selection_title.alignment = Alignment(horizontal="center", vertical="center")

#! 批改完成情况标题：跨列居中+加粗
NEW_START_COL = end_col_selection + 1
end_col_correction = NEW_START_COL + num_sheets - 1
first_sheet.merge_cells(
    start_row=1, start_column=NEW_START_COL,
    end_row=1, end_column=end_col_correction
)
correction_title = first_sheet.cell(row=1, column=NEW_START_COL)
correction_title.value = "批改完成情况"
correction_title.font = bold_font
correction_title.alignment = Alignment(horizontal="center", vertical="center")

# 填充选题情况和批改完成情况的题名
for i, section in enumerate(["选题情况", "批改完成情况"]):
    start_col = start_col_selection if i == 0 else NEW_START_COL
    for j, name in enumerate(sheet_names):
        cell = first_sheet.cell(row=2, column=start_col + j, value=name)
        cell.alignment = Alignment(horizontal="left")

# 填入1表示选该题，不填任何东西（即空单元格）表示不选题
for j in range(num_sheets):
    first_sheet.cell(row=3, column=start_col_selection + j, value=1)
    first_sheet.cell(row=4, column=start_col_selection + j, value=random.choice([1, ""]))

#! 自动创建各题sheet + 自动填小题
for (raw_title, clean_title), name in zip(sheet_mapping, sheet_names):
    sheet = workbook.create_sheet(clean_title)
    # 基础信息
    sheet["A1"] = "用户名/id"
    sheet["A1"].font = bold_font
    sheet["A2"] = "满分值"
    sheet["A3"] = "=答卷基本信息!A4"
    sheet["A3"].font = green_font
    
    sheet["B1"] = "总分＼小题号"
    sheet["B2"] = "=SUM(C2:IV2)"
    sheet["B2"].font = green_font

    # 从docx读取当前题目的小题列表
    questions = question_data[raw_title]
    # 自动从C列开始填充小题号与满分
    start_col = 3  # C列
    for idx, (q_num, q_score) in enumerate(questions):
        col = start_col + idx
        # 小题号（第1行）
        sheet.cell(row=1, column=col, value=q_num)
        # 小题满分（第2行）
        sheet.cell(row=2, column=col, value=q_score)

# 填充批改情况公式
for j, name in enumerate(sheet_names):
    selection_col = get_column_letter(start_col_selection + j)
    selection_cell = f"{selection_col}3"
    sheet_b2_ref = f"'{name}'!B2"
    target_col = NEW_START_COL + j
    first_sheet.cell(row=3, column=target_col).value = f'=IF(AND({selection_cell}<>"", ISBLANK({sheet_b2_ref})), "✘", "✔")'
    first_sheet.cell(row=3, column=target_col).font = green_font

#! 使用说明
instructions_start_col = end_col_correction + 2
instructions_end_col = instructions_start_col + 12

first_sheet.merge_cells(
    start_row=1, start_column=instructions_start_col,
    end_row=1, end_column=instructions_end_col
)
instructions_title = first_sheet.cell(row=1, column=instructions_start_col)
instructions_title.value = "使用说明"
instructions_title.font = bold_font
instructions_title.alignment = Alignment(horizontal="center", vertical="center")

instructions = [
    "1. 绿色字体单元格为公式单元格，请勿删除！其下方内容需通过下拉填充生成（勿手填）",
    "2. 有新参赛者时，请务必先在第一个工作表填写 '用户名/id数字' 和 选题情况，再到后面的工作表去批改具体题目。选题情况中：'1' 表示选该题，空表示未选",
    "3. 批改完成情况中：'✔' 表示已完成批改，'✘' 表示未完成。自行下拉即可查看批改情况。",
    "4. 批改具体题目时：满分值行用于填写各题满分标准（按自设分值修改）；用户分行对应实际得分。"
]

for row_idx, content in enumerate(instructions, start=2):
    first_sheet.merge_cells(
        start_row=row_idx, start_column=instructions_start_col,
        end_row=row_idx, end_column=instructions_end_col
    )
    cell = first_sheet.cell(row=row_idx, column=instructions_start_col)
    cell.value = content
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

#! 最后一个sheet用于总分统计
last_sheet = workbook.create_sheet("总分统计")
last_sheet["A1"] = "用户名/id"
last_sheet["A1"].font = bold_font
last_sheet["A2"] = "满分值"
last_sheet["A3"] = "=答卷基本信息!A4"
last_sheet["A3"].font = green_font

last_sheet["B1"] = "全卷总分"
last_sheet["B2"] = "=" + "+".join([f"'{sheet}'!B2" for sheet in sheet_names])
last_sheet["B2"].font = green_font
last_sheet["B3"] = "=" + "+".join([f"'{sheet}'!B3" for sheet in sheet_names])
last_sheet["B3"].font = green_font

last_sheet["C1"] = "最高三题总分"

choose_index = "{" + ",".join(str(i) for i in range(1, len(sheet_names)+1)) + "}"

sheet_refs_b2 = ",".join(f"'{sheet}'!B2" for sheet in sheet_names)
sheet_refs_b3 = ",".join(f"'{sheet}'!B3" for sheet in sheet_names)

# 条件片段
cond_b2 = f"IF(CHOOSE({choose_index},{sheet_refs_b2})>0,CHOOSE({choose_index},{sheet_refs_b2}))"
cond_b3 = f"IF(CHOOSE({choose_index},{sheet_refs_b3})>0,CHOOSE({choose_index},{sheet_refs_b3}))"

# 第1大 + 第2大 + 第3大，分别取，取不到就给 0 (因为可能选题 < 3题)
last_sheet["C2"] = "=" + " + ".join([
    f"IFERROR(LARGE({cond_b2},1),0)",
    f"IFERROR(LARGE({cond_b2},2),0)",
    f"IFERROR(LARGE({cond_b2},3),0)"
])

last_sheet["C3"] = "=" + " + ".join([
    f"IFERROR(LARGE({cond_b3},1),0)",
    f"IFERROR(LARGE({cond_b3},2),0)",
    f"IFERROR(LARGE({cond_b3},3),0)"
])

last_sheet["C2"].font = green_font
last_sheet["C3"].font = green_font

last_sheet["D1"] = "排名"

# 排名公式：先比C列(最高三题)，再比B列(全卷总分)，下拉即用
last_sheet["D3"] = "=SUMPRODUCT(($C$3:$C$1000>C3)*1)+SUMPRODUCT(($C$3:$C$1000=C3)*($B$3:$B$1000>B3)*1)+1"
last_sheet["D3"].font = green_font


#! 设置单元格居中
for sheet in workbook.worksheets:
    for row in sheet.iter_rows():
        for cell in row:
            if (sheet.title == "答卷基本信息" and (instructions_start_col <= cell.column <= instructions_end_col)) or \
               (sheet.title == "答卷基本信息" and cell.row == 2 and (start_col_selection <= cell.column <= end_col_selection or NEW_START_COL <= cell.column <= end_col_correction)):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")

import os
from pathlib import Path

def main():
    # 确保 dist 目录存在
    dist_dir = Path(__file__).resolve().parent.parent / "dist"
    dist_dir.mkdir(parents=True, exist_ok=True)
    output_path = dist_dir / "final.xlsx"
    workbook.save(output_path)
    print(f"✓ 文件已保存到 {output_path}")

if __name__ == "__main__":
    main()
    