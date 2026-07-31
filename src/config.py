from pathlib import Path
from typing import Literal

# 当前赛季年份
YEAR: int = 2026

# 是否启用文件夹自动识别模式
# True → 扫描 doc 下的子文件夹（如 初中组/、高中组/），按文件夹名生成对应 PDF/XLSX
# False → 直接读取 doc 根目录下的所有 .docx，使用 GROUP 常量
USE_FOLDER: bool = True
GROUP: Literal["大学组", "高中组", "初中组"] = "大学组"

TOC_TOP_MARGIN: int = 30  # 目录顶部边距，单位 mm

ROOT = Path(__file__).resolve().parent.parent # 项目根目录

DOC_DIR = ROOT / "doc"
DIST_DIR = ROOT / "dist"
TEMP_DIR = DIST_DIR / "temp_pdf"
PERIODIC_TABLE_IMG = ROOT / "assets" / "periodic_table.png"

# 候选中文字体
FONT_CANDIDATES = [
    ("SimSun", "C:/Windows/Fonts/simsun.ttc"),
    ("SimHei", "C:/Windows/Fonts/simhei.ttf"),
    ("Microsoft YaHei", "C:/Windows/Fonts/msyh.ttc"),
    ("Microsoft YaHei Bold", "C:/Windows/Fonts/msyhbd.ttc"),
    ("KaiTi", "C:/Windows/Fonts/simkai.ttf"),
]
