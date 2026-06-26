from pathlib import Path
from typing import Literal

# 当前赛季年份
YEAR: int = 2025
# 组别
GROUP: Literal["大学组", "高中组", "初中组"] = "大学组"

TOC_TOP_MARGIN: int = 60  # 目录顶部边距，单位 mm

ROOT = Path(__file__).resolve().parent.parent # 项目根目录

DOC_DIR = ROOT / "doc"
DIST_DIR = ROOT / "dist"
TEMP_DIR = DIST_DIR / "temp_pdf"

# 候选中文字体，供 utils/fonts.py 使用
FONT_CANDIDATES = [
    ("SimSun", "C:/Windows/Fonts/simsun.ttc"),
    ("SimHei", "C:/Windows/Fonts/simhei.ttf"),
    ("Microsoft YaHei", "C:/Windows/Fonts/msyh.ttc"),
    ("Microsoft YaHei Bold", "C:/Windows/Fonts/msyhbd.ttc"),
    ("KaiTi", "C:/Windows/Fonts/simkai.ttf"),
]
