from pathlib import Path

YEAR = 2025

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
