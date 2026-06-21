import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from src.config import FONT_CANDIDATES

_cjk_font_name = None

def get_cjk_font():
    """返回已注册的中文字体名称，首次调用时自动注册第一个可用字体。"""
    global _cjk_font_name
    if _cjk_font_name is not None:
        return _cjk_font_name

    for family, path in FONT_CANDIDATES:
        if os.path.isfile(path):
            pdfmetrics.registerFont(TTFont("CJKFont", path))
            print(f"Registered CJK font: {family}")
            _cjk_font_name = "CJKFont"
            return _cjk_font_name

    raise FileNotFoundError(
        "No CJK font found. Install SimSun, Microsoft YaHei, SimHei or KaiTi, "
        "or update FONT_CANDIDATES in config.py."
    )
