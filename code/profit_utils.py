# profit_utils.py
import re
from typing import Any, Optional

def normalize_item_name(name: Any) -> str:
    """
    规范化项目名：去掉所有空白字符（含中间空格、全角空格、tab、换行等）。
    不删除星号/三角/※/△ 等符号。
    """
    if name is None:
        return ""
    s = str(name).replace("\u3000", " ")
    s = re.sub(r"\s+", "", s)
    return s

def to_number(v: Any) -> Optional[float]:
    """保留原符号转 float；空/“-”/非数字返回 None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s == "" or s == "-":
        return None
    s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return None

def abs_if_negative(x: Optional[float]) -> Optional[float]:
    """仅当 x<0 时转正；None 不动。"""
    if x is None:
        return None
    return abs(float(x)) if float(x) < 0 else float(x)

def sub(a: Optional[float], b: Optional[float]) -> Optional[float]:
    """a-b；None 视为 0（但若两者都 None 返回 None）。"""
    if a is None and b is None:
        return None
    aa = 0.0 if a is None else float(a)
    bb = 0.0 if b is None else float(b)
    return aa - bb

def to_wanyuan(x: Optional[float]) -> Optional[float]:
    return None if x is None else float(x) / 10000.0

def negate(x: Optional[float]) -> Optional[float]:
    return None if x is None else -float(x)