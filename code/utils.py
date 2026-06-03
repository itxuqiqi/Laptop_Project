# utils.py
from __future__ import annotations

def norm(x):
    """统一字符串：去空格、转小写；None 返回空串"""
    return "" if x is None else str(x).strip().lower()

def to_float(v):
    """把单元格内容转 float；空/“-”/非数字返回 None"""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s == "-":
            return None
        s = s.replace(",", "")
        try:
            return float(s)
        except Exception:
            return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(v)
    except Exception:
        return None

def find_row_by_a(ws, title):
    """在A列精确找到某个标题所在行（去首尾空格），找不到返回None"""
    t = "" if title is None else str(title).strip()
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None:
            continue
        if str(v).strip() == t:
            return r
    return None

def find_row_by_a_contains(ws, text):
    """在A列查找包含text的行（去空格），找不到返回None"""
    t = str(text).strip()
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None:
            continue
        if t in str(v).strip():
            return r
    return None

def build_exist_keys(ws, start_row=7):
    """建立成本表已存在油品键集合：(A分类, D名称)"""
    existed = set()
    for r in range(start_row, ws.max_row + 1):
        a = norm(ws.cell(r, 1).value)
        d = norm(ws.cell(r, 4).value)
        if d:
            existed.add((a, d))
    return existed
