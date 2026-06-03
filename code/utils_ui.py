# utils_ui.py
import os
from typing import Optional

def basename(p: str) -> str:
    return os.path.basename(p) if p else ""

def format_path_display(path: str, max_len: int = 110) -> str:
    """显示：文件名 + 路径（可控长度）"""
    if not path:
        return ""
    bn = os.path.basename(path)
    if len(path) <= max_len:
        return f"{bn} | {path}"
    # 太长就截断中间
    head = path[: max_len // 2 - 3]
    tail = path[-(max_len // 2):]
    return f"{bn} | {head}...{tail}"
