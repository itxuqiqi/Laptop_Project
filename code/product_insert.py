"""product_insert.py

新增产品（3-1/3-2）插入逻辑。

用户需求（2026-01）：
1) 3-1/3-2 新增物料不再插到“最后一个物料类别锚点下方”。
2) 改为：插到“当前类别（如：一般贸易）中，最后一个存在物料号(C列)的明细行下方”。
   这样同一类别中带物料号的物料会连续排列，更便于核对。

说明：
- 当前产品底表 pack 不包含“类别”字段，因此仅对模板中 A=一般贸易 的明细块做自动插入。
"""

from utils import find_row_by_a, find_row_by_a_contains

# 兜底：模板结构若出现意外，仍可回退到旧定位逻辑（但尽量不走）
from row_locator import find_insert_row_after_category_anchor


def _norm_a(v) -> str:
    """A列锚点/分类规范化：去空格（含全角空格）。"""
    if v is None:
        return ""
    s = str(v).strip()
    # 去全角空格
    s = s.replace("\u3000", "")
    return s


def _find_anchor_row(ws, anchor_title: str) -> int | None:
    """更鲁棒地找锚点行：先精确匹配，再 contains 兜底。"""
    r = find_row_by_a(ws, anchor_title)
    if r is not None:
        return r
    # 有些模板锚点可能带前后缀/特殊空格
    r = find_row_by_a_contains(ws, anchor_title)
    return r


def _clean_code(v) -> str:
    """统一物料号口径：去空格、去 .0、去前导0。"""
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.lstrip("0").strip()


def _find_insert_row_after_last_code_in_category(
    ws,
    *,
    anchor_title: str,
    detail_category_value: str,
    start_row: int = 6,
    code_col: int = 3,
    name_col: int = 4,
):
    """
    3-1/3-2 新增物料插入位置（按用户口径优化）：

    目标：把新增物料插到“当前类别下最后一个存在物料号的明细行”下方，
    而不是简单插到类别锚点下方。

    规则：
    - 找到 A列锚点行 anchor_title（例如“一般贸易”）
    - 从 anchor_row+1 向下扫描，直到遇到下一个锚点/合计/标题行（经验：A列有值但D列为空）
    - 在扫描区间内，找最后一条满足：A列==detail_category_value 且 C列(物料号)非空 且 D列(名称)非空 的行
    - 若存在，插入位置=该行+1；否则回退到原逻辑（anchor_row+1）
    """

    anchor_row = _find_anchor_row(ws, anchor_title)
    if anchor_row is None:
        return None

    det = _norm_a(detail_category_value)
    r = anchor_row + 1
    last_with_code = None
    while r <= ws.max_row:
        a = ws.cell(r, 1).value
        d = ws.cell(r, name_col).value
        a_s = _norm_a(a)
        d_s = "" if d is None else str(d).strip()

        # 到了下一个锚点/合计/大标题区，停止
        if a_s and not d_s and r != anchor_row:
            break

        if a_s == det:
            code_s = _clean_code(ws.cell(r, code_col).value)
            if code_s and d_s:
                last_with_code = r

        r += 1

    return (last_with_code + 1) if last_with_code else (anchor_row + 1)


def build_exist_product_codes(ws, start_row=6):
    """
    以 C列物料号 去重（物料号唯一）。
    """
    existed = set()
    for r in range(start_row, ws.max_row + 1):
        name = ws.cell(r, 4).value
        if name is None or str(name).strip() == "":
            continue
        code_s = _clean_code(ws.cell(r, 3).value)
        if code_s:
            existed.add(code_s)
    return existed


def insert_new_products_into_general_trade(ws, product_pack: dict, start_row=6):
    """
    仅对“底表有物料号”的新产品自动插入到“一般贸易”锚点下。
    无物料号的底表记录不自动插入（避免重名/误插）。
    """
    by_code = (product_pack or {}).get("by_code", {}) or {}
    existed = build_exist_product_codes(ws, start_row=start_row)

    ptr = _find_insert_row_after_last_code_in_category(
        ws,
        anchor_title="一般贸易",
        detail_category_value="一般贸易",
        start_row=start_row,
        code_col=3,
        name_col=4,
    )

    # 兜底：若模板结构异常（找不到锚点/扫描失败），使用旧逻辑
    if ptr is None:
        ptr = find_insert_row_after_category_anchor(ws, anchor_title="一般贸易", detail_category_value="一般贸易")
    if ptr is None:
        print("【产品插入】找不到锚点：一般贸易，跳过新增产品插入")
        return

    items = []
    for code, rec in by_code.items():
        code_s = _clean_code(code)
        if not code_s:
            continue
        items.append((code_s, rec.get("name_raw", "")))
    items.sort(key=lambda x: x[0])

    inserted = 0
    for code_s, name_raw in items:
        if code_s in existed:
            continue

        ws.insert_rows(ptr)
        ws.cell(ptr, 1, value="一般贸易")   # A 分类
        ws.cell(ptr, 3, value=code_s)       # C 物料号
        ws.cell(ptr, 4, value=name_raw)     # D 产品名（展示）

        existed.add(code_s)
        ptr += 1
        inserted += 1

    if inserted:
        print(f"【产品插入】已插入新增一般贸易产品 {inserted} 条")
