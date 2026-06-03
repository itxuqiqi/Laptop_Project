# intransit_write.py
"""
在途明细写入（2-1 / 2-2 通用）

本次修复点：
1) 解决“insert_rows 导致后续写入错位”的根因：
   - 先写模板已存在的明细行（不插行）
   - 再统一插入新增明细（插行只发生在第二阶段）

2) 新增明细插入位置规则（你当前需求）：
   - 若该类别下已有明细（A列=类别 且 D列有名称），则插到“最后一个明细行”下方
   - 若该类别下没有任何明细，则插到该类别“...-合计”锚点行下方

3) 数据写入口径：
   - Q(17)=数量(万)，S(19)=金额(万)，R(18)=单价=S/Q（Q=0 则 0）
   - 缺失值一律用 0 填充
"""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple
from copy import copy

from utils import norm, find_row_by_a, to_float


Q_COL, R_COL, S_COL = 17, 18, 19


def _safe_num(v) -> float:
    f = to_float(v)
    return 0.0 if f is None else float(f)


def _unit_price(q: float, s: float) -> float:
    return 0.0 if abs(q) < 1e-12 else float(s) / float(q)


def _write_qrs(ws, row: int, qty_wan, amt_wan):
    q = _safe_num(qty_wan)
    s = _safe_num(amt_wan)
    r = _unit_price(q, s)
    ws.cell(row, Q_COL, value=q)
    ws.cell(row, R_COL, value=r)
    ws.cell(row, S_COL, value=s)


def _copy_row_style(ws, src_row: int, dst_row: int, max_col: int):
    """
    复制一行样式（避免 insert_rows 后新行没有格式）。
    只复制 style/number_format/alignment/border/fill/font。
    """
    if src_row < 1 or dst_row < 1:
        return
    for c in range(1, max_col + 1):
        sc = ws.cell(src_row, c)
        dc = ws.cell(dst_row, c)
        if sc.has_style:
            dc._style = copy(sc._style)
            dc.number_format = sc.number_format
            dc.alignment = copy(sc.alignment)
            dc.border = copy(sc.border)
            dc.fill = copy(sc.fill)
            dc.font = copy(sc.font)


def _build_template_maps(ws_cost) -> Tuple[Dict[Tuple[str, str], int], Dict[str, int]]:
    """
    existed_row_map: (cat_norm, name_norm) -> row
    last_detail_row: cat_norm -> max row where (A==cat and D has name)
    """
    existed_row_map: Dict[Tuple[str, str], int] = {}
    last_detail_row: Dict[str, int] = {}

    for r in range(1, ws_cost.max_row + 1):
        cat = ws_cost.cell(r, 1).value
        name = ws_cost.cell(r, 4).value
        if cat is None or name is None:
            continue
        cat_s = str(cat).strip()
        name_s = str(name).strip()
        if cat_s == "" or name_s == "":
            continue

        ck = norm(cat_s)
        nk = norm(name_s)
        existed_row_map[(ck, nk)] = r
        if ck not in last_detail_row or r > last_detail_row[ck]:
            last_detail_row[ck] = r

    return existed_row_map, last_detail_row


def _anchor_title_for_category(cat: str) -> str:
    """
    默认：合计锚点 = “{类别}-合计”
    若类别本身已经以“合计”结尾，则原样返回（兼容某些模板写法）。
    """
    cat_s = "" if cat is None else str(cat).strip()
    if cat_s.endswith("合计"):
        return cat_s
    # 常见模板是 “xxx-合计”
    if cat_s.endswith("-合计"):
        return cat_s
    return f"{cat_s}-合计"


def _group_items(items: List[dict]) -> Dict[str, List[dict]]:
    out: Dict[str, List[dict]] = {}
    for it in items or []:
        cat = "" if it.get("category") is None else str(it.get("category")).strip()
        name = "" if it.get("name") is None else str(it.get("name")).strip()
        if cat == "" or name == "":
            continue
        out.setdefault(cat, []).append(it)
    return out


def write_intransit_to_cost(ws_cost, items: List[dict]):
    """
    items: list[dict] 每条至少包含：
      - category: 分类（写入A列）
      - name: 名称（写入D列）
      - qty_wan: 数量(万)
      - amt_wan: 金额(万)
    """
    if not items:
        return

    # ---------- 1) 扫模板，建立映射 ----------
    existed_row_map, last_detail_row = _build_template_maps(ws_cost)

    # ---------- 2) 分成：模板已有 vs 新增 ----------
    exist_items: List[dict] = []
    new_items_by_cat: Dict[str, List[dict]] = {}

    for it in items:
        cat = "" if it.get("category") is None else str(it.get("category")).strip()
        name = "" if it.get("name") is None else str(it.get("name")).strip()
        if cat == "" or name == "":
            continue
        key = (norm(cat), norm(name))
        if key in existed_row_map:
            exist_items.append(it)
        else:
            new_items_by_cat.setdefault(cat, []).append(it)

    # ---------- 3) 第一遍：先写模板已有（不插行，避免错位） ----------
    for it in exist_items:
        cat = str(it.get("category")).strip()
        name = str(it.get("name")).strip()
        r = existed_row_map.get((norm(cat), norm(name)))
        if r is None:
            continue
        _write_qrs(ws_cost, r, it.get("qty_wan"), it.get("amt_wan"))

    # ---------- 4) 第二遍：统一插入新增（按“最后一个明细下方”，否则“合计下方”） ----------
    if not new_items_by_cat:
        return

    # 4.1 计算每个类别的插入起点 ptr
    insert_jobs: List[Tuple[int, str, List[dict]]] = []
    for cat, lst in new_items_by_cat.items():
        if not lst:
            continue

        ck = norm(cat)

        # 该类别最后一个明细行（若存在）
        last_r = last_detail_row.get(ck)

        # 若无明细，则用合计锚点行
        if last_r is None:
            anchor_title = _anchor_title_for_category(cat)
            anchor_r = find_row_by_a(ws_cost, anchor_title)
            if anchor_r is None:
                # 兜底：找不到合计锚点，则插在表尾
                last_r = ws_cost.max_row
            else:
                last_r = anchor_r

        ptr = int(last_r) + 1

        # 类别内新增明细写入顺序：按名称排序（保持稳定输出）
        lst_sorted = sorted(lst, key=lambda x: (norm(x.get("name")),))
        insert_jobs.append((ptr, cat, lst_sorted))

    # 4.2 为避免插入影响其它类别行号：按 ptr 从大到小执行（从下往上插）
    insert_jobs.sort(key=lambda x: x[0], reverse=True)

    max_col = ws_cost.max_column

    for ptr, cat, lst_sorted in insert_jobs:
        cur = ptr
        for it in lst_sorted:
            name = "" if it.get("name") is None else str(it.get("name")).strip()

            ws_cost.insert_rows(cur)

            # 复制样式：用上一行（cur-1）作为样式来源（若存在）
            if cur - 1 >= 1:
                _copy_row_style(ws_cost, cur - 1, cur, max_col)

            # 写入 A/D
            ws_cost.cell(cur, 1, value=cat)
            ws_cost.cell(cur, 4, value=name)

            # 写入 Q/R/S
            _write_qrs(ws_cost, cur, it.get("qty_wan"), it.get("amt_wan"))

            # 更新 existed_row_map / last_detail_row（用于本轮同类别多条连续插）
            existed_row_map[(norm(cat), norm(name))] = cur
            last_detail_row[norm(cat)] = cur

            cur += 1
