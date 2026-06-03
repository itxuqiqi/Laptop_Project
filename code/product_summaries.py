# product_summaries.py
from utils import find_row_by_a, to_float


def _write_stage(ws, row, col_qty, col_price, col_amt, qty_sum, amt_sum):
    ws.cell(row, col_qty, value=qty_sum)
    ws.cell(row, col_amt, value=amt_sum)
    if qty_sum is None or qty_sum == 0 or amt_sum is None:
        ws.cell(row, col_price, value="0")
    else:
        ws.cell(row, col_price, value=amt_sum / qty_sum)


def write_product_summaries(ws, start_row=6):
    """
    3-1/3-2 合计口径：
    - A=一般贸易 明细求和 -> 写到 A锚点“一般贸易-合计”
    - A=来料加工 明细求和 -> 写到 A锚点“来料加工-合计”
    - A锚点“一般贸易”“来料加工”通常是标题行（无需写数，可保持不动）
    - 同时把 T-V（账面）也写成与 Q-S 一致（合计行同理）
    """

    titles = {
        "general_total": "一般贸易-合计",
        "proc_total": "来料加工-合计",
    }
    r_gen = find_row_by_a(ws, titles["general_total"])
    r_proc = find_row_by_a(ws, titles["proc_total"])

    # 列块：E-G / H-J / K-M / N-P / Q-S / T-V
    blocks = [
        (5, 6, 7),    # 期初
        (8, 9, 10),   # 生产
        (11, 12, 13), # 销售
        (14, 15, 16), # 其他减少
        (17, 18, 19), # 期末
        (20, 21, 22), # 账面（=期末）
    ]

    def sum_cat(cat_value: str):
        sums = []
        for (cq, cp, ca) in blocks:
            sums.append({"qty": 0.0, "amt": 0.0})

        for r in range(start_row, ws.max_row + 1):
            a = ws.cell(r, 1).value
            d = ws.cell(r, 4).value
            if d is None or str(d).strip() == "":
                continue
            if str(a).strip() != cat_value:
                continue

            for i, (cq, _, ca) in enumerate(blocks):
                q = to_float(ws.cell(r, cq).value)
                a0 = to_float(ws.cell(r, ca).value)
                if q is not None:
                    sums[i]["qty"] += q
                if a0 is not None:
                    sums[i]["amt"] += a0

        return sums

    if r_gen is not None:
        sums = sum_cat("一般贸易")

        # 写 E..S（五段），然后 T..V 复制 Q..S
        for i, (cq, cp, ca) in enumerate(blocks[:5]):
            _write_stage(ws, r_gen, cq, cp, ca, sums[i]["qty"], sums[i]["amt"])

        # 账面 = 期末
        ws.cell(r_gen, 20, value=ws.cell(r_gen, 17).value)
        ws.cell(r_gen, 21, value=ws.cell(r_gen, 18).value)
        ws.cell(r_gen, 22, value=ws.cell(r_gen, 19).value)

    if r_proc is not None:
        sums = sum_cat("来料加工")

        for i, (cq, cp, ca) in enumerate(blocks[:5]):
            _write_stage(ws, r_proc, cq, cp, ca, sums[i]["qty"], sums[i]["amt"])

        ws.cell(r_proc, 20, value=ws.cell(r_proc, 17).value)
        ws.cell(r_proc, 21, value=ws.cell(r_proc, 18).value)
        ws.cell(r_proc, 22, value=ws.cell(r_proc, 19).value)
