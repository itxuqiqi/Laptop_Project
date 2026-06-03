# general_trade_summaries.py
from utils import to_float, find_row_by_a

def write_general_trade_summaries(ws_cost):
    """
    写回合计行（E~S）：
    - 一般贸易：进口原油 + 海洋原油（不含来料加工）
    - 外购：外购原料油
    - 来料加工：分类为“来料加工”
    - 合计：一般贸易 + 外购 + 来料加工

    数量/金额：求和
    单价：加权平均 = 金额合计 / 数量合计（数量为0则“-”）
    """
    cols = list(range(5, 20))  # E..S
    qty_cols = {5, 8, 11, 14, 17}
    amt_cols = {7, 10, 13, 16, 19}
    unit_cols = {6, 9, 12, 15, 18}

    unit_formula = {
        6: (7, 5),     # F=G/E
        9: (10, 8),    # I=J/H
        12: (13, 11),  # L=M/K
        15: (16, 14),  # O=P/N
        18: (19, 17),  # R=S/Q
    }

    cats = ["进口原油", "海洋原油", "外购原料油", "来料加工"]
    sums = {c: {col: 0.0 for col in cols} for c in cats}

    for r in range(8, ws_cost.max_row + 1):
        a = ws_cost.cell(r, 1).value
        d = ws_cost.cell(r, 4).value
        if d is None or str(d).strip() == "":
            continue

        cat = "" if a is None else str(a).strip()
        if cat not in sums:
            continue

        for col in (qty_cols | amt_cols):
            v = to_float(ws_cost.cell(r, col).value)
            if v is None:
                continue
            sums[cat][col] += v

    def calc_unit_prices(sd):
        for ucol, (acol, qcol) in unit_formula.items():
            q = sd.get(qcol, 0.0) or 0.0
            a = sd.get(acol, 0.0) or 0.0
            sd[ucol] = 0.0 if q == 0 else a / q

    for c in cats:
        calc_unit_prices(sums[c])

    def merge(a, b):
        out = {col: 0.0 for col in cols}
        for col in (qty_cols | amt_cols):
            out[col] = (a.get(col, 0.0) or 0.0) + (b.get(col, 0.0) or 0.0)
        calc_unit_prices(out)
        return out

    sum_import = sums["进口原油"]
    sum_ocean  = sums["海洋原油"]
    sum_buy    = sums["外购原料油"]
    sum_proc   = sums["来料加工"]

    sum_general = merge(sum_import, sum_ocean)
    sum_all = merge(merge(sum_general, sum_buy), sum_proc)

    targets = {
        "原油(一般贸易)-进口原油": sum_import,
        "原油(一般贸易)-海洋原油": sum_ocean,
        "原油(一般贸易)": sum_general,
        "外购原料油-合计": sum_buy,
        "来料加工-合计": sum_proc,
        "合计": sum_all,
    }

    for title, sd in targets.items():
        rr = find_row_by_a(ws_cost, title)
        if rr is None:
            print(f"【合计】找不到合计行：{title}，跳过")
            continue

        for col in cols:
            ws_cost.cell(rr, col, value=sd.get(col, 0.0))

        print(f"【合计】已写入：{title}（行{rr}）")
