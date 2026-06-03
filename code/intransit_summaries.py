# intransit_summaries.py
from utils import to_float, find_row_by_a

def sum_qrs_for_category_block(ws_cost, detail_category, anchor_title):
    Q_col, R_col, S_col = 17, 18, 19
    anchor_row = find_row_by_a(ws_cost, anchor_title)
    if anchor_row is None:
        print(f"【合计】找不到合计行：{anchor_title}")
        return

    q_sum = 0.0
    s_sum = 0.0

    for r in range(8, ws_cost.max_row + 1):
        a = ws_cost.cell(r, 1).value
        d = ws_cost.cell(r, 4).value
        if d is None or str(d).strip() == "":
            continue
        if str(a).strip() != detail_category:
            continue

        qv = to_float(ws_cost.cell(r, Q_col).value)
        sv = to_float(ws_cost.cell(r, S_col).value)
        if qv is not None:
            q_sum += qv
        if sv is not None:
            s_sum += sv

    ws_cost.cell(anchor_row, Q_col, value=q_sum)
    ws_cost.cell(anchor_row, S_col, value=s_sum)
    ws_cost.cell(anchor_row, R_col, value=0.0 if q_sum == 0 else s_sum / q_sum)

def write_intransit_summaries(ws_cost):
    sum_qrs_for_category_block(ws_cost, "DES/DAT/DAP在途（货权未转移）", "DES/DAT/DAP在途（货权未转移）-合计")
    sum_qrs_for_category_block(ws_cost, "在途原油-海洋原油", "在途原油-海洋原油-合计")
    sum_qrs_for_category_block(ws_cost, "在途原油-进口原油", "在途原油-进口原油-合计")

    Q_col, R_col, S_col = 17, 18, 19
    row_total = find_row_by_a(ws_cost, "在途原油-合计")
    row_ocean = find_row_by_a(ws_cost, "在途原油-海洋原油-合计")
    row_import = find_row_by_a(ws_cost, "在途原油-进口原油-合计")

    if row_total is None or row_ocean is None or row_import is None:
        print("【合计】找不到在途原油-合计或其子合计行，跳过总合计")
        return

    q_total = (to_float(ws_cost.cell(row_ocean, Q_col).value) or 0.0) + (to_float(ws_cost.cell(row_import, Q_col).value) or 0.0)
    s_total = (to_float(ws_cost.cell(row_ocean, S_col).value) or 0.0) + (to_float(ws_cost.cell(row_import, S_col).value) or 0.0)

    ws_cost.cell(row_total, Q_col, value=q_total)
    ws_cost.cell(row_total, S_col, value=s_total)
    ws_cost.cell(row_total, R_col, value=0.0 if q_total == 0 else s_total / q_total)
