# anchor_fill.py
from utils import find_row_by_a, find_row_by_a_contains

def fill_anchor_rows(ws_cost, fx_rate):
    """
    写锚点行（右移到 E-P）：
    - 项目名：合并 E-G / H-J / K-M / N-P 并写标题
    - 元/吨：从一般贸易合计行取单价（期初F、采购I、加工L、期末R）
    - 量：万吨：从一般贸易合计行取数量（期初E、采购H、加工K、期末Q）
    - 汇率（本\\期末）：只写 E 列
    """
    def _merge_keep_value(row, c1, c2, value):
        ws_cost.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
        ws_cost.cell(row=row, column=c1, value=value)

    row_proj = find_row_by_a_contains(ws_cost, "项目名")
    if row_proj is not None:
        _merge_keep_value(row_proj, 5, 7,  "期初库存原油（美元/桶:结算价）")
        _merge_keep_value(row_proj, 8, 10, "采购原油（美元/桶:结算价）")
        _merge_keep_value(row_proj, 11, 13, "加工原油（美元/桶:结算价）")
        _merge_keep_value(row_proj, 14, 16, "库存原油（美元/桶:结算价）")
    else:
        print("【锚点】找不到A列包含“项目名”的行，跳过项目名行写入")

    row_gen = find_row_by_a(ws_cost, "原油(一般贸易)")
    row_imp = find_row_by_a(ws_cost, "原油(一般贸易)-进口原油")
    row_oce = find_row_by_a(ws_cost, "原油(一般贸易)-海洋原油")

    if row_gen is None or row_imp is None or row_oce is None:
        print("【锚点】缺少一般贸易合计行（原油(一般贸易)/进口/海洋），无法填 元/吨 和 量：万吨 行")
    else:
        row_price = find_row_by_a_contains(ws_cost, "元/吨")
        if row_price is not None:
            ws_cost.cell(row_price, 5, value=ws_cost.cell(row_gen, 6).value)
            ws_cost.cell(row_price, 6, value=ws_cost.cell(row_imp, 6).value)
            ws_cost.cell(row_price, 7, value=ws_cost.cell(row_oce, 6).value)

            ws_cost.cell(row_price, 8, value=ws_cost.cell(row_gen, 9).value)
            ws_cost.cell(row_price, 9, value=ws_cost.cell(row_imp, 9).value)
            ws_cost.cell(row_price, 10, value=ws_cost.cell(row_oce, 9).value)

            ws_cost.cell(row_price, 11, value=ws_cost.cell(row_gen, 12).value)
            ws_cost.cell(row_price, 12, value=ws_cost.cell(row_imp, 12).value)
            ws_cost.cell(row_price, 13, value=ws_cost.cell(row_oce, 12).value)

            ws_cost.cell(row_price, 14, value=ws_cost.cell(row_gen, 18).value)
            ws_cost.cell(row_price, 15, value=ws_cost.cell(row_imp, 18).value)
            ws_cost.cell(row_price, 16, value=ws_cost.cell(row_oce, 18).value)
        else:
            print("【锚点】找不到A列包含“元/吨”的行，跳过单价锚点填充")

        row_qty = find_row_by_a_contains(ws_cost, "量：万吨")
        if row_qty is not None:
            ws_cost.cell(row_qty, 5, value=ws_cost.cell(row_gen, 5).value)
            ws_cost.cell(row_qty, 6, value=ws_cost.cell(row_imp, 5).value)
            ws_cost.cell(row_qty, 7, value=ws_cost.cell(row_oce, 5).value)

            ws_cost.cell(row_qty, 8, value=ws_cost.cell(row_gen, 8).value)
            ws_cost.cell(row_qty, 9, value=ws_cost.cell(row_imp, 8).value)
            ws_cost.cell(row_qty, 10, value=ws_cost.cell(row_oce, 8).value)

            ws_cost.cell(row_qty, 11, value=ws_cost.cell(row_gen, 11).value)
            ws_cost.cell(row_qty, 12, value=ws_cost.cell(row_imp, 11).value)
            ws_cost.cell(row_qty, 13, value=ws_cost.cell(row_oce, 11).value)

            ws_cost.cell(row_qty, 14, value=ws_cost.cell(row_gen, 17).value)
            ws_cost.cell(row_qty, 15, value=ws_cost.cell(row_imp, 17).value)
            ws_cost.cell(row_qty, 16, value=ws_cost.cell(row_oce, 17).value)
        else:
            print("【锚点】找不到A列包含“量：万吨”的行，跳过数量锚点填充")

    row_fx = find_row_by_a(ws_cost, "汇率（本\\期末）")
    if row_fx is None:
        row_fx = find_row_by_a_contains(ws_cost, "汇率（本")
    if row_fx is None:
        print("【锚点】找不到“汇率（本\\期末）”行，跳过汇率写入")
    else:
        ws_cost.cell(row_fx, 5, value=fx_rate)
        print(f"【锚点】已写入汇率到行{row_fx} 的 E 列")
