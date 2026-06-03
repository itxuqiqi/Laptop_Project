# financial_fill.py
import pandas as pd


def fill_financial_data(df_oil, ws_cost):
    """
    填充 2-1 / 2-2 明细行 E~S：

    需求口径（最终版）：
    - 数量列、金额列：单位换算为“万”（÷10000）
    - 单价列：不从底表取，统一按【金额 ÷ 数量】计算（数量为0则单价=0）
    - 若明细D列名称在df_oil里找不到：E~S 写 0
    """
    df = df_oil.copy()
    df["物料描述"] = df["物料描述"].apply(lambda x: "" if pd.isna(x) else str(x).strip().lower())

    def _clean_num(v):
        if isinstance(v, str):
            v = v.replace(",", "").strip()
        return pd.to_numeric(v, errors="coerce")

    # 仅依赖数量/金额字段（单价字段忽略）
    oil_fields = [
        "月初库存数量", "月初库存总成本",
        "本期进厂数量", "本期进厂总成本",
        "本期加工数量", "本期加工总成本",
        "其他数量", "其他总成本",
        "期末库存数量", "期末库存总成本",
    ]
    for c in oil_fields:
        if c not in df.columns:
            df[c] = pd.NA
        df[c] = df[c].apply(_clean_num)

    oil_map = {}
    for _, r in df.iterrows():
        name = r["物料描述"]
        if name and name not in oil_map:
            oil_map[name] = r

    # 数量/金额列映射：col -> (field, div10000)
    mapping_qty_amt = {
        5:  ("月初库存数量", True),      # E qty
        7:  ("月初库存总成本", True),    # G amt
        8:  ("本期进厂数量", True),      # H qty
        10: ("本期进厂总成本", True),    # J amt
        11: ("本期加工数量", True),      # K qty
        13: ("本期加工总成本", True),    # M amt
        14: ("其他数量", True),          # N qty
        16: ("其他总成本", True),        # P amt
        17: ("期末库存数量", True),      # Q qty
        19: ("期末库存总成本", True),    # S amt
    }

    # 单价列与(金额列, 数量列)对应
    mapping_unit = {
        6:  (7, 5),     # F = G / E
        9:  (10, 8),    # I = J / H
        12: (13, 11),   # L = M / K
        15: (16, 14),   # O = P / N
        18: (19, 17),   # R = S / Q
    }

    start_row = 8
    max_row = ws_cost.max_row

    def _f(v):
        try:
            return float(v)
        except Exception:
            return 0.0

    for row in range(start_row, max_row + 1):
        name_cell = ws_cost.cell(row=row, column=4).value
        name = "" if name_cell is None else str(name_cell).strip().lower()
        if not name:
            continue

        # 先清零 E..S
        for col in range(5, 20):
            ws_cost.cell(row=row, column=col, value=0.0)

        if name not in oil_map:
            continue

        r = oil_map[name]

        # 写数量/金额
        for col, (field, div10000) in mapping_qty_amt.items():
            v = r.get(field, pd.NA)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                ws_cost.cell(row=row, column=col, value=0.0)
            else:
                num = float(v)
                if div10000:
                    num = num / 10000.0
                ws_cost.cell(row=row, column=col, value=num)

        # 计算单价：金额/数量（万/万 => 元/吨），数量为0则0
        for ucol, (acol, qcol) in mapping_unit.items():
            qty = _f(ws_cost.cell(row=row, column=qcol).value)
            amt = _f(ws_cost.cell(row=row, column=acol).value)
            ws_cost.cell(row=row, column=ucol, value=0.0 if qty == 0 else amt / qty)
