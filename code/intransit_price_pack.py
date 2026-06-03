# intransit_price_pack.py
from __future__ import annotations

import re
import pandas as pd

from utils import norm, to_float


def _read_df(file_path: str, sheet_name: str) -> pd.DataFrame:
    xls = pd.ExcelFile(file_path)
    actual_sheet_name = sheet_name
    if sheet_name not in xls.sheet_names:
        target = str(sheet_name).strip()
        actual_sheet_name = next(
            (actual for actual in xls.sheet_names if str(actual).strip() == target),
            None,
        )
    if actual_sheet_name is None:
        return pd.DataFrame()
    # 暂估表表头在第 3 行
    df = pd.read_excel(file_path, sheet_name=actual_sheet_name, header=2)
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _norm_col(s) -> str:
    """规范化列名：去空白/换行、统一括号、转小写，便于按列名匹配。"""
    if s is None:
        return ""
    x = str(s)
    x = re.sub(r"\s+", "", x)
    x = x.replace("（", "(").replace("）", ")")
    return x.lower()


def _find_col_by_specs(df: pd.DataFrame, specs: list[list[str]]) -> str | None:
    """
    按列名查找列。
    specs 中每个子列表表示一组“必须同时包含”的关键词；多组之间是 OR。
    返回原始列名，找不到返回 None。
    """
    for c in df.columns:
        nc = _norm_col(c)
        for keys in specs:
            if all(_norm_col(k) in nc for k in keys):
                return c
    return None


def _require_col(df: pd.DataFrame, specs: list[list[str]], sheet_name: str, role: str) -> str:
    col = _find_col_by_specs(df, specs)
    if col is None:
        available = "、".join(str(c) for c in df.columns)
        raise ValueError(f"暂估表 sheet【{sheet_name}】找不到列：{role}。当前列名：{available}")
    return col


# 兼容旧函数名：其他代码如仍调用该函数，也改为按规范化列名匹配。
def _find_col_by_keywords(df: pd.DataFrame, keywords: list[str]) -> str | None:
    return _find_col_by_specs(df, [keywords])


NAME_SPECS = [["原油"], ["油种"]]
QTY_BBL_SPECS = [["提单净桶"], ["净桶"]]
AMT_USD_SPECS = [["应暂估金额", "美元"], ["暂估金额", "美元"], ["应暂估金额", "usd"], ["暂估金额", "usd"]]
UNIT_USD_SPECS = [
    ["应暂估价", "美元/桶"], ["应暂估价", "usd/桶"], ["应暂估价", "usd"],
    ["暂估单价", "美元/桶"], ["暂估单价", "usd/桶"], ["暂估单价", "usd"],
    ["暂估价", "美元/桶"], ["暂估价", "usd/桶"], ["暂估价", "usd"],
]
TYPE_SPECS = [["在途类型"], ["在途类别"], ["在途", "类型"], ["类型"]]


def _clean_name(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _clean_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return to_float(v)


def _tax_factor_for_ocean(name_raw: str) -> float:
    """
    海洋在途税换算：
    - 蓬莱（价内税0.95）：* 0.95
    - 其他海洋（1.13）：/ 1.13
    """
    s = "" if name_raw is None else str(name_raw).strip()
    if "蓬莱" in s:
        return 0.95
    return 1.0 / 1.13


def _clean_intransit_type(v) -> str:
    """
    其他在途类型字段清洗：
    - 去掉尾部逗号/顿号/空格
    - 去掉中间多余空格
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    s = re.sub(r"[，,、\s]+$", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def _is_summary_name(name: str) -> bool:
    if name is None:
        return True
    s = str(name).strip().replace(" ", "")
    if s == "":
        return True
    if s in {"合计", "小计", "总计", "汇总"}:
        return True
    if s.endswith("合计") and len(s) <= 6:
        return True
    return False


def build_intransit_price_pack(intransit_file: str) -> dict:
    """
    输出给 fill_sheet_2_3 使用的 pack：

    {
      "detail_price": {(cat_norm, name_norm): usd_per_bbl},
      "cat_price": {cat_norm: usd_per_bbl},
      "intransit_total_price": usd_per_bbl
    }

    取数方式：按列名匹配，不再依赖 B/G/I/J/L/M/N/P 等固定列号。
    """
    detail_price: dict[tuple[str, str], float | None] = {}
    cat_price: dict[str, float | None] = {}

    # =========================================================
    # 1) 海洋在途（按“应暂估价/暂估单价(美元/桶)”做税调整，再加权）
    # =========================================================
    df_o = _read_df(intransit_file, "海洋在途")

    ocean_amt_adj = 0.0  # Σ(桶 * 调整后单价)
    ocean_qty_bbl = 0.0  # Σ桶

    if not df_o.empty:
        sheet = "海洋在途"
        col_name = _require_col(df_o, NAME_SPECS, sheet, "原油名称/油种")
        col_qty = _require_col(df_o, QTY_BBL_SPECS, sheet, "提单净桶")
        col_unit = _find_col_by_specs(df_o, UNIT_USD_SPECS)
        col_amt = _find_col_by_specs(df_o, AMT_USD_SPECS)
        if col_unit is None and col_amt is None:
            raise ValueError(f"暂估表 sheet【{sheet}】找不到列：应暂估价/暂估单价（美元/桶）或 应暂估金额（美元）")

        rows = []
        for _, r in df_o.iterrows():
            name_raw = _clean_name(r.get(col_name))
            if not name_raw or _is_summary_name(name_raw):
                continue

            qty = _clean_num(r.get(col_qty))
            unit = _clean_num(r.get(col_unit)) if col_unit else None
            amt = _clean_num(r.get(col_amt)) if col_amt else None

            if qty is None or qty == 0:
                continue

            base_unit = unit
            if base_unit is None and amt is not None:
                base_unit = float(amt) / float(qty) if float(qty) != 0 else None

            rows.append((name_raw, float(qty), base_unit))

        if rows:
            dff = pd.DataFrame(rows, columns=["name_raw", "qty_bbl", "base_unit"])

            for name_raw, g in dff.groupby("name_raw"):
                qty_sum = float(pd.to_numeric(g["qty_bbl"], errors="coerce").sum(skipna=True))
                if qty_sum == 0:
                    detail_price[(norm("在途原油-海洋原油"), norm(name_raw))] = None
                    continue

                num = 0.0
                has = False
                for _, rr in g.iterrows():
                    q = rr["qty_bbl"]
                    u = rr["base_unit"]
                    if u is None or q is None:
                        continue
                    num += float(q) * float(u)
                    has = True

                base_unit_name = None if not has else (num / qty_sum)

                if base_unit_name is None:
                    adj_unit = None
                else:
                    tax_factor = _tax_factor_for_ocean(name_raw)
                    adj_unit = float(base_unit_name) * float(tax_factor)
                    ocean_amt_adj += qty_sum * adj_unit
                    ocean_qty_bbl += qty_sum

                detail_price[(norm("在途原油-海洋原油"), norm(name_raw))] = adj_unit

    ocean_cat_price = None if ocean_qty_bbl == 0 else (ocean_amt_adj / ocean_qty_bbl)
    cat_price[norm("在途原油-海洋原油")] = ocean_cat_price

    # =========================================================
    # 2) 进口在途（合计=Σ金额/Σ桶；明细按油品加权单价）
    # =========================================================
    df_i = _read_df(intransit_file, "进口在途")
    import_amt_usd = 0.0
    import_qty_bbl = 0.0

    if not df_i.empty:
        sheet = "进口在途"
        col_name = _require_col(df_i, NAME_SPECS, sheet, "原油名称/油种")
        col_qty = _require_col(df_i, QTY_BBL_SPECS, sheet, "提单净桶")
        col_unit = _find_col_by_specs(df_i, UNIT_USD_SPECS)
        col_amt = _find_col_by_specs(df_i, AMT_USD_SPECS)
        if col_unit is None and col_amt is None:
            raise ValueError(f"暂估表 sheet【{sheet}】找不到列：暂估单价（美元/桶）或 应暂估金额（美元）")

        rows = []
        for _, r in df_i.iterrows():
            name_raw = _clean_name(r.get(col_name))
            if not name_raw or _is_summary_name(name_raw):
                continue
            qty = _clean_num(r.get(col_qty))
            unit = _clean_num(r.get(col_unit)) if col_unit else None
            amt = _clean_num(r.get(col_amt)) if col_amt else None
            if qty is None or qty == 0:
                continue

            base_unit = unit
            if base_unit is None and amt is not None:
                base_unit = float(amt) / float(qty) if float(qty) != 0 else None

            rows.append((name_raw, float(qty), base_unit, amt))

        if rows:
            dff = pd.DataFrame(rows, columns=["name_raw", "qty_bbl", "unit_usd", "amt_usd"])

            for name_raw, g in dff.groupby("name_raw"):
                qty_sum = float(pd.to_numeric(g["qty_bbl"], errors="coerce").sum(skipna=True))
                if qty_sum == 0:
                    detail_price[(norm("在途原油-进口原油"), norm(name_raw))] = None
                    continue

                num = 0.0
                has = False
                for _, rr in g.iterrows():
                    q = rr["qty_bbl"]
                    u = rr["unit_usd"]
                    if u is None or q is None:
                        continue
                    num += float(q) * float(u)
                    has = True

                detail_price[(norm("在途原油-进口原油"), norm(name_raw))] = None if not has else (num / qty_sum)

            amt_sum = pd.to_numeric(dff["amt_usd"], errors="coerce").sum(skipna=True)
            qty_sum = pd.to_numeric(dff["qty_bbl"], errors="coerce").sum(skipna=True)

            if qty_sum and float(qty_sum) != 0 and amt_sum is not None and not pd.isna(amt_sum):
                import_amt_usd = float(amt_sum)
                import_qty_bbl = float(qty_sum)
            else:
                num = 0.0
                den = 0.0
                for _, rr in dff.iterrows():
                    q = rr["qty_bbl"]
                    u = rr["unit_usd"]
                    if q is None or u is None:
                        continue
                    num += float(q) * float(u)
                    den += float(q)
                import_amt_usd = num
                import_qty_bbl = den

    import_cat_price = None if import_qty_bbl == 0 else (import_amt_usd / import_qty_bbl)
    cat_price[norm("在途原油-进口原油")] = import_cat_price

    # =========================================================
    # 3) 其他在途（DES/DAT/DAP）
    #    明细单价 key 只按基础原油名；合计优先 Σ金额/Σ桶，否则 Σ(桶*单价)/Σ桶
    # =========================================================
    df_o2 = _read_df(intransit_file, "其他在途")
    other_amt_usd = 0.0
    other_qty_bbl = 0.0

    if not df_o2.empty:
        sheet = "其他在途"
        col_name = _require_col(df_o2, NAME_SPECS, sheet, "原油名称/油种")
        col_qty = _require_col(df_o2, QTY_BBL_SPECS, sheet, "提单净桶")
        col_unit = _find_col_by_specs(df_o2, UNIT_USD_SPECS)
        col_amt = _find_col_by_specs(df_o2, AMT_USD_SPECS)
        # 类型列可选；这里只做兼容，不参与 detail_price key。
        _ = _find_col_by_specs(df_o2, TYPE_SPECS)
        if col_unit is None and col_amt is None:
            raise ValueError(f"暂估表 sheet【{sheet}】找不到列：暂估单价（美元/桶）或 应暂估金额（美元）")

        rows = []
        for _, r in df_o2.iterrows():
            name_raw = _clean_name(r.get(col_name))
            if not name_raw or _is_summary_name(name_raw):
                continue

            qty = _clean_num(r.get(col_qty))
            unit = _clean_num(r.get(col_unit)) if col_unit else None
            amt = _clean_num(r.get(col_amt)) if col_amt else None

            if qty is None or qty == 0:
                continue

            rows.append((name_raw, float(qty), unit, amt))

        if rows:
            dff = pd.DataFrame(rows, columns=["name_raw", "qty_bbl", "unit_usd", "amt_usd"])

            for name_raw, g in dff.groupby("name_raw"):
                qty_sum = float(pd.to_numeric(g["qty_bbl"], errors="coerce").sum(skipna=True))
                if qty_sum == 0:
                    detail_price[(norm("DES/DAT/DAP在途（货权未转移）"), norm(name_raw))] = None
                    continue

                num = 0.0
                has = False
                for _, rr in g.iterrows():
                    q = rr["qty_bbl"]
                    u = rr["unit_usd"]
                    if q is None or u is None:
                        continue
                    num += float(q) * float(u)
                    has = True

                detail_price[(norm("DES/DAT/DAP在途（货权未转移）"), norm(name_raw))] = None if not has else (num / qty_sum)

            amt_sum = pd.to_numeric(dff["amt_usd"], errors="coerce").sum(skipna=True)
            qty_sum = pd.to_numeric(dff["qty_bbl"], errors="coerce").sum(skipna=True)
            if qty_sum and float(qty_sum) != 0 and amt_sum is not None and not pd.isna(amt_sum):
                other_amt_usd = float(amt_sum)
                other_qty_bbl = float(qty_sum)
            else:
                num = 0.0
                den = 0.0
                for _, rr in dff.iterrows():
                    q = rr["qty_bbl"]
                    u = rr["unit_usd"]
                    if q is None or u is None:
                        continue
                    num += float(q) * float(u)
                    den += float(q)
                other_amt_usd = num
                other_qty_bbl = den

    other_cat_price = None if other_qty_bbl == 0 else (other_amt_usd / other_qty_bbl)
    cat_price[norm("DES/DAT/DAP在途（货权未转移）")] = other_cat_price

    # =========================================================
    # 4) 在途原油-合计：只合 进口 + 海洋
    # =========================================================
    total_amt = import_amt_usd + ocean_amt_adj
    total_qty = import_qty_bbl + ocean_qty_bbl
    intransit_total_price = None if total_qty == 0 else (total_amt / total_qty)

    return {
        "detail_price": detail_price,
        "cat_price": cat_price,
        "intransit_total_price": intransit_total_price,
    }
