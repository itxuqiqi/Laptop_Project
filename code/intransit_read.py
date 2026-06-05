# intransit_read.py
from __future__ import annotations

import re
from typing import Iterable, Optional

import pandas as pd

from utils import to_float


def _norm_col(s) -> str:
    """规范化列名：去空白/换行、统一括号、转小写，便于按列名匹配。"""
    if s is None:
        return ""
    x = str(s)
    x = re.sub(r"\s+", "", x)
    x = x.replace("（", "(").replace("）", ")")
    return x.lower()


def _find_col(df: pd.DataFrame, specs: Iterable[Iterable[str]]) -> Optional[str]:
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


def _require_col(df: pd.DataFrame, specs: Iterable[Iterable[str]], sheet_name: str, role: str) -> str:
    col = _find_col(df, specs)
    if col is None:
        available = "、".join(str(c) for c in df.columns)
        raise ValueError(f"暂估表 sheet【{sheet_name}】找不到列：{role}。当前列名：{available}")
    return col


def _has_any_value(series: pd.Series) -> bool:
    return not series.dropna().empty


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


# 源表列名配置：只按表头文字取数，不再依赖 Excel 字母列。
NAME_SPECS = [["原油"], ["油种"]]
QTY_TON_SPECS = [["提单净吨"], ["净吨"]]
QTY_BBL_SPECS = [["提单净桶"], ["净桶"]]
AMT_RMB_SPECS = [["应暂估金额", "人民币"], ["暂估金额", "人民币"], ["应暂估金额", "rmb"], ["暂估金额", "rmb"]]
AMT_USD_SPECS = [["应暂估金额", "美元"], ["暂估金额", "美元"], ["应暂估金额", "usd"], ["暂估金额", "usd"]]
FX_SPECS = [["汇率"], ["exchange", "rate"]]
TYPE_SPECS = [["在途类型"], ["在途类别"], ["在途", "类型"], ["类型"]]
UNIT_USD_SPECS = [
    ["应暂估价", "美元/桶"], ["应暂估价", "usd/桶"], ["应暂估价", "usd"],
    ["暂估单价", "美元/桶"], ["暂估单价", "usd/桶"], ["暂估单价", "usd"],
    ["暂估价", "美元/桶"], ["暂估价", "usd/桶"], ["暂估价", "usd"],
]


def read_intransit_estimate(file_path: str, fx_rate: float | None = None):
    """
    读取原油暂估表的三个 sheet：海洋在途、进口在途、其他在途。
    汇总每种原油：提单净吨、暂估金额（人民币）。

    返回：list[dict]
      每条包含：category, name, qty_wan, amt_wan

    取数方式：按列名匹配，不再依赖 B/H/O/J/Q/S/P 等固定列号。
    """
    wb = pd.ExcelFile(file_path)
    results = []

    def _resolve_sheet_name(sheet_name: str) -> str | None:
        if sheet_name in wb.sheet_names:
            return sheet_name
        target = str(sheet_name).strip()
        for actual in wb.sheet_names:
            if str(actual).strip() == target:
                return actual
        return None

    def _read_sheet(sheet_name: str) -> pd.DataFrame:
        actual_sheet_name = _resolve_sheet_name(sheet_name)
        if actual_sheet_name is None:
            print(f"【提示】暂估表缺少sheet：{sheet_name}，跳过")
            return pd.DataFrame()
        # 暂估表表头在第 3 行
        df = pd.read_excel(file_path, sheet_name=actual_sheet_name, header=2)
        if df is None or df.empty:
            return pd.DataFrame()
        df.columns = [str(c).strip() for c in df.columns]
        return df

    def agg_sheet(
        sheet_name: str,
        category: str,
        *,
        amount_mode: str,
        with_type: bool = False,
    ):
        df = _read_sheet(sheet_name)
        if df.empty:
            return

        name_col = _require_col(df, NAME_SPECS, sheet_name, "原油名称/油种")
        qty_col = _require_col(df, QTY_TON_SPECS, sheet_name, "提单净吨")

        df2 = df.copy()
        df2["__name_raw__"] = df2[name_col].apply(lambda x: "" if pd.isna(x) else str(x).strip())
        df2 = df2[df2["__name_raw__"] != ""]
        df2 = df2[~df2["__name_raw__"].apply(_is_summary_name)]
        if df2.empty:
            return

        df2["__qty__"] = df2[qty_col].apply(lambda x: to_float(x) if not pd.isna(x) else None)

        # ---------------- 金额读取逻辑 ----------------
        # amount_mode:
        #   rmb_only          直接读人民币暂估金额
        #   rmb_or_usd_fx     优先人民币；人民币列不存在/为空时，美元暂估金额 * 本行汇率
        #   usd_fx            美元暂估金额 * 本行汇率
        if amount_mode == "rmb_only":
            rmb_col = _require_col(df2, AMT_RMB_SPECS, sheet_name, "应暂估金额（人民币）")
            df2["__amt_rmb__"] = df2[rmb_col].apply(lambda x: to_float(x) if not pd.isna(x) else None)

        elif amount_mode == "rmb_or_usd_fx":
            rmb_col = _find_col(df2, AMT_RMB_SPECS)
            if rmb_col is not None:
                df2["__amt_rmb__"] = df2[rmb_col].apply(lambda x: to_float(x) if not pd.isna(x) else None)
            else:
                df2["__amt_rmb__"] = None

            if not _has_any_value(df2["__amt_rmb__"]):
                usd_col = _require_col(df2, AMT_USD_SPECS, sheet_name, "应暂估金额（美元）")
                fx_col = _require_col(df2, FX_SPECS, sheet_name, "本行汇率")
                df2["__amt__"] = df2[usd_col].apply(lambda x: to_float(x) if not pd.isna(x) else None)
                df2["__fx__"] = df2[fx_col].apply(lambda x: to_float(x) if not pd.isna(x) else None)
                df2["__amt_rmb__"] = df2.apply(
                    lambda r: None if (r["__amt__"] is None or r["__fx__"] is None) else r["__amt__"] * r["__fx__"],
                    axis=1,
                )

        elif amount_mode == "usd_fx":
            usd_col = _require_col(df2, AMT_USD_SPECS, sheet_name, "应暂估金额（美元）")
            fx_col = _require_col(df2, FX_SPECS, sheet_name, "本行汇率")
            df2["__amt__"] = df2[usd_col].apply(lambda x: to_float(x) if not pd.isna(x) else None)
            df2["__fx__"] = df2[fx_col].apply(lambda x: to_float(x) if not pd.isna(x) else None)
            df2["__amt_rmb__"] = df2.apply(
                lambda r: None if (r["__amt__"] is None or r["__fx__"] is None) else r["__amt__"] * r["__fx__"],
                axis=1,
            )
        else:
            raise ValueError(f"未知暂估表金额读取模式：{amount_mode}")

        def _other_type_from_row(row) -> str:
            type_col = _find_col(df2, TYPE_SPECS)
            if type_col is not None:
                v = row.get(type_col)
                if not pd.isna(v) and str(v).strip():
                    return str(v).strip()
            for c in df2.columns:
                if str(c).startswith("Unnamed"):
                    v = row.get(c)
                    if not pd.isna(v):
                        s = str(v).strip()
                        if s.endswith("在途"):
                            return s
            return "其他在途"

        def _display_other_name(row) -> str:
            base = str(row["__name_raw__"]).strip()
            if base and not base.endswith("原油"):
                base = f"{base}原油"
            typ = _other_type_from_row(row)
            return f"{base}（{typ}）" if typ else base

        # 其他在途需要把类型拼到名称后；若源表未给明确类型，默认按“其他在途”展示。
        if with_type:
            df2["__name__"] = df2.apply(_display_other_name, axis=1)
        else:
            df2["__name__"] = df2["__name_raw__"]

        grp = df2.groupby("__name__", as_index=False).agg(
            qty_sum=("__qty__", "sum"),
            amt_sum=("__amt_rmb__", "sum"),
        )

        for _, rr in grp.iterrows():
            name = str(rr["__name__"]).strip()
            if _is_summary_name(name):
                continue
            qty = rr["qty_sum"]
            amt = rr["amt_sum"]
            results.append({
                "category": category,
                "name": name,
                "qty_wan": None if qty is None else qty / 10000.0,
                "amt_wan": None if amt is None else amt / 10000.0,
            })

    # 海洋在途：人民币暂估金额
    agg_sheet("海洋在途", "在途原油-海洋原油", amount_mode="rmb_only")

    # 进口在途：优先人民币；读不到/全空则 美元 * 本行汇率
    agg_sheet("进口在途", "在途原油-进口原油", amount_mode="rmb_or_usd_fx")

    # 其他在途：美元 * 本行汇率，并拼在途类型
    agg_sheet("其他在途", "DES/DAT/DAP在途（货权未转移）", amount_mode="usd_fx", with_type=True)

    return results


def read_intransit_usd_price_pack(file_path: str):
    """
    旧入口保留兼容。
    2-3 的暂估美元/桶 pack 统一委托 intransit_price_pack.build_intransit_price_pack，
    该实现已改为按列名取数。
    """
    from intransit_price_pack import build_intransit_price_pack

    return build_intransit_price_pack(file_path)
