# sales_revenue_read.py
import pandas as pd


def _clean_code(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    # 兼容 “123.0”
    if s.endswith(".0"):
        s = s[:-2]
    return s.lstrip("0").strip()


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = {str(c).strip(): c for c in df.columns}
    for k in candidates:
        if k in cols:
            return cols[k]
    # 宽松匹配
    for c in df.columns:
        cc = str(c).strip()
        for k in candidates:
            if k in cc:
                return c
    return None


def build_sales_revenue_pack(file_path: str, current_month: int) -> dict:
    """
    累计销售收入表（第一行标题，数据从第二行）：
    关键列：
      - 物料（物料号，唯一标识）
      - 数量（吨）
      - 本币金额（元）
      - 过账日期（形如 2025/1/7）
      - 文本（可选：物料描述兜底）

    输出 pack：
    {
      "year": 2025,
      "prev_year": 2024,
      "current_month": 11,
      "last_month": (2025, 10) or (2024,12 when current_month=1),
      "by_code": {
         "10001": {
             "desc": "xxx",
             "monthly": {(2025,1): {"qty_ton":..., "amt_yuan":...}, ...}
         },
         ...
      }
    }
    """
    if not (1 <= int(current_month) <= 12):
        raise ValueError(f"当前月月份必须是1~12，当前={current_month}")

    df = pd.read_excel(file_path, header=0, dtype=object)
    if df is None or df.empty:
        return {
            "year": None,
            "prev_year": None,
            "current_month": int(current_month),
            "last_month": None,
            "by_code": {},
        }

    col_code = _pick_col(df, ["物料", "物料号", "物料编码"])
    col_qty = _pick_col(df, ["数量"])
    col_amt = _pick_col(df, ["本币金额", "金额"])
    col_date = _pick_col(df, ["过账日期", "过账日", "日期"])
    col_desc = _pick_col(df, ["文本", "物料描述", "描述"])

    if col_code is None or col_qty is None or col_amt is None or col_date is None:
        raise ValueError(
            "累计销售收入表缺少必要列：物料/数量/本币金额/过账日期（请检查表头是否一致）"
        )

    df2 = df.copy()
    df2["__code__"] = df2[col_code].apply(_clean_code)
    df2 = df2[df2["__code__"] != ""]
    if df2.empty:
        return {
            "year": None,
            "prev_year": None,
            "current_month": int(current_month),
            "last_month": None,
            "by_code": {},
        }

    # 日期解析
    df2["__dt__"] = pd.to_datetime(df2[col_date], errors="coerce")
    df2 = df2[df2["__dt__"].notna()]
    if df2.empty:
        raise ValueError("累计销售收入表的“过账日期”无法解析为日期（例如 2025/1/7）")

    df2["__year__"] = df2["__dt__"].dt.year.astype(int)
    df2["__month__"] = df2["__dt__"].dt.month.astype(int)

    # 数值列
    df2["__qty__"] = pd.to_numeric(df2[col_qty].astype(str).str.replace(",", "").str.strip(), errors="coerce").fillna(0.0)
    df2["__amt__"] = pd.to_numeric(df2[col_amt].astype(str).str.replace(",", "").str.strip(), errors="coerce").fillna(0.0)

    # 目标年份：默认取数据中的最大年份（最常见：当年+可能含上年12月）
    year = int(df2["__year__"].max())
    prev_year = year - 1

    # 聚合：code + (year,month)
    grp = df2.groupby(["__code__", "__year__", "__month__"], as_index=False).agg(
        qty_ton=("__qty__", "sum"),
        amt_yuan=("__amt__", "sum"),
    )

    # 描述：用“文本/描述”列兜底（可空）
    desc_map = {}
    if col_desc is not None:
        tmp = df2[[ "__code__", col_desc ]].copy()
        tmp[col_desc] = tmp[col_desc].apply(lambda x: "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip())
        tmp = tmp[tmp[col_desc] != ""]
        for _, r in tmp.iterrows():
            c = r["__code__"]
            if c not in desc_map:
                desc_map[c] = r[col_desc]

    by_code: dict[str, dict] = {}
    for _, r in grp.iterrows():
        code = str(r["__code__"]).strip()
        y = int(r["__year__"])
        m = int(r["__month__"])
        qty = float(r["qty_ton"]) if r["qty_ton"] is not None else 0.0
        amt = float(r["amt_yuan"]) if r["amt_yuan"] is not None else 0.0

        if code not in by_code:
            by_code[code] = {"desc": desc_map.get(code, ""), "monthly": {}}

        by_code[code]["monthly"][(y, m)] = {"qty_ton": qty, "amt_yuan": amt}

    # last_month
    cm = int(current_month)
    if cm == 1:
        last_month = (prev_year, 12)
    else:
        last_month = (year, cm - 1)

    return {
        "year": year,
        "prev_year": prev_year,
        "current_month": cm,
        "last_month": last_month,
        "by_code": by_code,
    }
