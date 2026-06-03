# product_read.py
import pandas as pd


def _norm_name(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip().lower()


def _clean_code(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip().lstrip("0")


def _to_num(v):
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s == "-":
            return None
        s = s.replace(",", "")
        try:
            return float(s)
        except Exception:
            return None
    if isinstance(v, (int, float)):
        try:
            return float(v)
        except Exception:
            return None
    try:
        return float(v)
    except Exception:
        return None


def _nonnull_count(rec: dict) -> int:
    keys = [
        "begin_qty","begin_price","begin_amt",
        "prod_qty","prod_price","prod_amt",
        "sales_qty","sales_price","sales_amt",
        "other_qty","other_price","other_amt",
        "end_qty","end_price","end_amt",
    ]
    return sum(1 for k in keys if rec.get(k) is not None)


def read_product_sheet1_pack(file_path: str, sheet_name: str = "Sheet1"):
    """
    产品底表：Sheet1（第1行标题，数据从第2行开始；pandas header=0即可）
    A=物料号, B=产品名, D-R=期初/生产/销售/其他减少/期末 的 数量/单价/金额

    返回：
    {
      "by_code": {code_norm: rec},
      "by_name": {name_norm: rec}   # 仅用于“底表无物料号”的产品，以及模板行无物料号兜底
    }
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0, dtype=object)
    if df is None or df.empty:
        return {"by_code": {}, "by_name": {}}

    if df.shape[1] < 18:
        raise ValueError(f"产品表列数不足（至少需要到R列），当前列数={df.shape[1]}")

    by_code = {}
    by_name = {}

    for _, r in df.iterrows():
        code = _clean_code(r.iloc[0])          # A
        name_raw = "" if r.iloc[1] is None or (isinstance(r.iloc[1], float) and pd.isna(r.iloc[1])) else str(r.iloc[1]).strip()
        name_norm = _norm_name(name_raw)       # B

        if not code and not name_norm:
            continue

        vals = [r.iloc[i] for i in range(3, 18)]  # D..R
        nums = [_to_num(x) for x in vals]

        def get3(idx0):
            return nums[idx0], nums[idx0 + 1], nums[idx0 + 2]

        bq, bp, ba = get3(0)    # D/E/F
        pq, pp, pa = get3(3)    # G/H/I
        sq, sp, sa = get3(6)    # J/K/L
        oq, op, oa = get3(9)    # M/N/O
        eq, ep, ea = get3(12)   # P/Q/R

        rec = {
            "begin_qty": bq, "begin_price": bp, "begin_amt": ba,
            "prod_qty": pq,  "prod_price": pp, "prod_amt": pa,
            "sales_qty": sq, "sales_price": sp, "sales_amt": sa,
            "other_qty": oq, "other_price": op, "other_amt": oa,
            "end_qty": eq,   "end_price": ep, "end_amt": ea,

            "code_raw": code,
            "name_raw": name_raw,
            "name_norm": name_norm,
        }

        # 1) 优先：按物料号入库（物料号唯一）
        if code:
            if (code not in by_code) or (_nonnull_count(rec) > _nonnull_count(by_code[code])):
                by_code[code] = rec
            continue

        # 2) 无物料号：按名称兜底入库（仅用于“模板行也无物料号”的情况）
        if name_norm:
            if (name_norm not in by_name) or (_nonnull_count(rec) > _nonnull_count(by_name[name_norm])):
                by_name[name_norm] = rec

    return {"by_code": by_code, "by_name": by_name}
