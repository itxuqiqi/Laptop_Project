# oil_insert.py
import pandas as pd

from utils import norm, build_exist_keys
from row_locator import find_insert_row_after_category_anchor


def check_and_update_oil_data(df_oil, ws_cost, wb_cost):
    """
    ✅2-1/2-2 新增油品插入规则（合计在上、明细在下）：

    你给定的锚点对应关系（合计锚点A列标题 / 明细A列分类）：
      - 原油(一般贸易)-进口原油  / 进口原油
      - 原油(一般贸易)-海洋原油  / 海洋原油
      - 外购原料油-合计          / 外购原料油
      - 来料加工-合计            / 来料加工

    插入位置：
      - 若该类别下没有任何明细物料：插到“合计锚点行”下方（sum_row+1）
      - 若该类别下已有明细物料：插到“最后一个明细物料”下方（追加到末尾）

    判断唯一性：用 (A分类, D名称) 作为唯一键（均走 norm）。
    """

    existed = build_exist_keys(ws_cost, start_row=7)

    # 来源表“物料分类” -> 模板A列分类
    category_alias = {
        "进口原油": "进口原油",
        "海洋原油": "海洋原油",
        "外购原料油": "外购原料油",
    }

    # 合计锚点（A列） + 明细分类值（A列）
    anchor_map = {
        "进口原油": ("原油(一般贸易)-进口原油", "进口原油"),
        "海洋原油": ("原油(一般贸易)-海洋原油", "海洋原油"),
        "外购原料油": ("外购原料油-合计", "外购原料油"),
        "来料加工": ("来料加工-合计", "来料加工"),
    }

    df = df_oil.copy()

    # 兼容列缺失
    if "物料描述" not in df.columns:
        return
    if "物料号" not in df.columns:
        df["物料号"] = None
    if "物料分类" not in df.columns:
        df["物料分类"] = None

    # 清洗：保留原始名称用于写回，但判断/去重使用 norm
    df["__name_raw__"] = df["物料描述"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    df["__code__"] = df["物料号"].apply(lambda x: "" if pd.isna(x) else str(x).strip().lstrip("0"))
    df["__cat_raw__"] = df["物料分类"].apply(lambda x: "" if pd.isna(x) else str(x).strip())

    for _, r in df.iterrows():
        name_raw = r.get("__name_raw__", "")
        if not name_raw:
            continue

        name_key = norm(name_raw)

        # 分类判断：优先名称中包含“来料加工”，否则用“物料分类”映射
        if "来料加工" in name_key:
            mapped_cat = "来料加工"
        else:
            mapped_cat = category_alias.get(r.get("__cat_raw__", ""))

        if not mapped_cat:
            continue

        key = (norm(mapped_cat), name_key)
        if key in existed:
            continue

        anchor_title, detail_cat_value = anchor_map[mapped_cat]
        # ✅合计在上、明细在下：插到该类别明细块的末尾（无明细则在合计下方）
        ptr = find_insert_row_after_category_anchor(ws_cost, anchor_title, detail_cat_value)
        if ptr is None:
            print(f"【警告】找不到锚点：{anchor_title}，跳过插入：{mapped_cat} | {name_raw}")
            continue

        ws_cost.insert_rows(ptr)
        ws_cost.cell(ptr, 1, value=mapped_cat)          # A 分类
        ws_cost.cell(ptr, 3, value=r.get("__code__"))   # C 物料号
        ws_cost.cell(ptr, 4, value=name_raw)            # D 物料描述（保持原始）

        existed.add(key)
