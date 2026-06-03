# row_locator.py
from utils import norm, find_row_by_a

def find_insert_row_after_category_anchor(ws, anchor_title: str, detail_category_value: str):
    """
    在模板中，anchor_title 是A列锚点（如：原油(一般贸易)-进口原油 / 在途原油-进口原油-合计 等）
    detail_category_value 是明细行A列写入的分类值（如：进口原油 / 在途原油-进口原油 等）

    返回：应插入的新行行号（insert_rows用）
    规则：
    - 找到锚点行 anchor_row
    - 从 anchor_row+1 往下走，跳过空行/标题行
    - 只要 A列 == detail_category_value 且 D列非空，认为是该类别明细
    - 插入位置 = 该类别最后一条明细的下一行
    - 若该类别没有任何明细，插入位置 = anchor_row + 1
    """
    anchor_row = find_row_by_a(ws, anchor_title)
    if anchor_row is None:
        return None

    r = anchor_row + 1
    last_detail = None
    while r <= ws.max_row:
        a = ws.cell(r, 1).value
        d = ws.cell(r, 4).value

        a_s = "" if a is None else str(a).strip()
        d_s = "" if d is None else str(d).strip()

        # 到了下一个锚点/合计/大标题区，停止（经验：A列有值但D空，多半是锚点/标题）
        if a_s and not d_s and r != anchor_row:
            break

        if a_s == detail_category_value and d_s:
            last_detail = r

        r += 1

    return (last_detail + 1) if last_detail else (anchor_row + 1)
