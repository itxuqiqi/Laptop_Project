# registry.py

DATA_SOURCES = {
    # 成本表
    "cost_file": {
        "label": "成本报表（含2-1/2-2/2-3/3-1/3-2/4-1）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },

    # 原油数据源
    "oil_month": {
        "label": "原油当月数据（例如11月原油）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },
    "oil_ytd": {
        "label": "原油累计数据（例如1-11月累计原油）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },

    # 产品数据源（3-1/3-2）
    "product_month": {
        "label": "产品当月数据（Sheet1，A=物料号 B=产品名 D-R=期初/生产/销售/其他减少/期末）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },
    "product_ytd": {
        "label": "产品累计数据（Sheet1，同结构，用于3-2）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },

    # 在途暂估
    "intransit": {
        "label": "原油暂估表（海洋在途/进口在途/其他在途）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },

    # 生产经营费用（吨桶比）
    "expense": {
        "label": "生产经营费用表（吨桶比，sheet=原油(万)）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },

    # 平衡表（审查）
    "balance_table": {
        "label": "外购原(燃)料收拨存平衡表（审查用，含原燃料-合并石化 & 产品表-合并石化部）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },

    # 上月成本表
    "prev_cost_file": {
        "label": "上月成本报表（用于2-3期初回填）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls")],
    },

    # 汇率
    "fx_rate_month": {
        "label": "汇率（当月：用于2-1）",
        "kind": "number",
        "default": "7.20",
        "min": 0.000001,
    },
    "fx_rate_ytd": {
        "label": "汇率（累计：用于2-2）",
        "kind": "number",
        "default": "7.20",
        "min": 0.000001,
    },

    # ✅新增：累计销售收入表
    "sales_revenue": {
        "label": "累计销售收入表（用于4-1&收入表，含过账日期/数量/本币金额）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },

    # ✅新增：当前月份（1~12）
    "current_month": {
        "label": "当前月月份（1~12，用于区分当月/累计/上月）",
        "kind": "number",
        "default": "11",
        "min": 1,
        "max": 12,
    },

    # ✅新增：产销存-收入表
    "income_file": {
        "label": "产销存-收入表（sheet=收入，需要填充）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },
# 在 DATA_SOURCES 中添加（大约在第95行后添加）：
    # ✅新增：SAP 利新
    "profit_new": {
        "label": "SAP利润表-利新（12行表头，B=项目名，C/D/E=本月/本年/上年同期）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },
    # ✅新增：SAP 利旧
    "profit_old": {
        "label": "SAP利润表-利旧（12行表头，B=项目名，C/D/E=本月/本年/上年同期）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },
    "semi_cur": {
        "label": "SAP半成品库存-当月底表（A物料号 C描述 D总价值(元) F总数量(吨)，第1行表头）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },
    "semi_prev": {
        "label": "SAP半成品库存-上月底表（A物料号 C描述 D总价值(元) F总数量(吨)，第1行表头）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },

    # ✅新增：1-1 生产经营费用表所需的成本要素表
    "cost_elem_month": {
        "label": "SAP成本要素表-当月（B=成本要素文本，C=实际成本金额元）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },
    "cost_elem_ytd": {
        "label": "SAP成本要素表-累计（B=成本要素文本，C=实际成本金额元）",
        "kind": "file",
        "filetypes": [("Excel Files", "*.xlsx;*.xls;*.xlsm")],
    },
}

SHEETS = {
    "2-1": {
        "label": "当月成本表",
        "required_sources": ["cost_file", "oil_month", "intransit", "expense", "balance_table", "fx_rate_month"],
    },
    "2-2": {
        "label": "累计成本表",
        "required_sources": ["cost_file", "oil_ytd", "intransit", "expense", "balance_table", "fx_rate_ytd"],
    },
    "2-3": {
        "label": "当月+累计汇总展示",
        "required_sources": ["cost_file", "oil_month", "oil_ytd", "intransit", "expense", "fx_rate_month", "fx_rate_ytd", "prev_cost_file"],
    },

    "3-1": {
        "label": "产品当月（期初/生产/销售/其他减少/期末）",
        "required_sources": ["cost_file", "product_month", "balance_table"],
    },
    "3-2": {
        "label": "产品累计（同结构，数据底表为累计）",
        "required_sources": ["cost_file", "product_ytd", "balance_table"],
    },

    # ✅新增：4-1
    "4-1": {
        "label": "4-1 销售收入分析（当月/累计/环比/增利）",
        # ✅新增：product_month 用于“新增物料号”时从产品当月数据补名称
        "required_sources": ["cost_file", "sales_revenue", "current_month", "product_month", "balance_table"],
    },

    # ✅新增：收入表（产销存-收入表的sheet=收入）
    "收入": {
        "label": "产销存-收入表（sheet=收入）填充",
        # ✅新增：product_month 用于“新增物料号”时从产品当月数据补名称
        "required_sources": ["income_file", "sales_revenue", "current_month", "product_month"],
    },
    "4-1(按收入表回填)":{
    "label": "4-1 根据【导入/已调整】的产销存-收入表回填（4-1跟着收入表动）",
    "required_sources": ["cost_file", "income_file", "current_month", "balance_table"],
    },
    "4-2": {
    "label": "4-2 裸价税（税率/消费税来自产销存-裸价税sheet，单价来自收入sheet）",
    # ✅新增：product_ytd 用于“新增物料号”时从产品累计数据补名称
    "required_sources": ["cost_file", "income_file", "current_month", "product_ytd"],
},
# 在 SHEETS 中添加（大约在第170行后添加）：
    "6": {
        "label": "当月利润表（Sheet6）：从SAP利新/利旧回填D/E/F，并支持扣除投资损益项目",
        "required_sources": ["cost_file", "profit_new", "profit_old"],
    },
"7": {
        "label": "7 半成品库存表（当月/上月SAP底表回填，自动追加新增物料）",
        "required_sources": ["cost_file", "semi_cur", "semi_prev"],
    },

    # ✅新增：1-1 生产经营费用表
    "1-1": {
        "label": "1-1 生产经营费用表（原料/产品/成本要素/利新 + 手工口）",
        "required_sources": [
            "cost_file",
            "expense",            # 读取“万元”sheet 的上月累计口径
            "profit_new",         # 管理/财务/销售费用取利新
            "cost_elem_month",    # 成本要素当月
            "cost_elem_ytd",      # 成本要素累计
        ],

        "depends_on": ["2-1", "2-2", "3-1", "3-2", "6"],
    },
# ✅新增：1-1 生产经营费用表
    "1-2": {
        "label": "吨油加工成本表（来自1-1）",
    "required_sources": [],   # ✅不额外要数据源，只依赖成本表内的 1-1
    "depends_on": ["1-1"],     # ✅关键：勾选1-2必须先跑1-1
    },
# ✅新增：1-1 生产经营费用表
    "A-1": {
         "label": "经营指标汇总表（本月）",
    "required_sources": [],
    "depends_on": ["2-1", "3-1", "4-1", "1-2", "6", "7"],
    },
# ✅新增：1-1 生产经营费用表
    "A-2": {
        "label": "经营指标汇总表（累计）",
    "required_sources": [],
    "depends_on": ["2-2", "3-2", "4-1", "1-2", "6", "7"],
    },
}
