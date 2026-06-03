"""
ui_app.py
成本报表自动填充界面 - 企业增强版（Logo + 企业主题 + 右上角运行按钮）
✅ 手工调整口：表格版（支持Excel整行/多行粘贴）
✅ 运行按钮固定右上角，不会被挤下去
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

from PIL import Image, ImageTk  # ✅ 用于显示 Logo

from registry import DATA_SOURCES, SHEETS
from state import AppState
from utils_ui import format_path_display
from runner import run_workbook
from profit_sheet6_fill import DEFAULT_EXCLUDE_INV_ITEMS
import sys
import os
from functools import partial


def resource_path(relative_path):
    """获取资源文件的绝对路径，适用于开发环境和PyInstaller打包环境"""
    try:
        # PyInstaller创建临时文件夹存储资源
        base_path = sys._MEIPASS
    except Exception:
        # 开发环境
        base_path = os.path.abspath(".")

    path = os.path.join(base_path, relative_path)
    return path

# ==========================================================
# ✅ 企业主题色
# ==========================================================
THEME_BLUE = "#003366"
THEME_RED = "#B40000"
THEME_GRAY = "#666666"
THEME_LIGHT_GRAY = "#999999"


# ==========================================================
# ✅ 1-1 手工表格：固定锚点列表（与你业务一致）
# ==========================================================
MANUAL_11_ANCHORS = [
    "变动费用-外购动力-新鲜水（吨）",
    "变动费用-外购动力-电（千瓦时、元/千瓦时）",
    "变动费用-外购动力-蒸汽（吨)",
    "变动费用-外购动力-氮气（标立)",
    "变动费用-外购燃料",
    "外供劳务及动力（减项）",
    # "成本费用-加：期初半成品",
    # "成本费用-减：期末半成品",
    "成本费用-其他减少",
    "商品产品总成本-产成品其他减少",
    "商品产品总成本-其他",
]

MANUAL_11_COLS = [
    ("mq", "本月数量(吨)"),
    ("mp", "本月单价(元)"),
    ("ma", "本月金额(万元)"),
    ("yq", "累计数量(吨)"),
    ("yp", "累计单价(元)"),
    ("ya", "累计金额(万元)"),
]


class CostFillerApp:
    def __init__(self):
        self.state = AppState()

        self.root = tk.Tk()
        self.root.title("成本报表自动填充系统（企业增强版）")
        self.root.geometry("1080x720")
        self.root.configure(bg="white")

        self._vars_sheet = {}          # sheet -> BooleanVar
        self._vars_status = {}         # source_key -> StringVar
        self._labels_status = {}       # source_key -> Label
        self._entries_number = {}      # number source_key -> Entry

        # 1-1 手工表格控件
        self._tv_manual_11 = None
        self._manual_11_items = {}     # anchor -> tree item id

        # Logo引用必须持久化，否则会被GC
        self._logo_imgtk = None

        # 进度条/状态
        self._progress_var = tk.DoubleVar(value=0.0)
        self._progress_text = tk.StringVar(value="")
        self._progress_bar = None
        self._progress_label = None
        self._btn_run = None

        self._build_ui()
        self._refresh_required_highlight()

    def run(self):
        self.root.mainloop()

    # ==========================================================
    # ✅ scroll helper（只滚动数据源区域）
    # ==========================================================
    def _make_scrollable_area(self, parent):
        outer = tk.Frame(parent, bg="white")
        outer.pack(fill="both", expand=True, padx=12, pady=8)

        canvas = tk.Canvas(outer, highlightthickness=0, bg="white")
        vbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)

        vbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg="white")
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_config(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_config(event):
            canvas.itemconfigure(win_id, width=event.width)

        inner.bind("<Configure>", _on_inner_config)
        canvas.bind("<Configure>", _on_canvas_config)

        # 鼠标滚轮支持
        def _on_mousewheel(e):
            if getattr(e, "delta", 0):
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        # Linux 滚轮
        def _on_mousewheel_linux(e):
            if e.num == 4:
                canvas.yview_scroll(-3, "units")
            elif e.num == 5:
                canvas.yview_scroll(3, "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", _on_mousewheel_linux)
        canvas.bind_all("<Button-5>", _on_mousewheel_linux)

        return outer, inner

    # ==========================================================
    # ✅ UI 构建
    # ==========================================================
    def _build_ui(self):

        # 若重复构建（某些场景下会触发），先销毁旧 header，避免出现重复 Logo/按钮
        if getattr(self, "_frm_header", None) is not None:
            try:
                self._frm_header.destroy()
            except Exception:
                pass
            self._frm_header = None

        # ==========================================================
        # ✅ 顶部 Logo + 标题（企业风格）
        # ==========================================================
        frm_header = tk.Frame(self.root, bg="white")
        frm_header.pack(fill="x", padx=12, pady=(12, 2))
        self._frm_header = frm_header

        # # Logo
        # logo_path = "logo.png"   # ✅ 你也可以换成 logo.jpeg

        try:
            # img = Image.open(logo_path)
            # img = img.resize((130, 85))
            # self._logo_imgtk = ImageTk.PhotoImage(img)
            # 使用资源路径函数获取logo
            logo_path = resource_path("logo.png")
            print(f"Logo路径: {logo_path}")  # 调试信息，打包后可删除

            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((130, 85))
                self._logo_imgtk = ImageTk.PhotoImage(img)

                lbl_logo = tk.Label(frm_header, image=self._logo_imgtk, bg="white")
                lbl_logo.pack(side="left", padx=(0, 14))
            else:
                # 尝试备用路径
                alt_paths = [
                    os.path.join(os.path.dirname(sys.executable), 'logo.png'),  # exe同目录
                    'logo.png',  # 直接相对路径
                ]

                found = False
                for alt_path in alt_paths:
                    if os.path.exists(alt_path):
                        img = Image.open(alt_path)
                        img = img.resize((130, 85))
                        self._logo_imgtk = ImageTk.PhotoImage(img)

                        lbl_logo = tk.Label(frm_header, image=self._logo_imgtk, bg="white")
                        lbl_logo.pack(side="left", padx=(0, 14))
                        found = True
                        break

                if not found:
                    raise FileNotFoundError("找不到logo文件")

            # 注意：不要在这里再创建一次 lbl_logo（否则会出现两个Logo）

        except Exception:
            lbl_logo = tk.Label(frm_header, text="(Logo未加载)", fg=THEME_LIGHT_GRAY, bg="white")
            lbl_logo.pack(side="left", padx=(0, 14))

        # 标题区
        title_area = tk.Frame(frm_header, bg="white")
        title_area.pack(side="left", fill="x", expand=True)

        tk.Label(
            title_area,
            text="成本报表自动填充系统",
            font=("微软雅黑", 20, "bold"),
            fg=THEME_BLUE,
            bg="white",
        ).pack(anchor="w")

        tk.Label(
            title_area,
            text="数据源面板 + Sheet选择 + 一键运行",
            font=("微软雅黑", 11),
            fg=THEME_GRAY,
            bg="white",
        ).pack(anchor="w", pady=(2, 0))

        tk.Label(
            title_area,
            text="Version: 2025.12 | 内部工具（自动化回填）",
            font=("微软雅黑", 9),
            fg=THEME_LIGHT_GRAY,
            bg="white",
        ).pack(anchor="w", pady=(2, 0))

        # ==========================================================
        # ✅ 右上角按钮区：运行/清空（始终可见）
        # ==========================================================
        btn_area = tk.Frame(frm_header, bg="white")
        btn_area.pack(side="right", padx=8)

        self._btn_run = tk.Button(
            btn_area,
            text="运行",
            width=12,
            height=2,
            bg=THEME_BLUE,
            fg="white",
            font=("微软雅黑", 10, "bold"),
            command=self._run
        )
        self._btn_run.pack(side="top", pady=(6, 4))

        tk.Button(
            btn_area,
            text="清空选择",
            width=12,
            height=1,
            command=self._clear_all
        ).pack(side="top")

        # ==========================================================
        # ✅ 分割线
        # ==========================================================
        sep = tk.Frame(self.root, bg=THEME_BLUE, height=2)
        sep.pack(fill="x", padx=12, pady=(6, 10))

        # ==========================================================
        # ✅ 进度条（UI主线程刷新）
        # ==========================================================
        frm_progress = tk.Frame(self.root, bg="white")
        frm_progress.pack(fill="x", padx=12, pady=(0, 8))

        self._progress_label = tk.Label(
            frm_progress,
            textvariable=self._progress_text,
            bg="white",
            fg=THEME_GRAY,
            anchor="w",
        )
        self._progress_label.pack(fill="x")

        self._progress_bar = ttk.Progressbar(
            frm_progress,
            variable=self._progress_var,
            maximum=1,
            mode="determinate",
        )
        self._progress_bar.pack(fill="x", pady=(4, 0))

        # ==========================================================
        # ① sheet选择
        # ==========================================================
        frm_sheet = tk.LabelFrame(self.root, text="① 选择要运行的Sheet（可多选）", bg="white", fg=THEME_BLUE)
        frm_sheet.pack(fill="x", padx=12, pady=6)

        # ✅ 全选 / 全不选
        frm_sheet_ops = tk.Frame(frm_sheet, bg="white")
        frm_sheet_ops.grid(row=0, column=0, columnspan=4, sticky="e", padx=8, pady=(6, 0))

        tk.Button(
            frm_sheet_ops,
            text="全选",
            width=8,
            command=self._select_all_sheets
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            frm_sheet_ops,
            text="全不选",
            width=8,
            command=self._select_none_sheets
        ).pack(side="left")

        max_cols = 4
        # 从第1行开始放复选框（第0行给操作按钮用了）
        idx = 0
        for sh, meta in SHEETS.items():
            var = tk.BooleanVar(value=(sh == "2-1"))
            self._vars_sheet[sh] = var

            cb = tk.Checkbutton(
                frm_sheet,
                text=f"{sh}  |  {meta.get('label', '')}",
                variable=var,
                command=self._on_sheet_change,
                bg="white",
                fg="black",
                activebackground="white",
            )

            r = (idx // max_cols) + 1
            c = idx % max_cols
            cb.grid(row=r, column=c, padx=10, pady=6, sticky="w")
            idx += 1

        # ==========================================================
        # ② 数据源面板（可滚动）
        # ==========================================================
        _, scroll_inner = self._make_scrollable_area(self.root)

        frm_src = tk.LabelFrame(scroll_inner, text="② 数据源面板（必填项标红；选中变绿；右侧显示路径）", bg="white", fg=THEME_BLUE)
        frm_src.pack(fill="x", padx=0, pady=0)

        r = 0
        for key, meta in DATA_SOURCES.items():
            label = meta["label"]
            kind = meta["kind"]

            tk.Label(frm_src, text=label, bg="white").grid(row=r, column=0, sticky="w", padx=8, pady=6)

            if kind == "file":
                btn = tk.Button(frm_src, text="选择文件", command=lambda k=key: self._choose_file(k))
                btn.grid(row=r, column=1, padx=8, pady=6)

                v = tk.StringVar(value="❌ 未选择")
                self._vars_status[key] = v
                lbl = tk.Label(frm_src, textvariable=v, fg="red", bg="white")
                lbl.grid(row=r, column=2, sticky="w", padx=8)
                self._labels_status[key] = lbl

            elif kind == "number":
                ent = tk.Entry(frm_src, width=16)
                ent.grid(row=r, column=1, sticky="w", padx=8, pady=6)
                ent.insert(0, meta.get("default", ""))

                self._entries_number[key] = ent

                v = tk.StringVar(value="（必填：请输入数值）")
                self._vars_status[key] = v
                lbl = tk.Label(frm_src, textvariable=v, fg="red", bg="white")
                lbl.grid(row=r, column=2, sticky="w", padx=8)
                self._labels_status[key] = lbl

            else:
                raise ValueError(f"未知数据源类型：{kind}")

            r += 1

        # ==========================================================
        # ②-1 扣除投资损益配置
        # ==========================================================
        frm_cfg = tk.LabelFrame(self.root, text="②-1 扣除投资损益配置（可新增/删除）", bg="white", fg=THEME_BLUE)
        frm_cfg.pack(fill="x", padx=12, pady=6)

        self._exclude_items = list(DEFAULT_EXCLUDE_INV_ITEMS)
        self.state.set_source("exclude_inv_items", list(self._exclude_items))

        lb = tk.Listbox(frm_cfg, height=4, width=120)
        lb.grid(row=0, column=0, columnspan=3, padx=8, pady=6, sticky="we")

        for it in self._exclude_items:
            lb.insert(tk.END, it)

        ent_add = tk.Entry(frm_cfg, width=80)
        ent_add.grid(row=1, column=0, padx=8, pady=6, sticky="w")

        def _sync_exclude_to_state():
            items = [lb.get(i) for i in range(lb.size())]
            self.state.set_source("exclude_inv_items", items)

        def _add_item():
            v = ent_add.get().strip()
            if not v:
                return
            lb.insert(tk.END, v)
            ent_add.delete(0, tk.END)
            _sync_exclude_to_state()

        def _del_selected():
            sel = list(lb.curselection())
            sel.reverse()
            for i in sel:
                lb.delete(i)
            _sync_exclude_to_state()

        tk.Button(frm_cfg, text="新增项目", command=_add_item).grid(row=1, column=1, padx=6, pady=6)
        tk.Button(frm_cfg, text="删除选中", command=_del_selected).grid(row=1, column=2, padx=6, pady=6)

        _sync_exclude_to_state()

        # ==========================================================
        # ②-2 ✅ 1-1 手工调整口（表格版）
        # ==========================================================
        frm_manual = tk.LabelFrame(
            self.root,
            text="②-2 1-1 手工调整口（支持从Excel复制整行/多行粘贴；空值表示不覆盖）",
            bg="white",
            fg=THEME_BLUE
        )
        frm_manual.pack(fill="both", expand=False, padx=12, pady=6)

        self._build_manual_table(frm_manual)

        # 底部提示（可选）
        tip = tk.Label(
            self.root,
            text="提示：运行/清空按钮已固定右上角。表格粘贴：从Excel复制 6 列（数量/单价/金额/累计数量/累计单价/累计金额）直接粘贴即可。",
            fg=THEME_GRAY,
            bg="white"
        )
        tip.pack(fill="x", padx=12, pady=6)

    # ==========================================================
    # ✅ 构建手工输入表格 + 粘贴逻辑
    # ==========================================================
    def _build_manual_table(self, parent):
        tv = ttk.Treeview(parent, columns=[c[0] for c in MANUAL_11_COLS], show="headings", height=10)
        self._tv_manual_11 = tv

        tv.pack(side="left", fill="both", expand=True, padx=8, pady=6)

        # 设置列头
        for key, label in MANUAL_11_COLS:
            tv.heading(key, text=label)
            tv.column(key, width=130, anchor="center")

        # 左侧再加一个 anchor 列（使用 #0）
        tv["show"] = "tree headings"
        tv.column("#0", width=280, anchor="w")
        tv.heading("#0", text="锚点")

        # 插入固定锚点行
        for a in MANUAL_11_ANCHORS:
            iid = tv.insert("", "end", text=a, values=[""] * 6)
            self._manual_11_items[a] = iid

        # 右侧按钮
        btns = tk.Frame(parent, bg="white")
        btns.pack(side="right", padx=8, pady=6, fill="y")

        tk.Button(btns, text="清空全部", command=self._manual_clear_all).pack(fill="x", pady=3)
        tk.Button(btns, text="从剪贴板粘贴到选中行", command=self._manual_paste_to_selected).pack(fill="x", pady=3)
        tk.Button(btns, text="复制选中行到剪贴板", command=self._manual_copy_selected).pack(fill="x", pady=3)

        # 双击编辑
        tv.bind("<Double-1>", self._manual_edit_cell)

        # Ctrl+V 粘贴（粘贴到选中行）
        tv.bind_all("<Control-v>", lambda e: self._manual_paste_to_selected())
        tv.bind_all("<Control-V>", lambda e: self._manual_paste_to_selected())

        # 同步一次到 state
        self._sync_manual_table_to_state()

    def _manual_clear_all(self):
        for a, iid in self._manual_11_items.items():
            self._tv_manual_11.item(iid, values=[""] * 6)
        self._sync_manual_table_to_state()

    def _manual_copy_selected(self):
        tv = self._tv_manual_11
        sel = tv.selection()
        if not sel:
            return
        iid = sel[0]
        vals = tv.item(iid, "values")
        txt = "\t".join([str(x) for x in vals])
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)

    def _manual_paste_to_selected(self):
        tv = self._tv_manual_11
        sel = tv.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选中要粘贴的锚点行（左侧锚点）")
            return

        try:
            clip = self.root.clipboard_get()
        except Exception:
            return

        lines = [ln.strip() for ln in clip.splitlines() if ln.strip()]
        if not lines:
            return

        all_iids = tv.get_children("")
        start_idx = all_iids.index(sel[0])

        for i, line in enumerate(lines):
            if start_idx + i >= len(all_iids):
                break
            iid = all_iids[start_idx + i]

            parts = [p.strip() for p in line.split("\t")]
            if len(parts) <= 1:
                parts = [p.strip() for p in line.replace(",", "\t").split("\t")]
            if len(parts) <= 1:
                parts = [p.strip() for p in line.split() if p.strip()]

            while len(parts) < 6:
                parts.append("")
            parts = parts[:6]

            tv.item(iid, values=parts)

        self._sync_manual_table_to_state()

    def _manual_edit_cell(self, event):
        tv = self._tv_manual_11
        region = tv.identify("region", event.x, event.y)
        if region != "cell":
            return

        rowid = tv.identify_row(event.y)
        colid = tv.identify_column(event.x)
        if not rowid or not colid:
            return

        col_index = int(colid.replace("#", "")) - 1
        if col_index < 0:
            return

        x, y, w, h = tv.bbox(rowid, colid)
        old = tv.item(rowid, "values")[col_index]

        ent = tk.Entry(tv)
        ent.place(x=x, y=y, width=w, height=h)
        ent.insert(0, old)
        ent.focus()

        def _save(_e=None):
            new = ent.get().strip()
            vals = list(tv.item(rowid, "values"))
            vals[col_index] = new
            tv.item(rowid, values=vals)
            ent.destroy()
            self._sync_manual_table_to_state()

        ent.bind("<Return>", _save)
        ent.bind("<FocusOut>", _save)

    def _sync_manual_table_to_state(self):
        tv = self._tv_manual_11
        manual = {}
        for a, iid in self._manual_11_items.items():
            vals = tv.item(iid, "values")
            manual[a] = {
                "mq": vals[0],
                "mp": vals[1],
                "ma": vals[2],
                "yq": vals[3],
                "yp": vals[4],
                "ya": vals[5],
            }
        self.state.set_source("manual_1_1_table", manual)

    # ==========================================================
    # required/highlight
    # ==========================================================
    def _expand_deps_ui(self, sh: str):
        deps = SHEETS.get(sh, {}).get("depends_on", []) or []
        for d in deps:
            if d in self._vars_sheet and not self._vars_sheet[d].get():
                self._vars_sheet[d].set(True)
                self._expand_deps_ui(d)

    def _on_sheet_change(self):
        for sh, var in self._vars_sheet.items():
            if var.get():
                self._expand_deps_ui(sh)

        self._sync_selected_sheets()
        self._refresh_required_highlight()

    def _sync_selected_sheets(self):
        self.state.selected_sheets.clear()
        for sh, var in self._vars_sheet.items():
            if var.get():
                self.state.selected_sheets.add(sh)

    def _select_all_sheets(self):
        """一键全选所有Sheet（并自动勾选依赖）。"""
        for sh, var in self._vars_sheet.items():
            var.set(True)
        self._on_sheet_change()

    def _select_none_sheets(self):
        """一键全不选所有Sheet。"""
        for sh, var in self._vars_sheet.items():
            var.set(False)
        self._sync_selected_sheets()
        self._refresh_required_highlight()

    def _required_sources(self):
        req = set()
        for sh in self.state.selected_sheets:
            req |= set(SHEETS[sh]["required_sources"])
        return req

    def _refresh_required_highlight(self):
        self._sync_selected_sheets()
        req = self._required_sources()

        for key, meta in DATA_SOURCES.items():
            lbl = self._labels_status.get(key)
            if not lbl:
                continue

            required = (key in req)
            ok = self._is_source_ready(key)

            if not required:
                lbl.config(fg=THEME_LIGHT_GRAY)
            else:
                lbl.config(fg="green" if ok else "red")

            if meta["kind"] == "number":
                self._vars_status[key].set("✅ 已输入" if ok else "❌ 请输入有效数值")

    def _is_source_ready(self, key: str) -> bool:
        meta = DATA_SOURCES[key]
        if meta["kind"] == "file":
            return bool(self.state.get_source(key))

        if meta["kind"] == "number":
            ent = self._entries_number.get(key)
            if ent is None:
                return False
            try:
                v = float(ent.get().strip())
            except Exception:
                return False

            mn = meta.get("min")
            mx = meta.get("max")
            if mn is not None and v < float(mn):
                return False
            if mx is not None and v > float(mx):
                return False

            return True

        return False

    # ==========================================================
    # actions
    # ==========================================================
    def _choose_file(self, key: str):
        meta = DATA_SOURCES[key]
        filetypes = meta.get("filetypes", [("All Files", "*.*")])
        path = filedialog.askopenfilename(title=f"选择：{meta['label']}", filetypes=filetypes)
        if not path:
            return
        self.state.set_source(key, path)

        self._vars_status[key].set("✅ 已选：" + format_path_display(path))
        self._refresh_required_highlight()

    def _clear_all(self):
        for key, meta in DATA_SOURCES.items():
            if meta["kind"] == "file":
                self.state.clear_source(key)
                self._vars_status[key].set("❌ 未选择")

        for key, ent in self._entries_number.items():
            ent.delete(0, tk.END)
            ent.insert(0, DATA_SOURCES[key].get("default", ""))

        self._manual_clear_all()
        self._refresh_required_highlight()

    def _validate(self):
        if not self.state.selected_sheets:
            return False, ["至少勾选一个要运行的sheet（例如2-1/2-2）"]

        req = self._required_sources()
        missing = []

        for key in req:
            meta = DATA_SOURCES[key]
            if meta["kind"] == "file":
                if not self.state.get_source(key):
                    missing.append(meta["label"])
            elif meta["kind"] == "number":
                if not self._is_source_ready(key):
                    missing.append(meta["label"])

        return (len(missing) == 0), missing

    def _gather_sources(self):
        sources = dict(self.state.sources)
        for key, ent in self._entries_number.items():
            sources[key] = float(ent.get().strip())

        # ✅ 手工表格数据（直接就是 dict）
        sources["manual_1_1_table"] = sources.get("manual_1_1_table") or {}
        return sources

    # ==========================================================
    # ✅ 进度条与运行状态
    # ==========================================================
    def _set_running_state(self, running: bool):
        if self._btn_run is not None:
            self._btn_run.configure(state=("disabled" if running else "normal"))
        if running:
            self._progress_var.set(0.0)
            self._progress_text.set("准备开始...")
            if self._progress_bar is not None:
                self._progress_bar.configure(maximum=1, mode="determinate")
        else:
            # 结束后不强制清空，保留最终状态
            if self._progress_text.get().strip() == "":
                self._progress_text.set("空闲")

    def _on_progress_event(self, ev: dict):
        """主线程中更新进度条。runner progress_cb 会传入 dict。"""
        try:
            total = int(ev.get("total") or 1)
            step = int(ev.get("step") or 0)
        except Exception:
            total, step = 1, 0

        sheet = str(ev.get("sheet") or "")
        phase = str(ev.get("phase") or "")
        detail = str(ev.get("detail") or "")

        if self._progress_bar is not None:
            try:
                self._progress_bar.configure(maximum=max(1, total))
            except Exception:
                pass

        if phase == "start":
            self._progress_var.set(max(0, step - 1))
            msg = f"正在处理：{sheet}（{step}/{total}）"
            if detail:
                msg += f" - {detail}"
            self._progress_text.set(msg)
        elif phase == "done":
            self._progress_var.set(step)
            self._progress_text.set(f"已完成：{sheet}（{step}/{total}）")
        else:
            # 其它自定义阶段
            msg = f"{sheet}（{step}/{total}）"
            if phase:
                msg = f"{phase}: {msg}"
            if detail:
                msg += f" - {detail}"
            self._progress_text.set(msg)

        try:
            self.root.update_idletasks()
        except Exception:
            pass

    def _run(self):
        ok, missing = self._validate()
        if not ok:
            messagebox.showerror("缺少必填数据源", "本次运行缺少以下必填项：\n\n" + "\n".join(missing))
            return

        selected = sorted(list(self.state.selected_sheets))
        sources = self._gather_sources()

        out = filedialog.asksaveasfilename(
            title="保存处理后的成本报表（cost_file 输出）",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not out:
            return

        # 进入运行态
        self._set_running_state(True)

        def _progress_cb(ev: dict):
            try:
                self.root.after(0, partial(self._on_progress_event, ev))
            except Exception:
                pass

        def _worker():
            try:
                wb_cost, audit_path, income_out_path = run_workbook(
                    selected, sources, cost_out_path=out, progress_cb=_progress_cb
                )

                if wb_cost is not None:
                    # 保存成本表
                    wb_cost.save(out)

                msg = f"已完成：{', '.join(selected)}\n\n审查日志：{audit_path}"
                if wb_cost is not None:
                    msg = f"{msg}\n成本表：{out}"
                if income_out_path:
                    msg = f"{msg}\n收入表：{income_out_path}"

                # ✅ 绑定 msg，避免闭包变量问题
                self.root.after(0, lambda msg=msg: messagebox.showinfo("成功", msg))

            except Exception:
                # ✅ 绑定 err，避免 e 的 free variable 问题；同时带堆栈方便排查
                import traceback
                err = traceback.format_exc()
                self.root.after(0, lambda err=err: messagebox.showerror("运行失败", f"发生错误：\n\n{err}"))
            finally:
                self.root.after(0, lambda: self._set_running_state(False))

        t = threading.Thread(target=_worker, daemon=True)
        t.start()


