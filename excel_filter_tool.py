# -*- coding: utf-8 -*-
"""
表格筛选工具（苹果风 · 记忆与预设版）
==================================================
功能：
  1. 打开 Excel（.xlsx / .xlsm），自动识别表头；像 Excel 一样点击每列表头进行筛选。
  2. 自动记忆：上次的表头指纹 + 筛选设置 + 另存目录 + 文件名模板。
     再次打开同一结构的文件时自动套用；另存对话框定位到上次位置。
  3. 筛选预设：可把当前筛选 + 另存位置 + 文件名设置保存为预设，一键应用 / 删除。
  4. 苹果浅灰配色界面（纯 tkinter 绘制，颜色稳定、无系统兼容问题）。
  5. 一键导出，保留原表格式。

依赖：openpyxl（界面使用 Python 自带 tkinter）。
运行：双击本文件，或在命令行执行  python excel_filter_tool.py
"""

import datetime as _dt
import hashlib
import json
import os
import sys
import threading

# ----------------------------------------------------------------------------
# 依赖检查
# ----------------------------------------------------------------------------
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except Exception:
    print("错误：当前 Python 环境缺少 tkinter 图形界面库。")
    print("请到 https://www.python.org 重新安装 Python，并确保勾选 “tcl/tk and IDLE”。")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except Exception:
    try:
        import subprocess
        subprocess.call([sys.executable, "-m", "pip", "install", "openpyxl"])
        import openpyxl
    except Exception:
        try:
            root_tmp = tk.Tk()
            root_tmp.withdraw()
            messagebox.showerror(
                "缺少依赖",
                "缺少 openpyxl 库，且自动安装失败。\n\n请手动在命令行执行：\n"
                "    pip install openpyxl\n\n然后重新运行本程序。",
            )
            root_tmp.destroy()
        except Exception:
            pass
        sys.exit(1)


import tkinter.font as tkfont


PREVIEW_LIMIT = 200        # 预览表最多显示行数（导出为全部）
CHECKBOX_MAX = 100         # 该列不同值超过此数时改用 Listbox 多选

# 主题常量
THEME_BG = "#F5F5F7"
THEME_BG_DARK = "#E8E8ED"
THEME_CARD = "#FAFAFC"
THEME_TEXT = "#1D1D1F"
THEME_SUB = "#86868B"
THEME_ACCENT = "#007AFF"
THEME_ACTIVE_BG = "#C9E4FF"
THEME_BORDER = "#E3E3E8"
THEME_HEADER_BG = "#EFEFF3"
FONT = "Microsoft YaHei UI"


def _round_rect_points(canvas, x1, y1, x2, y2, r):
    """返回圆角矩形轮廓的坐标点列表，用于 Canvas 绘制。"""
    r = max(0, min(r, (x2 - x1) // 2, (y2 - y1) // 2))
    return [
        x1 + r, y1,
        x2 - r, y1,
        x2, y1,
        x2, y1 + r,
        x2, y2 - r,
        x2, y2,
        x2 - r, y2,
        x1 + r, y2,
        x1, y2,
        x1, y2 - r,
        x1, y1 + r,
        x1, y1,
    ]


class RoundedButton(tk.Canvas):
    """
    自绘圆角矩形按钮。
    bg / fg / hover / active 可自定义；点击触发 command。
    """
    def __init__(self, parent, text="", command=None, width=90, height=30, radius=12,
                 bg="#FFFFFF", fg=THEME_TEXT, hover="#F2F2F5", active=THEME_ACCENT,
                 active_fg="#FFFFFF", font=(FONT, 9), padx=0):
        super().__init__(parent, width=width, height=height, highlightthickness=0, bd=0,
                         bg=parent.cget("bg") if parent.cget("bg") else THEME_BG, cursor="hand2")
        self._text = text
        self._command = command
        self._radius = radius
        self._bg = bg
        self._fg = fg
        self._hover = hover
        self._active = active
        self._active_fg = active_fg
        self._font = font
        self._padx = padx
        self._hovered = False
        self._pressed = False

        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    def _on_enter(self, _):
        self._hovered = True
        self._draw()

    def _on_leave(self, _):
        self._hovered = False
        self._pressed = False
        self._draw()

    def _on_press(self, _):
        self._pressed = True
        self._draw()

    def _on_release(self, e):
        was_pressed = self._pressed
        self._pressed = False
        self._draw()
        if was_pressed and self._command and 0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height():
            self._command()

    def _draw(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 2 or h <= 2:
            return
        if self._pressed:
            fill = self._active
            text_color = self._active_fg
        elif self._hovered:
            fill = self._hover
            text_color = self._fg
        else:
            fill = self._bg
            text_color = self._fg

        pts = _round_rect_points(self, 1, 1, w - 1, h - 1, self._radius)
        self.create_polygon(pts, smooth=True, fill=fill, outline=fill)
        self.create_text(w / 2, h / 2, text=self._text, fill=text_color, font=self._font)

    def set_text(self, text):
        self._text = text
        self._draw()

    def set_style(self, bg=None, fg=None):
        """运行时切换背景/前景色（用于表头激活态）。"""
        if bg is not None:
            self._bg = bg
        if fg is not None:
            self._fg = fg
        self._draw()

    def set_enabled(self, enabled):
        """禁用时降低不透明度/变灰。"""
        if not enabled:
            self._fg_saved = self._fg
            self._fg = THEME_SUB
            self._command_saved = self._command
            self._command = None
        else:
            if hasattr(self, "_fg_saved"):
                self._fg = self._fg_saved
            if hasattr(self, "_command_saved"):
                self._command = self._command_saved
        self._draw()


class RoundedCard(tk.Canvas):
    """圆角卡片容器：用于给表格区域加圆角边框。"""
    def __init__(self, parent, radius=16, bg=THEME_CARD, border=THEME_BORDER, pad=2):
        super().__init__(parent, highlightthickness=0, bd=0,
                         bg=parent.cget("bg") if parent.cget("bg") else THEME_BG)
        self._radius = radius
        self._bg = bg
        self._border = border
        self._pad = pad
        self.bind("<Configure>", lambda e: self._draw())

    def _draw(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 4 or h <= 4:
            return
        pts = _round_rect_points(self, 1, 1, w - 1, h - 1, self._radius)
        self.create_polygon(pts, smooth=True, fill=self._bg, outline=self._border, width=1)


def apply_apple_style(root):
    """应用苹果浅灰风格（基础样式）。"""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Treeview", background=THEME_CARD, fieldbackground=THEME_CARD,
                    foreground=THEME_TEXT, borderwidth=0, rowheight=26, font=(FONT, 9))
    style.configure("Treeview.Heading", background=THEME_HEADER_BG, foreground=THEME_TEXT,
                    relief="flat", font=(FONT, 9, "bold"))
    style.map("Treeview", background=[("selected", THEME_ACCENT)],
              foreground=[("selected", "#FFFFFF")])
    style.configure("TScrollbar", background=THEME_BG_DARK, troughcolor=THEME_BG,
                    borderwidth=0, arrowcolor=THEME_SUB)
    style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF",
                    foreground=THEME_TEXT, borderwidth=0, padding=4)


# ----------------------------------------------------------------------------
# 工具函数
# ----------------------------------------------------------------------------
def normalize_value(v):
    if v is None:
        return ""
    if isinstance(v, _dt.datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, _dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, _dt.time):
        return v.strftime("%H:%M")
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return repr(v)
    return str(v)


def blank_label():
    return "（空白）"


def json_value(v):
    """把单元格原始值转成 JSON 可序列化的基本类型。"""
    if v is None:
        return None
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time)):
        return {"__type__": "date", "iso": v.isoformat()}
    if isinstance(v, bool):
        return bool(v)
    if isinstance(v, (int, float)):
        return v
    return str(v)


def from_json_value(j):
    if isinstance(j, dict) and j.get("__type__") == "date":
        s = j.get("iso", "")
        for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
            try:
                return _dt.datetime.strptime(s, fmt)
            except ValueError:
                continue
        return s
    return j


class SourcedCell:
    __slots__ = ("value", "style")

    def __init__(self, value, style):
        self.value = value
        self.style = style


# ----------------------------------------------------------------------------
# 配置与预设持久化
# ----------------------------------------------------------------------------
class ConfigStore:
    """
    配置文件（JSON）结构：
    {
      "last": {
          "fingerprint": "<表头指纹>",
          "selected": {"<列索引>": ["<显示值>", ...]},
          "contains": {"<列索引>": "<文字>"},
          "export_dir": "<路径>",
          "name_template": "<模板>"
      },
      "presets": {
          "<预设名>": { 同 last 结构，可无 fingerprint  }
      }
    }
    注意：selected 以“显示文本”存储，跨文件通用（前提是表头相同）。
    """
    def __init__(self):
        base = os.path.expanduser("~")
        self.dir = os.path.join(base, ".excel_filter_tool")
        self.path = os.path.join(self.dir, "config.json")

    def load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return {"last": {}, "presets": {}}
        if not isinstance(data, dict):
            return {"last": {}, "presets": {}}
        data.setdefault("last", {})
        data.setdefault("presets", {})
        return data

    def save(self, data):
        try:
            os.makedirs(self.dir, exist_ok=True)
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


# ----------------------------------------------------------------------------
# 筛选模型（纯逻辑）
# ----------------------------------------------------------------------------
class FilterModel:
    def __init__(self):
        self.headers = []
        self.rows = []
        self.display_rows = []
        self.value_index = {}
        self.columns = {}
        self.selected = {}
        self.contains = {}

    def load(self, path):
        self.headers = []
        self.rows = []
        self.display_rows = []
        self.value_index = {}
        self.columns = {}
        self.selected = {}
        self.contains = {}

        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active

        header_row_idx = None
        for r in range(1, ws.max_row + 1):
            if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                header_row_idx = r
                break
        if header_row_idx is None:
            raise ValueError("没有找到任何非空内容，文件可能是空的。")

        ncols = ws.max_column
        for c in range(1, ncols + 1):
            v = ws.cell(header_row_idx, c).value
            self.headers.append(str(v) if v is not None else "")

        for r in range(header_row_idx + 1, ws.max_row + 1):
            row = []
            has_content = False
            for c in range(1, ncols + 1):
                cell = ws.cell(r, c)
                if cell.value not in (None, ""):
                    has_content = True
                row.append(SourcedCell(cell.value, cell.style))
            if has_content:
                self.rows.append(row)

        for ci in range(ncols):
            raw = [row[ci].value for row in self.rows]
            seen = set()
            ordered = []
            for v in raw:
                if v not in seen:
                    seen.add(v)
                    ordered.append(v)

            non_blank = [v for v in ordered if v is not None]
            has_blank = len(non_blank) != len(ordered)

            def sort_key(v):
                if isinstance(v, (_dt.datetime, _dt.date)):
                    return (0, v)
                if isinstance(v, (int, float)):
                    return (1, float(v))
                return (2, str(v))

            non_blank.sort(key=sort_key)
            if has_blank:
                non_blank.append(None)

            display = [normalize_value(v) for v in non_blank]
            for i, v in enumerate(non_blank):
                if v is None:
                    display[i] = blank_label()

            self.columns[ci] = {"values": non_blank, "display": display}
            self.selected[ci] = set(range(len(non_blank)))

        self.display_rows = [[normalize_value(cell.value) for cell in row] for row in self.rows]
        self.value_index = {}
        for ci in range(ncols):
            self.value_index[ci] = {v: idx for idx, v in enumerate(self.columns[ci]["values"])}

    # -- 指纹与状态序列化 -----------------------------------------------------
    def fingerprint(self):
        key = "".join(self.headers)
        return hashlib.sha1(key.encode("utf-8")).hexdigest()

    def get_state(self):
        """导出当前筛选状态（以显示文本存储，跨文件通用）。"""
        selected = {}
        for ci in range(len(self.headers)):
            vals = self.columns[ci]["display"]
            if not self.col_is_all_selected(ci):
                selected[str(ci)] = [vals[i] for i in sorted(self.selected[ci])]
        contains = {}
        for ci in range(len(self.headers)):
            c = self.contains.get(ci, "").strip()
            if c:
                contains[str(ci)] = c
        return {"selected": selected, "contains": contains}

    def apply_state(self, state):
        """按显示文本恢复筛选状态。"""
        self.reset_filters()
        if not state:
            return
        sel = state.get("selected") or {}
        for ci in range(len(self.headers)):
            key = str(ci)
            if key in sel:
                wanted = set(sel[key])
                cur = set()
                for i, d in enumerate(self.columns[ci]["display"]):
                    if d in wanted:
                        cur.add(i)
                # 若没有任何匹配（数据变了），保持全选
                self.selected[ci] = cur if cur else set(range(len(self.columns[ci]["values"])))
        con = state.get("contains") or {}
        for ci in range(len(self.headers)):
            key = str(ci)
            if key in con:
                self.contains[ci] = con[key]

    def reset_filters(self):
        for ci in self.columns:
            self.selected[ci] = set(range(len(self.columns[ci]["values"])))
            self.contains[ci] = ""

    def col_is_all_selected(self, ci):
        return len(self.selected[ci]) == len(self.columns[ci]["values"])

    def col_is_active(self, ci):
        return (not self.col_is_all_selected(ci)) or bool(self.contains.get(ci, "").strip())

    def filtered_indices(self):
        ncols = len(self.headers)
        disp = self.display_rows
        validx = self.value_index

        contains = {ci: self.contains.get(ci, "").strip().lower()
                    for ci in range(ncols) if self.contains.get(ci, "").strip()}
        sel_cols = {ci for ci in range(ncols) if not self.col_is_all_selected(ci)}
        active = sorted(sel_cols | set(contains))

        out = []
        for i, row in enumerate(self.rows):
            ok = True
            di = disp[i]
            for ci in active:
                if ci in sel_cols:
                    if validx[ci].get(row[ci].value, -1) not in self.selected[ci]:
                        ok = False
                        break
                if ci in contains:
                    if contains[ci] not in di[ci].lower():
                        ok = False
                        break
            if ok:
                out.append(i)
        return out


# ----------------------------------------------------------------------------
# 列筛选下拉弹窗
# ----------------------------------------------------------------------------
class ColumnFilterPopup:
    def __init__(self, app, ci):
        self.app = app
        self.ci = ci
        self.model = app.model

        self.meta = self.model.columns[ci]
        self.values = self.meta["values"]
        self.displays = self.meta["display"]
        self.n = len(self.values)
        self.mode = "checkbox" if self.n <= CHECKBOX_MAX else "listbox"
        self.visible_indices = list(range(self.n))
        self._search_after = None

        header = self.model.headers[ci] or f"列{ci + 1}"

        self.top = tk.Toplevel(app.root, bg=THEME_BG)
        self.top.title("筛选：" + header)
        self.top.attributes("-topmost", True)
        self.top.resizable(False, True)
        self.top.transient(app.root)

        pad = {"padx": 10, "pady": 4}
        tk.Label(self.top, text=f"{header}", bg=THEME_BG, fg=THEME_TEXT,
                 font=(FONT, 11, "bold"), anchor="w").pack(fill="x", **pad)
        tk.Label(self.top, text=f"{self.n} 个值", bg=THEME_BG, fg=THEME_SUB,
                 font=(FONT, 8), anchor="w").pack(fill="x", padx=10, pady=(0, 4))

        # 搜索值
        srch = tk.Frame(self.top, bg=THEME_BG)
        srch.pack(fill="x", **pad)
        tk.Label(srch, text="搜索值", bg=THEME_BG, fg=THEME_SUB, font=(FONT, 9)).pack(side="left")
        self.search_var = tk.StringVar()
        se = tk.Entry(srch, textvariable=self.search_var, width=16, relief="flat",
                      bg="#FFFFFF", fg=THEME_TEXT, highlightthickness=1,
                      highlightbackground=THEME_BORDER, highlightcolor=THEME_ACCENT)
        se.pack(side="left", fill="x", expand=True, padx=(8, 0), ipady=2)
        se.bind("<KeyRelease>", self._on_search_event)

        # 全选 / 清空
        btns = tk.Frame(self.top, bg=THEME_BG)
        btns.pack(fill="x", **pad)
        self._mk_btn(btns, "全选", self._select_all).pack(side="left")
        self._mk_btn(btns, "清空", self._select_none).pack(side="left", padx=(6, 0))

        # 值区域
        self.value_area = tk.Frame(self.top, bg=THEME_BG)
        self.value_area.pack(fill="both", expand=True, **pad)
        self._build_value_area()

        # 包含文字
        cwrap = tk.Frame(self.top, bg=THEME_BG)
        cwrap.pack(fill="x", **pad)
        tk.Label(cwrap, text="包含文字", bg=THEME_BG, fg=THEME_SUB, font=(FONT, 9)).pack(side="left")
        self.contains_var = tk.StringVar(value=self.model.contains.get(ci, ""))
        ce = tk.Entry(cwrap, textvariable=self.contains_var, width=14, relief="flat",
                      bg="#FFFFFF", fg=THEME_TEXT, highlightthickness=1,
                      highlightbackground=THEME_BORDER, highlightcolor=THEME_ACCENT)
        ce.pack(side="left", fill="x", expand=True, padx=(8, 0), ipady=2)
        ce.bind("<KeyRelease>", self._on_contains)

        # 底部按钮
        bottom = tk.Frame(self.top, bg=THEME_BG)
        bottom.pack(fill="x", **pad)
        self._mk_btn(bottom, "清除本列筛选", self._clear_column).pack(side="left")
        self._mk_btn(bottom, "完成", self._close).pack(side="right")

        self.top.protocol("WM_DELETE_WINDOW", self._close)
        self.top.update_idletasks()
        self._place_near_header()

    def _mk_btn(self, parent, text, cmd):
        return RoundedButton(parent, text=text, command=cmd, width=76, height=28, radius=14,
                             bg=THEME_HEADER_BG, fg=THEME_ACCENT, hover="#E4E4EA",
                             active=THEME_ACCENT, active_fg="#FFFFFF")

    def _build_value_area(self):
        if self.mode == "checkbox":
            canvas = tk.Canvas(self.value_area, width=250, height=300, highlightthickness=0,
                               bg=THEME_BG)
            sb = tk.Scrollbar(self.value_area, orient="vertical", command=canvas.yview)
            self.list_frame = tk.Frame(canvas, bg=THEME_BG)
            self.list_frame.bind("<Configure>",
                                 lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
            canvas.configure(yscrollcommand=sb.set)
            canvas.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            self.vars = []
            self.cbs = []
            cur_sel = self.model.selected[self.ci]
            for vi in range(self.n):
                v = tk.BooleanVar(value=(vi in cur_sel))
                cb = tk.Checkbutton(self.list_frame, text=self.displays[vi], variable=v,
                                    anchor="w", command=self._apply, bg=THEME_BG, fg=THEME_TEXT,
                                    activebackground=THEME_BG, selectcolor="#FFFFFF",
                                    font=(FONT, 9))
                cb.pack(fill="x")
                self.vars.append(v)
                self.cbs.append(cb)
        else:
            self.listbox = tk.Listbox(self.value_area, selectmode="multiple", exportselection=False,
                                      height=18, activestyle="dotbox", relief="flat",
                                      bg="#FFFFFF", fg=THEME_TEXT, selectbackground=THEME_ACCENT,
                                      selectforeground="#FFFFFF", highlightthickness=1,
                                      highlightbackground=THEME_BORDER, font=(FONT, 9))
            sb = tk.Scrollbar(self.value_area, orient="vertical", command=self.listbox.yview)
            self.listbox.configure(yscrollcommand=sb.set)
            self.listbox.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            self._refill_listbox()
            cur_sel = self.model.selected[self.ci]
            for li, vi in enumerate(self.visible_indices):
                if vi in cur_sel:
                    self.listbox.selection_set(li)
            self.listbox.bind("<<ListboxSelect>>", lambda e: self._apply())

    def _refill_listbox(self):
        self.listbox.delete(0, "end")
        for vi in self.visible_indices:
            self.listbox.insert("end", self.displays[vi])

    def _on_search_event(self, e):
        if self._search_after:
            self.top.after_cancel(self._search_after)
        self._search_after = self.top.after(120, self._on_search)

    def _on_search(self):
        kw = self.search_var.get().strip().lower()
        if self.mode == "checkbox":
            for vi, cb in enumerate(self.cbs):
                if kw in self.displays[vi].lower():
                    cb.pack(fill="x")
                else:
                    cb.pack_forget()
        else:
            self.visible_indices = [vi for vi, d in enumerate(self.displays) if kw in d.lower()]
            self._refill_listbox()
            cur_sel = self.model.selected[self.ci]
            for li, vi in enumerate(self.visible_indices):
                if vi in cur_sel:
                    self.listbox.selection_set(li)

    def _on_contains(self, *_):
        self.model.contains[self.ci] = self.contains_var.get()
        self.app.schedule_refresh(150)

    def _sync_from_ui(self):
        if self.mode == "checkbox":
            sel = {vi for vi, v in enumerate(self.vars) if v.get()}
        else:
            sel = {self.visible_indices[li] for li in self.listbox.curselection()}
        self.model.selected[self.ci] = sel
        self.model.contains[self.ci] = self.contains_var.get()

    def _apply(self):
        self._sync_from_ui()
        self.app.schedule_refresh(60)

    def _select_all(self):
        self.search_var.set("")
        if self.mode == "checkbox":
            for v in self.vars:
                v.set(True)
        else:
            self.visible_indices = list(range(self.n))
            self._refill_listbox()
            self.listbox.selection_set(0, "end")
        self._apply()

    def _select_none(self):
        if self.mode == "checkbox":
            for v in self.vars:
                v.set(False)
        else:
            self.listbox.selection_clear(0, "end")
        self._apply()

    def _clear_column(self):
        self.model.selected[self.ci] = set(range(self.n))
        self.model.contains[self.ci] = ""
        self.contains_var.set("")
        self.search_var.set("")
        if self.mode == "checkbox":
            for v in self.vars:
                v.set(True)
        else:
            self.visible_indices = list(range(self.n))
            self._refill_listbox()
            self.listbox.selection_set(0, "end")
        self.app.schedule_refresh(60)
        self._close()

    def _close(self):
        try:
            self.top.destroy()
        except Exception:
            pass

    def _place_near_header(self):
        try:
            btn = self.app.header_buttons.get(self.ci)
            if btn is not None and btn.winfo_viewable():
                x = btn.winfo_rootx()
                y = btn.winfo_rooty() + btn.winfo_height()
                sw = self.top.winfo_screenwidth()
                pw = self.top.winfo_width()
                x = max(0, min(x, sw - pw - 10))
                self.top.geometry(f"+{x}+{y}")
        except Exception:
            pass


# ----------------------------------------------------------------------------
# 主界面
# ----------------------------------------------------------------------------
class App:
    COL_MIN = 90
    COL_MAX = 380

    def __init__(self, root):
        self.root = root
        self.model = FilterModel()
        self.current_path = None
        self.col_widths = []
        self.header_buttons = {}
        self.popup = None
        self._last_idxs = None
        self._refresh_after = None

        self.cfg = ConfigStore()
        self.data = self.cfg.load()
        self.export_dir = os.path.expanduser("~")
        self.name_template = ""

        self._setup_window()

        self._build_toolbar()
        self._build_table_area()
        self._build_statusbar()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_state_no_file()

    def _setup_window(self):
        root = self.root
        root.title("表格筛选工具")
        root.geometry("1180x720")
        root.minsize(900, 560)
        root.configure(bg=THEME_BG)
        apply_apple_style(root)

    # -- 界面骨架 -------------------------------------------------------------
    def _build_toolbar(self):
        bar = tk.Frame(self.root, bg=THEME_BG, padx=14, pady=10)
        bar.pack(side="top", fill="x")

        self.btn_open = self._pill_button(bar, "选择 Excel 文件", self._choose_file, accent=True)
        self.btn_open.pack(side="left")

        self.lbl_path = tk.Label(bar, text="未选择文件", anchor="w", bg=THEME_BG,
                                 fg=THEME_SUB, font=(FONT, 9))
        self.lbl_path.pack(side="left", padx=10, fill="x", expand=True)

        # 预设区
        tk.Label(bar, text="预设", bg=THEME_BG, fg=THEME_SUB, font=(FONT, 9)).pack(side="left")
        self.preset_var = tk.StringVar(value="")
        self.preset_combo = ttk.Combobox(bar, textvariable=self.preset_var, state="readonly",
                                         width=16, font=(FONT, 9))
        self.preset_combo.pack(side="left", padx=(6, 0))
        self.preset_combo.bind("<<ComboboxSelected>>", self._on_preset_selected)
        self._refresh_preset_combo()

        self._pill_button(bar, "保存预设", self._save_preset).pack(side="left", padx=(6, 0))
        self._pill_button(bar, "删除预设", self._delete_preset).pack(side="left", padx=(6, 0))

        self.btn_reset = self._pill_button(bar, "重置筛选", self._reset_filters)
        self.btn_reset.pack(side="right", padx=(0, 6))

        self.btn_export = self._pill_button(bar, "导出 Excel", self._export, accent=True)
        self.btn_export.pack(side="right")

        tk.Label(bar, text="点击列标题筛选", bg=THEME_BG, fg=THEME_SUB,
                 font=(FONT, 8)).pack(side="right", padx=(0, 10))

    def _pill_button(self, parent, text, cmd, accent=False):
        if accent:
            bg, fg, hover, abg = THEME_ACCENT, "#FFFFFF", "#3395FF", "#0066CC"
        else:
            bg, fg, hover, abg = "#FFFFFF", THEME_TEXT, "#F2F2F5", THEME_ACCENT
        return RoundedButton(parent, text=text, command=cmd, width=104, height=32, radius=16,
                             bg=bg, fg=fg, hover=hover, active=abg, active_fg="#FFFFFF")

    def _build_table_area(self):
        outer = tk.Frame(self.root, bg=THEME_BG)
        outer.pack(side="top", fill="both", expand=True)

        self.header_canvas = tk.Canvas(outer, height=36, highlightthickness=0, bd=0, bg=THEME_BG)
        self.header_frame = tk.Frame(self.header_canvas, bg=THEME_BG)
        self.header_window = self.header_canvas.create_window((0, 0), window=self.header_frame, anchor="nw")
        self.header_frame.bind("<Configure>",
                               lambda e: self.header_canvas.configure(scrollregion=self.header_canvas.bbox("all")))

        self.tree = ttk.Treeview(outer, show="", columns=())
        self.vsb = ttk.Scrollbar(outer, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(outer, orient="horizontal", command=self._xscroll_both)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self._tree_xscroll)

        self.header_canvas.grid(row=0, column=0, sticky="ew")
        self.tree.grid(row=1, column=0, sticky="nsew")
        self.vsb.grid(row=1, column=1, sticky="ns")
        self.hsb.grid(row=2, column=0, sticky="ew")
        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=1)

        self.tree.bind("<MouseWheel>", self._on_mousewheel)
        self.tree.bind("<Shift-MouseWheel>", self._on_shiftwheel)
        self.header_canvas.bind("<Shift-MouseWheel>", self._on_shiftwheel)

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="就绪")
        status = tk.Label(self.root, textvariable=self.status_var, anchor="w",
                          bg=THEME_BG, fg=THEME_SUB, font=(FONT, 9), padx=14, pady=6)
        status.pack(side="bottom", fill="x")

    # -- 滚动 ---------------------------------------------------------------
    def _xscroll_both(self, *args):
        self.tree.xview(*args)
        self._sync_header()

    def _tree_xscroll(self, *args):
        self.hsb.set(*args)
        self._sync_header()

    def _sync_header(self):
        self.header_canvas.xview_moveto(self.tree.xview()[0])

    def _on_mousewheel(self, e):
        delta = -1 if e.delta > 0 else 1
        self.tree.yview_scroll(delta, "units")
        return "break"

    def _on_shiftwheel(self, e):
        delta = -1 if e.delta > 0 else 1
        self.tree.xview_scroll(delta, "units")
        self._sync_header()
        return "break"

    # -- 状态管理 -------------------------------------------------------------
    def _set_state_no_file(self):
        self.btn_export.set_enabled(False)
        self.btn_reset.set_enabled(False)
        self.status_var.set("请先选择一个 Excel 文件。")

    def _set_state_loaded(self):
        self.btn_export.set_enabled(True)
        self.btn_reset.set_enabled(True)

    # -- 文件选择 -------------------------------------------------------------
    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self._load_file(path)

    def _load_file(self, path):
        self.status_var.set("正在读取文件…")
        self.root.update_idletasks()
        try:
            self.model.load(path)
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取该文件：\n{e}")
            self.status_var.set("读取失败。")
            return
        self.current_path = path
        self.lbl_path.config(text=os.path.basename(path))
        self._compute_col_widths()
        self._build_headers()

        # 恢复上次设置 / 预设状态
        self._restore_after_load()

        self._set_state_loaded()
        self._last_idxs = None
        self.refresh(force=True)

    def _restore_after_load(self):
        """根据表头指纹恢复上次筛选设置；同时设置另存目录与文件名模板。"""
        last = self.data.get("last") or {}
        self.export_dir = last.get("export_dir") or os.path.dirname(self.current_path)
        self.name_template = last.get("name_template") or ""

        # 默认优先应用当前选中的预设（若下拉里还选中着某个预设）
        applied = False
        presets = self.data.get("presets") or {}
        cur_preset = self.preset_var.get()
        if cur_preset and cur_preset in presets:
            self._apply_preset(presets[cur_preset])
            applied = True

        if not applied:
            fp = self.model.fingerprint()
            if last.get("fingerprint") == fp and (last.get("selected") or last.get("contains")):
                self.model.apply_state({"selected": last.get("selected"),
                                        "contains": last.get("contains")})
                self.status_var.set("已自动套用上次的筛选设置。")
            else:
                self.model.reset_filters()

    # -- 表头 / 列宽 ---------------------------------------------------------
    def _compute_col_widths(self):
        widths = []
        for ci in range(len(self.model.headers)):
            header_len = len(self.model.headers[ci])
            sample_lens = [len(d) for d in self.model.columns[ci]["display"][:60]]
            mx = max([header_len] + sample_lens) if sample_lens else header_len
            w = mx * 13
            widths.append(max(self.COL_MIN, min(self.COL_MAX, w)))
        self.col_widths = widths

    def _build_headers(self):
        for w in self.header_frame.winfo_children():
            w.destroy()
        self.header_buttons = {}

        ncols = len(self.model.headers)
        for ci in range(ncols):
            btn = RoundedButton(self.header_frame, text=self._header_text(ci),
                                command=lambda c=ci: self._open_popup(c),
                                width=self.col_widths[ci], height=30, radius=10,
                                bg=THEME_HEADER_BG, fg=THEME_TEXT,
                                hover="#E4E4EA", active=THEME_ACTIVE_BG, active_fg=THEME_TEXT)
            btn.grid(row=0, column=ci, padx=2, pady=2)
            self.header_buttons[ci] = btn

        self.header_canvas.configure(scrollregion=self.header_canvas.bbox("all"))
        self._update_header_markers()

    def _header_text(self, ci):
        h = self.model.headers[ci] or f"列{ci + 1}"
        return h + (" ●" if self.model.col_is_active(ci) else " ▾")

    def _update_header_markers(self):
        for ci, btn in self.header_buttons.items():
            btn.set_text(self._header_text(ci))
            if self.model.col_is_active(ci):
                btn.set_style(bg=THEME_ACTIVE_BG, fg=THEME_TEXT)
            else:
                btn.set_style(bg=THEME_HEADER_BG, fg=THEME_TEXT)

    # -- 下拉弹窗 -------------------------------------------------------------
    def _open_popup(self, ci):
        if self.popup is not None:
            try:
                self.popup.top.destroy()
            except Exception:
                pass
            self.popup = None
        self.popup = ColumnFilterPopup(self, ci)
        self.popup.top.bind("<Destroy>", self._on_popup_destroy)

    def _on_popup_destroy(self, *_):
        self.popup = None

    # -- 刷新（防抖 + 未变不重建）---------------------------------------------
    def schedule_refresh(self, delay=120):
        if self._refresh_after is not None:
            try:
                self.root.after_cancel(self._refresh_after)
            except Exception:
                pass
        self._refresh_after = self.root.after(delay, self.refresh)

    def refresh(self, force=False):
        if not self.model.rows:
            return
        try:
            idxs = self.model.filtered_indices()
        except Exception as e:
            self.status_var.set(f"筛选出错：{e}")
            return

        self._update_header_markers()

        if force or self._last_idxs is None or idxs != self._last_idxs:
            self._last_idxs = idxs
            self._update_table(idxs)

        total = len(self.model.rows)
        if len(idxs) > PREVIEW_LIMIT:
            self.status_var.set(
                f"匹配 {len(idxs)} 条 / 共 {total} 条（预览仅显示前 {PREVIEW_LIMIT} 条，导出为全部）")
        else:
            self.status_var.set(f"匹配 {len(idxs)} 条 / 共 {total} 条")

    def _update_table(self, idxs):
        tree = self.tree
        tree.delete(*tree.get_children())
        ncols = len(self.model.headers)
        tree["columns"] = list(range(ncols))
        for ci in range(ncols):
            tree.heading(ci, text="")
            tree.column(ci, width=self.col_widths[ci], minwidth=self.col_widths[ci],
                        stretch=False, anchor="w")

        disp = self.model.display_rows
        for i in idxs[:PREVIEW_LIMIT]:
            tree.insert("", "end", values=disp[i])

    # -- 预设 ---------------------------------------------------------------
    def _refresh_preset_combo(self):
        names = sorted((self.data.get("presets") or {}).keys())
        self.preset_combo["values"] = names

    def _current_preset_payload(self):
        return {
            "state": self.model.get_state(),
            "export_dir": self.export_dir,
            "name_template": self.name_template,
        }

    def _save_preset(self):
        if not self.model.rows:
            messagebox.showinfo("提示", "请先打开一个 Excel 文件，再保存预设。")
            return
        from tkinter import simpledialog
        name = simpledialog.askstring("保存预设", "请输入预设名称：", parent=self.root)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        presets = self.data.setdefault("presets", {})
        presets[name] = self._current_preset_payload()
        self.cfg.save(self.data)
        self._refresh_preset_combo()
        self.preset_var.set(name)
        self.status_var.set(f"已保存预设：{name}")

    def _delete_preset(self):
        name = self.preset_var.get()
        presets = self.data.get("presets") or {}
        if not name or name not in presets:
            messagebox.showinfo("提示", "请先在预设下拉里选择一个要删除的预设。")
            return
        if not messagebox.askyesno("删除预设", f"确定删除预设「{name}」吗？"):
            return
        del presets[name]
        self.cfg.save(self.data)
        self._refresh_preset_combo()
        self.preset_var.set("")
        self.status_var.set(f"已删除预设：{name}")

    def _on_preset_selected(self, *_):
        name = self.preset_var.get()
        presets = self.data.get("presets") or {}
        if not name or name not in presets:
            return
        if not self.model.rows:
            messagebox.showinfo("提示", "预设只记录了筛选设置；请先打开 Excel 文件，再套用预设。")
            return
        self._apply_preset(presets[name])
        self.status_var.set(f"已套用预设：{name}")

    def _apply_preset(self, payload):
        self.model.apply_state(payload.get("state") or {})
        d = payload.get("export_dir")
        if d and os.path.isdir(d):
            self.export_dir = d
        self.name_template = payload.get("name_template") or ""
        self.refresh(force=True)

    # -- 重置 / 导出 ----------------------------------------------------------
    def _reset_filters(self):
        self.model.reset_filters()
        if self.popup is not None:
            try:
                self.popup.top.destroy()
            except Exception:
                pass
            self.popup = None
        self.refresh(force=True)

    def _export(self):
        if not self.model.rows:
            return
        idxs = self.model.filtered_indices()
        if not idxs:
            if not messagebox.askyesno("导出为空", "当前筛选结果为空，仍要导出一个只有表头的文件吗？"):
                return

        initial_dir = self.export_dir if os.path.isdir(self.export_dir) else os.path.expanduser("~")
        path = filedialog.asksaveasfilename(
            title="另存为",
            defaultextension=".xlsx",
            initialdir=initial_dir,
            initialfile=self._render_name() or self._default_export_name(),
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return

        self.export_dir = os.path.dirname(path)
        self.status_var.set("正在导出…")
        self.root.update_idletasks()
        self.btn_export.set_enabled(False)

        def work():
            try:
                self._write_export(path, idxs)
            except Exception:
                self.root.after(0, lambda: self._export_failed(path))
            else:
                self.root.after(0, lambda: self._export_done(path))

        threading.Thread(target=work, daemon=True).start()

    def _default_export_name(self):
        base = os.path.splitext(os.path.basename(self.current_path or "筛选结果"))[0]
        return f"{base}-筛选结果.xlsx"

    def _render_name(self):
        """把文件名模板里的 {date} / {match} 替换为实际值。"""
        tpl = (self.name_template or "").strip()
        if not tpl:
            return ""
        d = _dt.datetime.now()
        name = tpl
        name = name.replace("{date}", d.strftime("%Y-%m-%d"))
        name = name.replace("{time}", d.strftime("%H%M"))
        try:
            name = name.replace("{match}", str(len(self.model.filtered_indices())))
        except Exception:
            name = name.replace("{match}", "0")
        if not name.lower().endswith((".xlsx", ".xlsm")):
            name += ".xlsx"
        return name

    def _write_export(self, path, idxs):
        src_path = self.current_path
        if src_path and os.path.exists(src_path):
            wb = openpyxl.load_workbook(src_path)
            ws = wb.active
            header_row_idx = self._find_header_row(ws)
            keep_data_rows = set(header_row_idx + 1 + i for i in idxs)
            for r in range(ws.max_row, header_row_idx, -1):
                if r not in keep_data_rows:
                    ws.delete_rows(r, 1)
            wb.save(path)
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            headers = self.model.headers
            for ci, h in enumerate(headers):
                ws.cell(1, ci + 1, h)
            for ri, i in enumerate(idxs, start=2):
                for ci in range(len(headers)):
                    cell = ws.cell(ri, ci + 1, self.model.rows[i][ci].value)
                    if self.model.rows[i][ci].style:
                        cell.style = self.model.rows[i][ci].style
            wb.save(path)

    def _find_header_row(self, ws):
        for r in range(1, ws.max_row + 1):
            if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                return r
        return 1

    def _export_done(self, path):
        self.btn_export.set_enabled(True)
        self.status_var.set(f"已导出：{os.path.basename(path)}")
        if messagebox.askyesno("导出完成", f"已导出到：\n{path}\n\n是否打开所在文件夹？"):
            self._open_folder(path)

    def _export_failed(self, path):
        self.btn_export.set_enabled(True)
        self.status_var.set("导出失败。")
        messagebox.showerror("导出失败", "导出过程中出现错误。\n\n请确认目标位置没有被其他程序占用。")

    def _open_folder(self, path):
        folder = os.path.dirname(path)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception:
            pass

    # -- 关闭时保存记忆 -------------------------------------------------------
    def _persist_last(self):
        if not self.model.rows:
            return
        last = self.data.setdefault("last", {})
        last["fingerprint"] = self.model.fingerprint()
        last.update(self.model.get_state())
        last["export_dir"] = self.export_dir
        last["name_template"] = self.name_template
        self.cfg.save(self.data)

    def _on_close(self):
        try:
            self._persist_last()
        except Exception:
            pass
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
