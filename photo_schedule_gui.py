# -*- coding: utf-8 -*-
"""
照片日程录入工具 - 图形界面版

功能：
  选择图片文件夹 -> 扫描并解析文件名 -> 表格预览（可直接编辑）-> 导出 Excel

解析引擎复用 photo_schedule.py（已单独测试通过）。
"""

import os
import re
import sys
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ------------------------------------------------------------
# 运行环境：兼容直接运行脚本 与 PyInstaller 打包后的 exe
# ------------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from photo_schedule import (          # noqa: E402
    parse_filename,
    generate_excel,
    generate_warning_list,
    format_date_short,
    IMAGE_EXTS,
)

APP_TITLE = "照片日程录入工具"
WINDOW_SIZE = "1340x800"

# 表格列定义：(内部列名, 显示标题, 宽度)
COLUMNS = [
    ("filename", "文件名", 220),
    ("date", "时间", 92),
    ("country", "国家", 60),
    ("province", "省份/州", 92),
    ("city", "城市", 72),
    ("venue", "地点", 210),
    ("exhibit", "展览", 190),
    ("activity", "活动", 120),
    ("show", "演出", 160),
    ("sports", "体育", 160),
    ("status", "状态", 78),
]

COL_NAMES = [c[0] for c in COLUMNS]

# 事件四列的内部列名 -> 中文列名
EVENT_COL_MAP = {
    "exhibit": "展览",
    "activity": "活动",
    "show": "演出",
    "sports": "体育",
}

# 不可编辑的列
READONLY_COLS = {"filename", "status"}

FONT_BASE = ("Microsoft YaHei UI", 10)
FONT_SMALL = ("Microsoft YaHei UI", 9)
FONT_BOLD = ("Microsoft YaHei UI", 10, "bold")

COLOR_WARN = "#FFF2CC"      # 需确认
COLOR_MODIFIED = "#E2EFDA"  # 已修改
COLOR_ERROR = "#FCE4E4"     # 严重问题（无日期）


# ============================================================
# 工具函数
# ============================================================

def parse_user_date(text):
    """
    解析用户手动输入的日期，支持多种写法：
      2021-01-11 / 2021/1/11 / 2021年1月11日 / 20210111
    解析失败返回 None。
    """
    if not text:
        return None
    s = text.strip()
    if not s:
        return None

    patterns = [
        r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$",
        r"^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?$",
        r"^(\d{4})(\d{2})(\d{2})$",
    ]
    for p in patterns:
        m = re.match(p, s)
        if m:
            try:
                return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
    return None


def record_to_row(rec):
    """把一条记录转成表格行的字符串元组"""
    events = {k: "" for k in EVENT_COL_MAP}
    if rec.get("event_col") and rec.get("event_text"):
        for internal, cn in EVENT_COL_MAP.items():
            if cn == rec["event_col"]:
                events[internal] = rec["event_text"]
                break

    return (
        rec.get("_filename", ""),
        format_date_short(rec.get("date")),
        rec.get("country") or "",
        rec.get("province") or "",
        rec.get("city") or "",
        rec.get("venue") or "",
        events["exhibit"],
        events["activity"],
        events["show"],
        events["sports"],
        status_of(rec),
    )


def status_of(rec):
    """计算行状态：需确认 / 已修改 / 正常"""
    if rec.get("warnings"):
        return "需确认"
    if rec.get("modified"):
        return "已修改"
    return "正常"


def tag_of(rec):
    """根据状态返回 Treeview 的着色 tag"""
    if not rec.get("date"):
        return "error"
    if rec.get("warnings"):
        return "warn"
    if rec.get("modified"):
        return "modified"
    return ""


def apply_edit(rec, col_name, value):
    """把用户在表格里的编辑写回记录"""
    value = (value or "").strip()

    if col_name == "date":
        rec["date"] = parse_user_date(value)
        if value and rec["date"] is None:
            rec.setdefault("_edit_warnings", []).append(f"日期格式无法识别: {value}")
    elif col_name == "country":
        rec["country"] = value or None
    elif col_name == "province":
        rec["province"] = value or None
    elif col_name == "city":
        rec["city"] = value or None
    elif col_name == "venue":
        rec["venue"] = value or None
    elif col_name in EVENT_COL_MAP:
        cn = EVENT_COL_MAP[col_name]
        if value:
            rec["event_col"] = cn
            rec["event_text"] = value
            # 清掉"事件无法自动归类"的告警
            rec["warnings"] = [
                w for w in rec.get("warnings", [])
                if not w.startswith("事件无法自动归类")
            ]
        else:
            if rec.get("event_col") == cn:
                rec["event_col"] = None
                rec["event_text"] = None

    rec["modified"] = True


def collect_image_files(folder, recursive):
    """收集文件夹下的图片文件，返回 [(目录, 文件名), ...]"""
    found = []
    if recursive:
        for root, _dirs, names in os.walk(folder):
            for n in names:
                if os.path.splitext(n)[1].lower() in IMAGE_EXTS:
                    found.append((root, n))
    else:
        try:
            entries = os.listdir(folder)
        except OSError:
            return found
        for n in entries:
            p = os.path.join(folder, n)
            if os.path.isfile(p) and os.path.splitext(n)[1].lower() in IMAGE_EXTS:
                found.append((folder, n))
    found.sort(key=lambda t: t[1].lower())
    return found


# ============================================================
# 主界面
# ============================================================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(1080, 620)

        self.folder = tk.StringVar()
        self.records = []          # 全部记录
        self.records_by_iid = {}   # tree item id -> record
        self.last_output = None

        # 选项
        self.var_recursive = tk.BooleanVar(value=False)
        self.var_auto_event = tk.BooleanVar(value=True)
        self.var_sort_date = tk.BooleanVar(value=True)
        self.var_only_warn = tk.BooleanVar(value=False)

        self._editor = None        # 当前单元格编辑器

        self._build_style()
        self._build_ui()
        self._bind_shortcuts()
        self.log("欢迎使用照片日程录入工具。请先选择图片文件夹，然后点击「开始扫描」。")

    # ---------------- 样式 ----------------
    def _build_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", font=FONT_SMALL, rowheight=24)
        style.configure("Treeview.Heading", font=FONT_BOLD)
        style.configure("TButton", font=FONT_BASE, padding=4)
        style.configure("TCheckbutton", font=FONT_BASE)
        style.configure("TLabel", font=FONT_BASE)
        style.configure("Accent.TButton", font=FONT_BOLD)

    # ---------------- 界面搭建 ----------------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # --- 1. 文件夹选择 ---
        frm_folder = ttk.LabelFrame(self.root, text=" 1. 图片文件夹 ")
        frm_folder.pack(fill="x", **pad)

        ttk.Label(frm_folder, text="路径：").pack(side="left", padx=(8, 2))
        ent_path = ttk.Entry(frm_folder, textvariable=self.folder, font=FONT_BASE)
        ent_path.pack(side="left", fill="x", expand=True, padx=4, pady=6)
        ent_path.configure(state="readonly")

        ttk.Button(frm_folder, text="浏览…", command=self.choose_folder).pack(side="left", padx=4)
        ttk.Button(frm_folder, text="打开文件夹", command=self.open_folder).pack(side="left", padx=(0, 8))

        # --- 2. 选项与操作 ---
        frm_opt = ttk.LabelFrame(self.root, text=" 2. 扫描选项 ")
        frm_opt.pack(fill="x", **pad)

        ttk.Checkbutton(frm_opt, text="包含子文件夹", variable=self.var_recursive).pack(side="left", padx=(8, 6))
        ttk.Checkbutton(frm_opt, text="自动识别事件类型", variable=self.var_auto_event).pack(side="left", padx=6)
        ttk.Checkbutton(frm_opt, text="按日期排序", variable=self.var_sort_date).pack(side="left", padx=6)
        ttk.Checkbutton(frm_opt, text="只看需确认", variable=self.var_only_warn,
                        command=self.refresh_table).pack(side="left", padx=6)

        ttk.Separator(frm_opt, orient="vertical").pack(side="left", fill="y", padx=10, pady=4)

        ttk.Button(frm_opt, text="开始扫描", style="Accent.TButton",
                   command=self.scan).pack(side="left", padx=6)
        ttk.Button(frm_opt, text="导出 Excel", style="Accent.TButton",
                   command=self.export).pack(side="left", padx=6)
        ttk.Button(frm_opt, text="删除选中行", command=self.delete_selected).pack(side="left", padx=6)
        ttk.Button(frm_opt, text="清空", command=self.clear_all).pack(side="left", padx=(0, 8))

        # --- 3. 预览表格 ---
        frm_table = ttk.LabelFrame(self.root, text=" 3. 解析结果预览（双击单元格可修改） ")
        frm_table.pack(fill="both", expand=True, **pad)

        self.tree = ttk.Treeview(frm_table, columns=COL_NAMES, show="headings", selectmode="extended")
        for name, title, width in COLUMNS:
            self.tree.heading(name, text=title)
            anchor = "center" if name in ("date", "country", "city", "status") else "w"
            self.tree.column(name, width=width, minwidth=50, anchor=anchor, stretch=(name in ("filename", "venue")))

        vsb = ttk.Scrollbar(frm_table, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frm_table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm_table.rowconfigure(0, weight=1)
        frm_table.columnconfigure(0, weight=1)

        self.tree.tag_configure("warn", background=COLOR_WARN)
        self.tree.tag_configure("modified", background=COLOR_MODIFIED)
        self.tree.tag_configure("error", background=COLOR_ERROR)

        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # --- 4. 选中行详情 ---
        frm_detail = ttk.LabelFrame(self.root, text=" 选中行详情 ")
        frm_detail.pack(fill="x", **pad)
        self.lbl_detail = ttk.Label(frm_detail, text="（未选中）", font=FONT_SMALL,
                                    wraplength=1280, justify="left")
        self.lbl_detail.pack(fill="x", padx=8, pady=4)

        # --- 5. 日志 ---
        frm_log = ttk.LabelFrame(self.root, text=" 运行日志 ")
        frm_log.pack(fill="x", **pad)
        self.txt_log = tk.Text(frm_log, height=6, font=FONT_SMALL, state="disabled",
                               wrap="word", background="#FAFAFA")
        log_sb = ttk.Scrollbar(frm_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_sb.set)
        self.txt_log.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=6)
        log_sb.pack(side="left", fill="y", pady=6)

        # --- 6. 状态栏 ---
        self.status = tk.StringVar(value="就绪")
        frm_status = ttk.Frame(self.root)
        frm_status.pack(fill="x", side="bottom")
        ttk.Label(frm_status, textvariable=self.status, font=FONT_SMALL).pack(side="left", padx=10, pady=4)

    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.choose_folder())
        self.root.bind("<F5>", lambda e: self.scan())
        self.root.bind("<Control-s>", lambda e: self.export())
        self.root.bind("<Delete>", lambda e: self.delete_selected())

    # ---------------- 日志与状态 ----------------
    def log(self, msg):
        self.txt_log.configure(state="normal")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{ts}] {msg}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def update_status(self):
        if not self.records:
            self.status.set("就绪　|　尚未扫描")
            return
        total = len(self.records)
        warn = sum(1 for r in self.records if r.get("warnings"))
        mod = sum(1 for r in self.records if r.get("modified") and not r.get("warnings"))
        nodate = sum(1 for r in self.records if not r.get("date"))
        self.status.set(
            f"共 {total} 张　|　正常 {total - warn - mod}　|　需确认 {warn}　|　"
            f"已修改 {mod}　|　缺日期 {nodate}"
        )

    # ---------------- 文件夹操作 ----------------
    def choose_folder(self):
        initial = self.folder.get() or BASE_DIR
        chosen = filedialog.askdirectory(title="请选择包含照片的文件夹", initialdir=initial)
        if chosen:
            self.folder.set(os.path.normpath(chosen))
            self.log(f"已选择文件夹：{self.folder.get()}")

    def open_folder(self):
        target = self.folder.get()
        if not target or not os.path.isdir(target):
            if self.last_output and os.path.isdir(os.path.dirname(self.last_output)):
                target = os.path.dirname(self.last_output)
            else:
                messagebox.showinfo("提示", "还没有可用的文件夹。")
                return
        try:
            os.startfile(target)  # noqa: S606 - Windows 打开资源管理器
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("打开失败", str(exc))

    # ---------------- 扫描 ----------------
    def scan(self):
        folder = self.folder.get()
        if not folder:
            messagebox.showwarning("提示", "请先选择图片文件夹。")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("错误", f"文件夹不存在：\n{folder}")
            return

        self.status.set("正在扫描…")
        self.root.update_idletasks()

        files = collect_image_files(folder, self.var_recursive.get())
        if not files:
            self.records = []
            self.records_by_iid = {}
            self.refresh_table()
            self.update_status()
            messagebox.showwarning("提示", "所选文件夹中没有找到图片文件。")
            self.log("扫描结束：未找到图片文件。")
            return

        records = []
        seq = 0
        for i, (dirpath, fname) in enumerate(files, 1):
            rec = parse_filename(fname)
            if not self.var_auto_event.get():
                rec["event_col"] = None
                rec["event_text"] = None
                rec["warnings"] = [w for w in rec["warnings"]
                                   if not w.startswith("事件无法自动归类")]
            # 子文件夹时文件名带上相对路径，便于定位
            rel = os.path.relpath(dirpath, folder) if self.var_recursive.get() else ""
            rec["_filename"] = os.path.join(rel, fname) if rel and rel != "." else fname
            rec["_dir"] = dirpath
            seq += 1
            rec["_id"] = f"r{seq:06d}"
            records.append(rec)

            if i % 200 == 0:
                self.status.set(f"正在解析 {i}/{len(files)} …")
                self.root.update_idletasks()

        if self.var_sort_date.get():
            records.sort(key=lambda r: (r["date"] is None,
                                        r["date"] or datetime.date.max,
                                        r.get("_filename", "")))

        self.records = records
        self.records_by_iid = {r["_id"]: r for r in records}
        self.refresh_table()
        self.update_status()

        warn = sum(1 for r in records if r.get("warnings"))
        self.log(f"扫描完成：共 {len(records)} 张图片，其中 {warn} 张需要人工确认。")
        self.status.set(f"扫描完成　|　{len(records)} 张，{warn} 张需确认")
        if warn:
            self.log("提示：黄色行为需确认，可在表格中双击单元格直接修改。")

    # ---------------- 表格刷新 ----------------
    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        only_warn = self.var_only_warn.get()
        shown = 0
        for rec in self.records:
            if only_warn and not rec.get("warnings"):
                continue
            iid = rec["_id"]
            self.tree.insert("", "end", iid=iid, values=record_to_row(rec), tags=(tag_of(rec),))
            shown += 1
        if only_warn:
            self.log(f"已过滤：仅显示需确认的 {shown} 行。")
        self.update_status()

    def refresh_row(self, rec):
        iid = rec.get("_id")
        if not iid or not self.tree.exists(iid):
            return
        self.tree.item(iid, values=record_to_row(rec), tags=(tag_of(rec),))

    # ---------------- 单元格编辑 ----------------
    def on_double_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        col_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not col_id or not row_id:
            return
        try:
            col_index = int(col_id.replace("#", "")) - 1
        except ValueError:
            return
        if col_index < 0 or col_index >= len(COL_NAMES):
            return
        col_name = COL_NAMES[col_index]
        if col_name in READONLY_COLS:
            self.log(f"「{dict((c[0], c[1]) for c in COLUMNS)[col_name]}」列不可编辑。")
            return

        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        self.destroy_editor()
        current = self.tree.set(row_id, col_name)
        var = tk.StringVar(value=current)
        entry = ttk.Entry(self.tree, textvariable=var, font=FONT_SMALL)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.select_range(0, "end")

        self._editor = (entry, var, row_id, col_name)
        entry.bind("<Return>", self.commit_edit)
        entry.bind("<KP_Enter>", self.commit_edit)
        entry.bind("<Escape>", self.destroy_editor)
        entry.bind("<FocusOut>", self.commit_edit)

    def commit_edit(self, event=None):
        if not self._editor:
            return
        entry, var, row_id, col_name = self._editor
        new_value = var.get()
        self.destroy_editor()

        rec = self.records_by_iid.get(row_id)
        if rec is None:
            return
        old_value = record_to_row(rec)[COL_NAMES.index(col_name)]
        if new_value.strip() == old_value.strip():
            return

        apply_edit(rec, col_name, new_value)
        self.refresh_row(rec)
        self.update_status()
        self.show_detail(rec)
        title = dict((c[0], c[1]) for c in COLUMNS)[col_name]
        self.log(f"已修改「{rec.get('_filename')}」的 {title}：{old_value or '(空)'} → {new_value.strip() or '(空)'}")

    def destroy_editor(self, event=None):
        if not self._editor:
            return
        entry = self._editor[0]
        self._editor = None
        try:
            entry.destroy()
        except tk.TclError:
            pass

    # ---------------- 选中行详情 ----------------
    def on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            self.lbl_detail.configure(text="（未选中）")
            return
        rec = self.records_by_iid.get(sel[0])
        if rec:
            self.show_detail(rec)

    def show_detail(self, rec):
        lines = [f"文件：{rec.get('_filename', '')}"]
        warns = list(rec.get("warnings", [])) + list(rec.get("_edit_warnings", []))
        if warns:
            lines.append("需确认：" + "；".join(warns))
        else:
            lines.append("状态：解析正常" + ("（已手动修改）" if rec.get("modified") else ""))
        self.lbl_detail.configure(text="　|　".join(lines))

    # ---------------- 行操作 ----------------
    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno("确认", f"确定从导出列表中移除选中的 {len(sel)} 行吗？\n（不会删除任何图片文件）"):
            return
        removed = 0
        for iid in sel:
            rec = self.records_by_iid.pop(iid, None)
            if rec in self.records:
                self.records.remove(rec)
                removed += 1
        self.refresh_table()
        self.log(f"已从导出列表移除 {removed} 行。")

    def clear_all(self):
        if not self.records:
            return
        if not messagebox.askyesno("确认", "清空当前所有解析结果？"):
            return
        self.records = []
        self.records_by_iid = {}
        self.refresh_table()
        self.lbl_detail.configure(text="（未选中）")
        self.log("已清空解析结果。")

    # ---------------- 导出 ----------------
    def export(self):
        if not self.records:
            messagebox.showwarning("提示", "还没有可导出的数据，请先扫描文件夹。")
            return

        folder = self.folder.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("错误", "输出文件夹不可用，请重新选择图片文件夹。")
            return

        warn_count = sum(1 for r in self.records if r.get("warnings"))
        if warn_count and not messagebox.askyesno(
                "存在待确认项",
                f"还有 {warn_count} 行标记为「需确认」。\n\n"
                "仍然导出吗？（这些行会在 Excel 中标黄）"):
            return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        xlsx_path = os.path.join(folder, f"日程记录_{timestamp}.xlsx")
        warning_path = os.path.join(folder, f"待人工确认_{timestamp}.txt")

        try:
            generate_excel(self.records, xlsx_path)
        except PermissionError:
            messagebox.showerror("导出失败", "文件被占用，请先关闭已打开的同名 Excel 再试。")
            return
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("导出失败", str(exc))
            return

        self.last_output = xlsx_path
        self.log(f"已导出 Excel：{xlsx_path}")

        if warn_count:
            try:
                generate_warning_list(self.records, warning_path)
                self.log(f"已导出问题清单：{warning_path}")
            except Exception as exc:  # pragma: no cover
                self.log(f"问题清单写入失败：{exc}")

        self.update_status()
        messagebox.showinfo(
            "导出完成",
            f"共导出 {len(self.records)} 条记录。\n\n"
            f"Excel：\n{xlsx_path}"
            + (f"\n\n问题清单：\n{warning_path}" if warn_count else "")
        )
        if messagebox.askyesno("导出完成", "是否打开所在文件夹？"):
            self.open_folder()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
