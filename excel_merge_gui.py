# -*- coding: utf-8 -*-
"""
Excel 第五列按行合并工具（美化自适应版）
"""
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook
import zipfile


# ---------------- 核心合并逻辑 ----------------
def merge_column(a_path, b_path, out_path, col=5, has_header=True):
    if not a_path.lower().endswith(".xlsx") or not b_path.lower().endswith(".xlsx"):
        raise ValueError(
            "文件格式错误：请确保 A 文件和 B 文件都是 .xlsx 格式！\n\n如果是 .xls 文件，请先用 Excel 另存为 .xlsx")

    try:
        wb_a = load_workbook(a_path, data_only=True)
        wb_b = load_workbook(b_path)
    except zipfile.BadZipFile:
        raise ValueError(
            "文件读取失败！\n\n可能的原因：\n1. 文件不是标准的 xlsx 格式\n2. 文件已损坏\n3. 文件设置了密码\n4. 文件正被其他软件（如Excel/WPS）占用。")
    except Exception as e:
        raise ValueError(f"读取文件时发生错误：\n{e}")

    ws_a = wb_a.worksheets[0]
    ws_b = wb_b.worksheets[0]

    first_row = 2 if has_header else 1
    last_row = min(ws_a.max_row, ws_b.max_row)

    duplicate_rows = []
    filled_rows = []

    for r in range(first_row, last_row + 1):
        a_val = ws_a.cell(row=r, column=col).value
        b_val = ws_b.cell(row=r, column=col).value

        a_empty = a_val is None or str(a_val).strip() == ""
        b_empty = b_val is None or str(b_val).strip() == ""

        if a_empty:
            continue
        if not b_empty:
            # 两边都有数据 → 合并数据，用顿号连接
            merged_val = f"{b_val}、{a_val}"
            ws_b.cell(row=r, column=col, value=merged_val)
            duplicate_rows.append((r, b_val, a_val, merged_val))
        else:
            # b 为空，a 有数据 → 直接填入
            ws_b.cell(row=r, column=col, value=a_val)
            filled_rows.append(r)

    try:
        wb_b.save(out_path)
    except PermissionError:
        raise ValueError(f"保存失败！\n\n输出文件正被其他程序占用，请关闭该文件或换一个输出文件名。")

    return filled_rows, duplicate_rows


# ---------------- GUI 主界面 ----------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Excel 第五列合并工具")
        root.geometry("600x400")
        # ★ 允许窗口自由缩放
        root.resizable(True, True)
        root.minsize(550, 350)

        # 全局样式美化
        self.style = ttk.Style()
        # 尝试使用更现代的主题 (clam 在 Windows/Mac 上表现都还可以)
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        # 自定义按钮样式
        self.style.configure("Action.TButton", font=("微软雅黑", 11, "bold"), padding=6)
        self.style.configure("Browse.TButton", padding=2)
        self.style.configure("TLabel", font=("微软雅黑", 10))
        self.style.configure("TEntry", padding=4)

        self.a_path = tk.StringVar()
        self.b_path = tk.StringVar()
        self.out_path = tk.StringVar()
        self.has_header = tk.BooleanVar(value=True)
        self.col_var = tk.StringVar(value="5")

        self.build_ui()

    def build_ui(self):
        # 主容器，留出边距，随窗口拉伸
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 文件选择区域
        self._create_file_row(main_frame, 0, "A 文件 (提供数据)", self.a_path, self.pick_a)
        self._create_file_row(main_frame, 1, "B 文件 (目标文件)", self.b_path, self.pick_b)
        self._create_file_row(main_frame, 2, "输出保存为", self.out_path, self.pick_out)

        # 2. 选项区域
        opt_frame = ttk.Frame(main_frame)
        opt_frame.grid(row=3, column=0, columnspan=3, pady=(15, 5), sticky="w")

        ttk.Label(opt_frame, text="操作列号 (1-9):").pack(side=tk.LEFT)
        spinbox = ttk.Spinbox(opt_frame, from_=1, to=9, textvariable=self.col_var, width=5)
        spinbox.pack(side=tk.LEFT, padx=(5, 20))

        ttk.Checkbutton(opt_frame, text="第一行是表头 (从第二行开始合并)", variable=self.has_header).pack(side=tk.LEFT)

        # 3. 开始按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=20)
        ttk.Button(btn_frame, text="🚀 开始合并", style="Action.TButton", command=self.start).pack()

        # 4. 状态栏
        self.status = ttk.Label(main_frame, text="请选择文件后点击「开始合并」", foreground="gray")
        self.status.grid(row=5, column=0, columnspan=3, sticky="s")

        # 配置 grid 权重，确保拉伸时输入框变宽
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)  # 把状态栏往下推

    def _create_file_row(self, parent, row, label_text, var, cmd):
        ttk.Label(parent, text=label_text, width=16, anchor="w").grid(row=row, column=0, pady=8, sticky="w")
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, padx=10, sticky="ew")
        ttk.Button(parent, text="选择…", style="Browse.TButton", command=cmd).grid(row=row, column=2, sticky="e")

    def pick_a(self):
        p = filedialog.askopenfilename(title="选择 A 文件", filetypes=[("Excel 文件", "*.xlsx")])
        if p:
            self.a_path.set(p)

    def pick_b(self):
        p = filedialog.askopenfilename(title="选择 B 文件", filetypes=[("Excel 文件", "*.xlsx")])
        if p:
            self.b_path.set(p)
            if not self.out_path.get():
                base, ext = os.path.splitext(p)
                self.out_path.set(base + "_合并后" + ext)

    def pick_out(self):
        p = filedialog.asksaveasfilename(title="保存为", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if p:
            self.out_path.set(p)

    def start(self):
        a, b, o = self.a_path.get(), self.b_path.get(), self.out_path.get()
        if not (a and b and o):
            messagebox.showwarning("提示", "请把 A、B 文件和输出路径都选好。")
            return
        try:
            col = int(self.col_var.get())
        except ValueError:
            messagebox.showerror("错误", "列号必须是整数！")
            return

        try:
            self.status.config(text="正在合并中，请稍候…", foreground="blue")
            self.root.update()
            filled, dup = merge_column(a, b, o, col, self.has_header.get())
        except Exception as e:
            messagebox.showerror("出错了", str(e))
            self.status.config(text="合并失败，请检查文件。", foreground="red")
            return

        # ★ 成功之后：清空输入框
        self.a_path.set("")
        self.b_path.set("")
        self.out_path.set("")
        self.status.config(text="✅ 合并成功！输入已清空，可进行下一次操作。", foreground="green")

        # 弹出结果展示窗
        ResultWindow(self.root, o, filled, dup)


# ---------------- 结果展示窗口 (美化版) ----------------
class ResultWindow:
    def __init__(self, parent, out_path, filled, dup):
        self.win = tk.Toplevel(parent)
        self.win.title("处理完成")
        self.win.geometry("650x450")
        self.win.transient(parent)  # 置顶在主窗口之上

        main_f = ttk.Frame(self.win, padding=15)
        main_f.pack(fill=tk.BOTH, expand=True)

        # 头部提示
        ttk.Label(main_f, text=f"文件已保存至：\n{out_path}", font=("微软雅黑", 10, "bold")).pack(anchor="w",
                                                                                                 pady=(0, 10))

        info_text = f"共向空单元格补充了 {len(filled)} 行数据。"
        if dup:
            info_text += f" 且发现 {len(dup)} 行重合数据，已用「、」合并。"
        else:
            info_text += " 未发现重合数据。"

        ttk.Label(main_f, text=info_text, foreground="#333").pack(anchor="w", pady=(0, 10))

        # 如果有重合数据，展示表格
        if dup:
            ttk.Label(main_f, text="重合数据明细：").pack(anchor="w")

            # 创建 Treeview 表格展示
            cols = ("行号", "B 文件原数据", "A 文件新数据", "合并后结果")
            tree = ttk.Treeview(main_f, columns=cols, show="headings", height=10)

            for c in cols:
                tree.heading(c, text=c)
                tree.column(c, width=100, anchor="center")

            tree.column("合并后结果", width=200)
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)

            # 滚动条
            scrollbar = ttk.Scrollbar(main_f, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscroll=scrollbar.set)
            scrollbar.pack(side=tk.LEFT, fill=tk.Y, pady=5)

            # 插入数据
            for r, old_b, old_a, new_val in dup:
                tree.insert("", tk.END, values=(r, old_b, old_a, new_val))

        # 关闭按钮
        ttk.Button(main_f, text="我知道了", command=self.win.destroy).pack(side=tk.BOTTOM, pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()