# -*- coding: utf-8 -*-
"""
Excel 指定行、指定列合并工具（美化自适应版）
"""
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook
import zipfile


# ---------------- 核心合并逻辑 ----------------
def merge_column(a_path, b_path, out_path, col=5, start_row=2):
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

    last_row = min(ws_a.max_row, ws_b.max_row)

    duplicate_rows = []  # 两边数据不同，已拼接的
    same_rows = []  # 两边数据相同，跳过的
    filled_rows = []  # B 为空，直接填充的

    for r in range(start_row, last_row + 1):
        a_val = ws_a.cell(row=r, column=col).value
        b_val = ws_b.cell(row=r, column=col).value

        a_empty = a_val is None or str(a_val).strip() == ""
        b_empty = b_val is None or str(b_val).strip() == ""

        if a_empty:
            continue

        if not b_empty:
            # 两边都有数据
            str_a = str(a_val).strip()
            str_b = str(b_val).strip()

            if str_a == str_b:
                # ★ 新增：内容完全一致，不做任何修改，保留原样
                same_rows.append((r, b_val, a_val))
            else:
                # 内容不同 → 拼接合并
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

    return filled_rows, duplicate_rows, same_rows


# ---------------- GUI 主界面 ----------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Excel 行/列合并工具")
        root.geometry("620x420")
        # 允许窗口自由缩放
        root.resizable(True, True)
        root.minsize(580, 380)

        # 全局样式美化
        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.style.configure("Action.TButton", font=("微软雅黑", 11, "bold"), padding=6)
        self.style.configure("Browse.TButton", padding=2)
        self.style.configure("TLabel", font=("微软雅黑", 10))
        self.style.configure("TEntry", padding=4)

        self.a_path = tk.StringVar()
        self.b_path = tk.StringVar()
        self.out_path = tk.StringVar()

        # 默认起始行2，操作列号5
        self.start_row_var = tk.StringVar(value="2")
        self.col_var = tk.StringVar(value="5")

        self.build_ui()

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 文件选择区域
        self._create_file_row(main_frame, 0, "A 文件 (提供数据)", self.a_path, self.pick_a)
        self._create_file_row(main_frame, 1, "B 文件 (目标文件)", self.b_path, self.pick_b)
        self._create_file_row(main_frame, 2, "输出保存为", self.out_path, self.pick_out)

        # 2. 选项区域
        opt_frame = ttk.Frame(main_frame)
        opt_frame.grid(row=3, column=0, columnspan=3, pady=(20, 5), sticky="w")

        ttk.Label(opt_frame, text="起始行号:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_frame, from_=1, to=9999, textvariable=self.start_row_var, width=6).pack(side=tk.LEFT,
                                                                                                padx=(5, 20))

        ttk.Label(opt_frame, text="操作列号:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_frame, from_=1, to=9, textvariable=self.col_var, width=5).pack(side=tk.LEFT, padx=(5, 0))

        # 3. 开始按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=30)
        ttk.Button(btn_frame, text="🚀 开始合并", style="Action.TButton", command=self.start).pack()

        # 4. 状态栏
        self.status = ttk.Label(main_frame, text="请选择文件并设置好行/列，点击「开始合并」", foreground="gray")
        self.status.grid(row=5, column=0, columnspan=3, sticky="s")

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

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
            start_row = int(self.start_row_var.get())
            if col < 1 or start_row < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "列号和起始行号必须是大于 0 的整数！")
            return

        try:
            self.status.config(text="正在合并中，请稍候…", foreground="blue")
            self.root.update()
            filled, dup, same = merge_column(a, b, o, col, start_row)
        except Exception as e:
            messagebox.showerror("出错了", str(e))
            self.status.config(text="合并失败，请检查文件。", foreground="red")
            return

        # 成功之后：清空输入框
        self.a_path.set("")
        self.b_path.set("")
        self.out_path.set("")
        self.status.config(text="✅ 合并成功！输入已清空，可进行下一次操作。", foreground="green")

        # 弹出结果展示窗
        ResultWindow(self.root, o, filled, dup, same)


# ---------------- 结果展示窗口 ----------------
class ResultWindow:
    def __init__(self, parent, out_path, filled, dup, same):
        self.win = tk.Toplevel(parent)
        self.win.title("处理完成")
        self.win.geometry("700x480")
        self.win.transient(parent)

        main_f = ttk.Frame(self.win, padding=15)
        main_f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_f, text=f"文件已保存至：\n{out_path}", font=("微软雅黑", 10, "bold")).pack(anchor="w",
                                                                                                 pady=(0, 10))

        # 组合提示文案
        info_text = f"共向空单元格补充了 {len(filled)} 行数据。\n"
        if dup or same:
            if dup:
                info_text += f"发现 {len(dup)} 行数据不一致，已用「、」拼接合并。\n"
            if same:
                info_text += f"发现 {len(same)} 行数据完全一致，已保持原样跳过。"
        else:
            info_text += "未发现重复位置或重合数据。"

        ttk.Label(main_f, text=info_text, foreground="#333").pack(anchor="w", pady=(0, 10))

        # 展示重合数据明细
        if dup or same:
            ttk.Label(main_f, text="重合数据明细：").pack(anchor="w")

            # 新增一列“处理结果”，区分是拼接还是跳过
            cols = ("行号", "B 文件原数据", "A 文件新数据", "处理结果", "状态")
            tree = ttk.Treeview(main_f, columns=cols, show="headings", height=10)

            # 设置列宽
            widths = (60, 150, 150, 200, 100)
            for c, w in zip(cols, widths):
                tree.heading(c, text=c)
                tree.column(c, width=w, anchor="center")

            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)

            scrollbar = ttk.Scrollbar(main_f, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscroll=scrollbar.set)
            scrollbar.pack(side=tk.LEFT, fill=tk.Y, pady=5)

            # 插入“不一致拼接”的记录
            for r, old_b, old_a, new_val in dup:
                tree.insert("", tk.END, values=(r, old_b, old_a, new_val, "拼接合并"))

            # 插入“一致跳过”的记录
            for r, old_b, old_a in same:
                tree.insert("", tk.END, values=(r, old_b, old_a, old_b, "内容一致(保留)"))

        ttk.Button(main_f, text="我知道了", command=self.win.destroy).pack(side=tk.BOTTOM, pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()