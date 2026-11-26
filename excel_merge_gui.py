# -*- coding: utf-8 -*-
"""
Excel 第五列按行合并工具（GUI 版）
打包命令见文末
"""
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook


# ---------------- 核心合并逻辑 ----------------
def merge_column(a_path, b_path, out_path, col=5, has_header=True):
    wb_a = load_workbook(a_path, data_only=True)
    wb_b = load_workbook(b_path)
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
            duplicate_rows.append((r, a_val, b_val))
        else:
            ws_b.cell(row=r, column=col, value=a_val)
            filled_rows.append(r)

    wb_b.save(out_path)
    return filled_rows, duplicate_rows


# ---------------- GUI ----------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Excel 第五列合并工具")
        root.geometry("560x320")
        root.resizable(False, False)

        self.a_path = tk.StringVar()
        self.b_path = tk.StringVar()
        self.out_path = tk.StringVar()
        self.has_header = tk.BooleanVar(value=True)
        self.col_var = tk.StringVar(value="5")

        self._row("A 文件（提供数据）", self.a_path, self.pick_a)
        self._row("B 文件（目标文件）", self.b_path, self.pick_b)
        self._row("输出文件", self.out_path, self.pick_out)

        opt = tk.Frame(root, pady=6)
        opt.pack(fill="x", padx=12)
        tk.Label(opt, text="操作列号：").pack(side="left")
        tk.Entry(opt, textvariable=self.col_var, width=5).pack(side="left")
        tk.Checkbutton(opt, text="第一行是表头", variable=self.has_header).pack(side="left", padx=10)

        tk.Button(root, text="开始合并", width=20, height=2,
                  bg="#4CAF50", fg="white", font=("微软雅黑", 11, "bold"),
                  command=self.start).pack(pady=12)

        self.status = tk.Label(root, text="请选择文件后点击「开始合并」", fg="gray")
        self.status.pack()

    def _row(self, label, var, cmd):
        f = tk.Frame(self.root, pady=4)
        f.pack(fill="x", padx=12)
        tk.Label(f, text=label, width=16, anchor="w").pack(side="left")
        tk.Entry(f, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(f, text="选择…", command=cmd).pack(side="left")

    def pick_a(self):
        p = filedialog.askopenfilename(title="选择 A 文件",
                                       filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p:
            self.a_path.set(p)

    def pick_b(self):
        p = filedialog.askopenfilename(title="选择 B 文件",
                                       filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p:
            self.b_path.set(p)
            if not self.out_path.get():
                base, ext = os.path.splitext(p)
                self.out_path.set(base + "_合并后" + ext)

    def pick_out(self):
        p = filedialog.asksaveasfilename(title="保存为",
                                         defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            self.out_path.set(p)

    def start(self):
        a, b, o = self.a_path.get(), self.b_path.get(), self.out_path.get()
        if not (a and b and o):
            messagebox.showwarning("提示", "请先把 A、B、输出三个路径都选好。")
            return
        try:
            col = int(self.col_var.get())
        except ValueError:
            messagebox.showerror("错误", "列号必须是整数")
            return

        try:
            self.status.config(text="合并中，请稍候…")
            self.root.update()
            filled, dup = merge_column(a, b, o, col, self.has_header.get())
        except Exception as e:
            messagebox.showerror("出错了", f"{e}\n\n{traceback.format_exc()}")
            self.status.config(text="失败")
            return

        self.status.config(text=f"完成，已保存到：{o}")
        self.show_result(o, filled, dup)

    def show_result(self, out_path, filled, dup):
        win = tk.Toplevel(self.root)
        win.title("合并结果")
        win.geometry("600x420")
        txt = tk.Text(win, wrap="word")
        txt.pack(fill="both", expand=True)

        txt.insert("end", f"输出文件：{out_path}\n\n")
        txt.insert("end", f"共补充 {len(filled)} 行数据：{filled}\n\n")
        if dup:
            txt.insert("end", f"★ 重合 {len(dup)} 行（B 文件原数据已保留）：\n")
            for r, av, bv in dup:
                txt.insert("end", f"    第 {r} 行 → A: {av!r}  |  B(保留): {bv!r}\n")
        else:
            txt.insert("end", "未发现重合数据。\n")
        txt.config(state="disabled")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()