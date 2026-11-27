# -*- coding: utf-8 -*-
"""
多文件 Excel 指定行/列合并工具 (终极智能去重版)
"""
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook
import zipfile


# ---------------- 辅助函数：智能拆分名字 ----------------
def parse_names(val):
    """将单元格数据统一拆分，支持顿号、中英文逗号，并去除空格和空字符串"""
    if val is None:
        return []
    # 把中英文逗号统一替换为顿号，去掉所有空格
    s = str(val).replace(",", "、").replace("，", "、").replace(" ", "").strip()
    # 按顿号拆分并过滤空串
    return [x for x in s.split("、") if x]


# ---------------- 核心合并逻辑 ----------------
def merge_column(a_paths, b_path, out_path, col=5, start_row=2):
    if not b_path.lower().endswith(".xlsx") or not all(p.lower().endswith(".xlsx") for p in a_paths):
        raise ValueError("文件格式错误：请确保 A、B 文件都是 .xlsx 格式！")

    try:
        wb_b = load_workbook(b_path)
    except zipfile.BadZipFile:
        raise ValueError("B 文件读取失败！可能被加密、损坏或正在被占用。")
    except Exception as e:
        raise ValueError(f"读取 B 文件时发生错误：\n{e}")

    ws_b = wb_b.worksheets[0]
    max_row_b = ws_b.max_row

    # 记录历史情况
    history = {}

    for a_path in a_paths:
        try:
            wb_a = load_workbook(a_path, data_only=True)
        except Exception as e:
            raise ValueError(f"读取 A 文件 [{os.path.basename(a_path)}] 失败：\n{e}")

        ws_a = wb_a.worksheets[0]
        last_row = min(ws_a.max_row, max_row_b)
        a_name = os.path.basename(a_path)

        for r in range(start_row, last_row + 1):
            a_val = ws_a.cell(row=r, column=col).value
            b_val = ws_b.cell(row=r, column=col).value

            if a_val is None or str(a_val).strip() == "":
                continue

            # 拆分当前 B 单元格为列表
            current_b_list = parse_names(b_val)

            # 初始化该行的记录
            if r not in history:
                history[r] = {
                    "original": "、".join(current_b_list),
                    "current": "、".join(current_b_list),
                    "logs": []
                }

            # 拆分当前 A 单元格为列表
            incoming_a_list = parse_names(a_val)

            # 找出 A 中 B 没有的新元素
            added_items = []
            for item in incoming_a_list:
                if item not in current_b_list:
                    current_b_list.append(item)
                    added_items.append(item)

            # 如果 A 里的元素全都在 B 里了，直接跳过
            if not added_items:
                history[r]["logs"].append(f"{a_name} 已存在(跳过)")
                continue

            # 如果有新元素，重新组合并写回 Excel
            merged_val = "、".join(current_b_list)
            ws_b.cell(row=r, column=col, value=merged_val)

            history[r]["current"] = merged_val
            history[r]["logs"].append(f"从 {a_name} 添加了: {'、'.join(added_items)}")

    try:
        wb_b.save(out_path)
    except PermissionError:
        raise ValueError("保存失败！输出文件正被占用，请关闭后重试。")

    # 整理最终展示数据
    results = []
    for r, data in history.items():
        if data["logs"]:
            results.append({
                "row": r,
                "original": data["original"],
                "current": data["current"],
                "logs": " | ".join(data["logs"])
            })

    results.sort(key=lambda x: x["row"])
    return results


# ---------------- GUI 主界面 ----------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("多文件 Excel 行/列合并工具")
        root.geometry("720x550")
        root.resizable(True, True)
        root.minsize(650, 480)

        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.style.configure("Action.TButton", font=("微软雅黑", 11, "bold"), padding=6)
        self.style.configure("Browse.TButton", padding=2)
        self.style.configure("TLabel", font=("微软雅黑", 10))
        self.style.configure("TEntry", padding=4)
        self.style.configure("TButton", font=("微软雅黑", 9))

        self.b_path = tk.StringVar()
        self.out_path = tk.StringVar()

        self.start_row_var = tk.StringVar(value="2")
        self.col_var = tk.StringVar(value="5")

        self.build_ui()

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 1. A 文件列表区域 ---
        a_frame = ttk.Frame(main_frame)
        a_frame.grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky="ew")

        ttk.Label(a_frame, text="A 文件列表(按序合并):", width=18, anchor="w").pack(side=tk.LEFT, anchor="n")

        list_frame = ttk.Frame(a_frame)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.a_listbox = tk.Listbox(list_frame, height=6, font=("微软雅黑", 9), selectmode=tk.EXTENDED)
        self.a_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        a_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.a_listbox.yview)
        self.a_listbox.configure(yscrollcommand=a_scroll.set)
        a_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        btn_a_frame = ttk.Frame(a_frame)
        btn_a_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0))
        ttk.Button(btn_a_frame, text="➕ 添加", command=self.add_a_files).pack(fill=tk.X, pady=2)
        ttk.Button(btn_a_frame, text="🗑️ 清空", command=self.clear_a_files).pack(fill=tk.X, pady=2)

        # --- 2. B 文件及输出区域 ---
        self._create_file_row(main_frame, 1, "B 文件 (目标文件)", self.b_path, self.pick_b)
        self._create_file_row(main_frame, 2, "输出保存为", self.out_path, self.pick_out)

        # --- 3. 选项区域 ---
        opt_frame = ttk.Frame(main_frame)
        opt_frame.grid(row=3, column=0, columnspan=3, pady=(20, 5), sticky="w")

        ttk.Label(opt_frame, text="起始行号:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_frame, from_=1, to=9999, textvariable=self.start_row_var, width=6).pack(side=tk.LEFT,
                                                                                                padx=(5, 20))

        ttk.Label(opt_frame, text="操作列号:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_frame, from_=1, to=9, textvariable=self.col_var, width=5).pack(side=tk.LEFT, padx=(5, 0))

        # --- 4. 开始按钮 ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=30)
        ttk.Button(btn_frame, text="🚀 开始合并", style="Action.TButton", command=self.start).pack()

        # --- 5. 状态栏 ---
        self.status = ttk.Label(main_frame, text="请添加 A 文件并选择 B 文件，点击开始合并", foreground="gray")
        self.status.grid(row=5, column=0, columnspan=3, sticky="s")

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

    def _create_file_row(self, parent, row, label_text, var, cmd):
        ttk.Label(parent, text=label_text, width=18, anchor="w").grid(row=row, column=0, pady=8, sticky="w")
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, padx=10, sticky="ew")
        ttk.Button(parent, text="选择…", style="Browse.TButton", command=cmd).grid(row=row, column=2, sticky="e")

    def add_a_files(self):
        paths = filedialog.askopenfilenames(title="选择 A 文件（可多选）", filetypes=[("Excel 文件", "*.xlsx")])
        if paths:
            existing = self.a_listbox.get(0, tk.END)
            for p in paths:
                if p not in existing:
                    self.a_listbox.insert(tk.END, p)

    def clear_a_files(self):
        self.a_listbox.delete(0, tk.END)

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
        a_paths = list(self.a_listbox.get(0, tk.END))
        b = self.b_path.get()
        o = self.out_path.get()

        if not a_paths:
            messagebox.showwarning("提示", "请至少添加一个 A 文件。")
            return
        if not b or not o:
            messagebox.showwarning("提示", "请选择 B 文件和输出路径。")
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
            results = merge_column(a_paths, b, o, col, start_row)
        except Exception as e:
            messagebox.showerror("出错了", str(e))
            self.status.config(text="合并失败，请检查文件。", foreground="red")
            return

        self.clear_a_files()
        self.b_path.set("")
        self.out_path.set("")
        self.status.config(text="✅ 合并成功！输入已清空，可进行下一次操作。", foreground="green")

        ResultWindow(self.root, o, results)


# ---------------- 结果展示窗口 ----------------
class ResultWindow:
    def __init__(self, parent, out_path, results):
        self.win = tk.Toplevel(parent)
        self.win.title("处理完成")
        self.win.geometry("800x500")
        self.win.transient(parent)

        main_f = ttk.Frame(self.win, padding=15)
        main_f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_f, text=f"文件已保存至：\n{out_path}", font=("微软雅黑", 10, "bold")).pack(anchor="w",
                                                                                                 pady=(0, 10))

        if results:
            info_text = f"共处理了 {len(results)} 行有变动的单元格（已自动过滤重复项）。"
        else:
            info_text = "文件合并完成，未发现需要新增的数据。"

        ttk.Label(main_f, text=info_text, foreground="#333").pack(anchor="w", pady=(0, 10))

        if results:
            ttk.Label(main_f, text="处理明细（按行汇总）：").pack(anchor="w")

            cols = ("行号", "B文件最初数据", "最终结果", "处理明细")
            tree = ttk.Treeview(main_f, columns=cols, show="headings", height=12)

            widths = (60, 180, 200, 320)
            for c, w in zip(cols, widths):
                tree.heading(c, text=c)
                tree.column(c, width=w, anchor="center")

            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)

            scrollbar = ttk.Scrollbar(main_f, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscroll=scrollbar.set)
            scrollbar.pack(side=tk.LEFT, fill=tk.Y, pady=5)

            for item in results:
                tree.insert("", tk.END, values=(item["row"], item["original"], item["current"], item["logs"]))

        ttk.Button(main_f, text="我知道了", command=self.win.destroy).pack(side=tk.BOTTOM, pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()