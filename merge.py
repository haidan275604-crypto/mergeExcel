# -*- coding: utf-8 -*-
"""
a、b 两个表结构完全一致，逐行合并第 5 列：
  · 同一行 a、b 第 5 列都有数据 → 重合，保留 b 原值不动，记录行号
  · 同一行 b 第 5 列为空、a 有值   → 用 a 的值填入
  · 其他情况不做处理
合并完成后打印重合的行号
"""

from openpyxl import load_workbook


def merge_column(a_path, b_path, out_path,
                 col=5, has_header=True,
                 sheet_a=0, sheet_b=0):
    wb_a = load_workbook(a_path, data_only=True)
    wb_b = load_workbook(b_path)

    ws_a = wb_a.worksheets[sheet_a] if isinstance(sheet_a, int) else wb_a[sheet_a]
    ws_b = wb_b.worksheets[sheet_b] if isinstance(sheet_b, int) else wb_b[sheet_b]

    first_row = 2 if has_header else 1
    last_row  = min(ws_a.max_row, ws_b.max_row)   # 结构一致，取二者较短者更保险

    duplicate_rows = []   # 重合行号
    filled_rows    = []   # 成功填充的行号

    for r in range(first_row, last_row + 1):
        a_val = ws_a.cell(row=r, column=col).value
        b_val = ws_b.cell(row=r, column=col).value

        a_empty = a_val is None or str(a_val).strip() == ""
        b_empty = b_val is None or str(b_val).strip() == ""

        if a_empty:
            continue                     # a 没数据，跳过
        if not b_empty:
            duplicate_rows.append(r)     # 两边都有数据 → 重合，b 保持不动
        else:
            ws_b.cell(row=r, column=col, value=a_val)   # b 为空 → 填入
            filled_rows.append(r)

    wb_b.save(out_path)

    # ---------- 结果提示 ----------
    print(f"【完成】共向 B 文件第 {col} 列补充了 {len(filled_rows)} 行数据。")
    if filled_rows:
        print(f"     填充行号：{filled_rows}")

    if duplicate_rows:
        print(f"【重合】共 {len(duplicate_rows)} 行 A、B 第 {col} 列都有数据（B 未改动）：")
        for r in duplicate_rows:
            print(f"    第 {r} 行 → A: {ws_a.cell(row=r, column=col).value!r}  |  "
                  f"B(保留): {ws_b.cell(row=r, column=col).value!r}")
    else:
        print("【重合】未发现重合数据。")


if __name__ == "__main__":
    merge_column(
        a_path="a.xlsx",
        b_path="b.xlsx",
        out_path="b_合并后.xlsx",
        col=5,
        has_header=True,     # 没有表头就改成 False
    )