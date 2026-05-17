#!/usr/bin/env python3
"""XLS → XLSX 変換ツール（GUIワンボタン版）"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import xlrd
import openpyxl


def convert(src: str, dst: str | None = None) -> str:
    if dst is None:
        dst = os.path.splitext(src)[0] + ".xlsx"

    wb_xls = xlrd.open_workbook(src, formatting_info=True)
    wb_xlsx = openpyxl.Workbook()
    wb_xlsx.remove(wb_xlsx.active)

    for sheet_name in wb_xls.sheet_names():
        ws_xls = wb_xls.sheet_by_name(sheet_name)
        ws_xlsx = wb_xlsx.create_sheet(title=sheet_name)

        for col_idx, col_info in ws_xls.colinfo_map.items():
            if col_info.width > 0:
                col_letter = openpyxl.utils.get_column_letter(col_idx + 1)
                ws_xlsx.column_dimensions[col_letter].width = col_info.width / 256

        for row_idx, row_info in ws_xls.rowinfo_map.items():
            if row_info.height > 0:
                ws_xlsx.row_dimensions[row_idx + 1].height = row_info.height / 20

        for row_idx in range(ws_xls.nrows):
            for col_idx in range(ws_xls.ncols):
                cell = ws_xls.cell(row_idx, col_idx)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(value, wb_xls.datemode)
                    except Exception:
                        pass
                ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1, value=value)

    wb_xlsx.save(dst)
    return dst


def verify(src: str, dst: str) -> list[str]:
    """変換前後のファイルを比較し、差異のリストを返す。空リストなら一致。"""
    diffs = []

    wb_xls = xlrd.open_workbook(src, formatting_info=True)
    wb_xlsx = openpyxl.load_workbook(dst)

    xls_sheets = wb_xls.sheet_names()
    xlsx_sheets = wb_xlsx.sheetnames

    if xls_sheets != xlsx_sheets:
        diffs.append(f"シート構成が異なります:\n  変換前: {xls_sheets}\n  変換後: {xlsx_sheets}")
        return diffs

    for sheet_name in xls_sheets:
        ws_xls = wb_xls.sheet_by_name(sheet_name)
        ws_xlsx = wb_xlsx[sheet_name]

        for row_idx in range(ws_xls.nrows):
            for col_idx in range(ws_xls.ncols):
                cell = ws_xls.cell(row_idx, col_idx)
                xls_val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        xls_val = xlrd.xldate_as_datetime(xls_val, wb_xls.datemode)
                    except Exception:
                        pass

                xlsx_val = ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1).value

                # 空文字とNoneは同一扱い
                xls_norm = None if xls_val == "" else xls_val
                xlsx_norm = None if xlsx_val == "" else xlsx_val
                if xls_norm != xlsx_norm:
                    loc = f"シート「{sheet_name}」行{row_idx+1} 列{col_idx+1}"
                    diffs.append(f"{loc}: 変換前={xls_val!r} / 変換後={xlsx_val!r}")

        for col_idx, col_info in ws_xls.colinfo_map.items():
            if col_info.width == 0:
                continue
            expected = round(col_info.width / 256, 2)
            letter = openpyxl.utils.get_column_letter(col_idx + 1)
            actual = round(ws_xlsx.column_dimensions[letter].width, 2)
            if abs(expected - actual) >= 0.1:
                diffs.append(
                    f"シート「{sheet_name}」列{letter} 列幅: 変換前={expected} / 変換後={actual}"
                )

    return diffs


def on_click():
    path = filedialog.askopenfilename(
        title="変換するXLSファイルを選択",
        filetypes=[("Excel 97-2003", "*.xls")],
    )
    if not path:
        return

    try:
        output = convert(path)
        diffs = verify(path, output)

        if diffs:
            detail = "\n".join(f"・{d}" for d in diffs[:10])
            if len(diffs) > 10:
                detail += f"\n... 他 {len(diffs) - 10} 件"
            messagebox.showwarning(
                "警告: 差異が見つかりました",
                f"変換は完了しましたが、{len(diffs)} 件の差異が検出されました。\n\n{detail}"
            )
        else:
            messagebox.showinfo(
                "完了",
                f"変換・検証が完了しました！\n\n"
                f"保存先: {output}\n\n"
                f"✓ 変換前後のデータは完全に一致しています。"
            )
    except Exception as e:
        messagebox.showerror("エラー", str(e))


root = tk.Tk()
root.title("XLS → XLSX 変換ツール")
root.resizable(False, False)

frame = tk.Frame(root, padx=40, pady=30)
frame.pack()

tk.Label(frame, text="XLSファイルをXLSXに変換します", font=("", 12)).pack(pady=(0, 20))

btn = tk.Button(
    frame,
    text="ファイルを選んで変換",
    font=("", 13, "bold"),
    bg="#4CAF50",
    fg="white",
    padx=20,
    pady=10,
    cursor="hand2",
    command=on_click,
)
btn.pack()

tk.Label(frame, text="変換後のファイルは元ファイルと同じフォルダに保存されます", font=("", 9), fg="gray").pack(pady=(15, 0))

root.mainloop()
