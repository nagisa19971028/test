#!/usr/bin/env python3
"""XLS → XLSX 変換ツール（GUIワンボタン版）"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.styles.colors import Color

_HOR_ALIGN = {
    0: "general", 1: "left", 2: "center", 3: "right",
    4: "fill", 5: "justify", 6: "centerContinuous", 7: "distributed",
}
_VERT_ALIGN = {
    0: "top", 1: "center", 2: "bottom", 3: "justify", 4: "distributed",
}
_BORDER_STYLE = {
    0: None, 1: "thin", 2: "medium", 3: "dashed", 4: "dotted",
    5: "thick", 6: "double", 7: "hair", 8: "mediumDashed",
    9: "dashDot", 10: "mediumDashDot", 11: "dashDotDot",
    12: "mediumDashDotDot", 13: "slantDashDot",
}
_UNDERLINE = {
    0: None, 1: "single", 2: "double", 0x21: "singleAccounting", 0x22: "doubleAccounting",
}


def _rgb(colour_map, index):
    if index is None or index == 0x7FFF:
        return None
    rgb = colour_map.get(index)
    if rgb is None:
        return None
    return f"FF{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def _make_font(wb_xls, cell):
    xf = wb_xls.xf_list[cell.xf_index]
    f = wb_xls.font_list[xf.font_index]
    colour = _rgb(wb_xls.colour_map, f.colour_index)
    return Font(
        name=f.name, size=f.height / 20,
        bold=bool(f.bold), italic=bool(f.italic),
        underline=_UNDERLINE.get(f.underline_type),
        strike=bool(f.struck_out),
        color=Color(rgb=colour) if colour else None,
    )


def _make_alignment(wb_xls, cell):
    xf = wb_xls.xf_list[cell.xf_index]
    a = xf.alignment
    return Alignment(
        horizontal=_HOR_ALIGN.get(a.hor_align, "general"),
        vertical=_VERT_ALIGN.get(a.vert_align, "bottom"),
        wrap_text=bool(a.text_wrapped),
    )


def _make_fill(wb_xls, cell):
    xf = wb_xls.xf_list[cell.xf_index]
    bg = xf.background
    if bg.fill_pattern == 0:
        return None
    colour = _rgb(wb_xls.colour_map, bg.pattern_colour_index)
    if not colour:
        return None
    return PatternFill(fill_type="solid", fgColor=colour)


def _make_side(colour_map, line_type, colour_index):
    style = _BORDER_STYLE.get(line_type)
    if not style:
        return Side()
    colour = _rgb(colour_map, colour_index)
    return Side(style=style, color=Color(rgb=colour) if colour else None)


def _make_border(wb_xls, cell):
    xf = wb_xls.xf_list[cell.xf_index]
    b = xf.border
    cm = wb_xls.colour_map
    return Border(
        left=_make_side(cm, b.left_line_style,   b.left_colour_index),
        right=_make_side(cm, b.right_line_style,  b.right_colour_index),
        top=_make_side(cm, b.top_line_style,    b.top_colour_index),
        bottom=_make_side(cm, b.bottom_line_style, b.bottom_colour_index),
    )


def _num_format(wb_xls, cell):
    xf = wb_xls.xf_list[cell.xf_index]
    fmt = wb_xls.format_map.get(xf.format_key)
    if fmt and fmt.format_str and fmt.format_str != "General":
        return fmt.format_str
    return None


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
                letter = openpyxl.utils.get_column_letter(col_idx + 1)
                ws_xlsx.column_dimensions[letter].width = col_info.width / 256

        for row_idx, row_info in ws_xls.rowinfo_map.items():
            if row_info.height > 0:
                ws_xlsx.row_dimensions[row_idx + 1].height = row_info.height / 20

        for rlo, rhi, clo, chi in ws_xls.merged_cells:
            ws_xlsx.merge_cells(
                start_row=rlo + 1, end_row=rhi,
                start_column=clo + 1, end_column=chi,
            )

        merged_skip = set()
        for rlo, rhi, clo, chi in ws_xls.merged_cells:
            for r in range(rlo, rhi):
                for c in range(clo, chi):
                    if r != rlo or c != clo:
                        merged_skip.add((r, c))

        for row_idx in range(ws_xls.nrows):
            for col_idx in range(ws_xls.ncols):
                if (row_idx, col_idx) in merged_skip:
                    continue
                cell = ws_xls.cell(row_idx, col_idx)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(value, wb_xls.datemode)
                    except Exception:
                        pass

                out = ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1, value=value)
                out.font = _make_font(wb_xls, cell)
                out.alignment = _make_alignment(wb_xls, cell)
                out.border = _make_border(wb_xls, cell)
                fill = _make_fill(wb_xls, cell)
                if fill:
                    out.fill = fill
                fmt = _num_format(wb_xls, cell)
                if fmt:
                    out.number_format = fmt

    wb_xlsx.save(dst)
    return dst


def verify(src: str, dst: str) -> list[str]:
    diffs = []
    wb_xls = xlrd.open_workbook(src, formatting_info=True)
    wb_xlsx = openpyxl.load_workbook(dst)

    if wb_xls.sheet_names() != wb_xlsx.sheetnames:
        diffs.append(f"シート構成: {wb_xls.sheet_names()} → {wb_xlsx.sheetnames}")
        return diffs

    for sheet_name in wb_xls.sheet_names():
        ws_xls = wb_xls.sheet_by_name(sheet_name)
        ws_xlsx = wb_xlsx[sheet_name]

        xls_merges = {(rlo, rhi, clo, chi) for rlo, rhi, clo, chi in ws_xls.merged_cells}
        xlsx_merges = {
            (r.min_row - 1, r.max_row, r.min_col - 1, r.max_col)
            for r in ws_xlsx.merged_cells.ranges
        }
        if xls_merges != xlsx_merges:
            diffs.append(f"シート「{sheet_name}」結合セルが異なります")

        merged_skip_v = set()
        for rlo, rhi, clo, chi in ws_xls.merged_cells:
            for r in range(rlo, rhi):
                for c in range(clo, chi):
                    if r != rlo or c != clo:
                        merged_skip_v.add((r, c))

        for row_idx in range(ws_xls.nrows):
            for col_idx in range(ws_xls.ncols):
                if (row_idx, col_idx) in merged_skip_v:
                    continue
                loc = f"シート「{sheet_name}」行{row_idx+1} 列{col_idx+1}"
                cell = ws_xls.cell(row_idx, col_idx)
                out = ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1)

                xls_val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        xls_val = xlrd.xldate_as_datetime(xls_val, wb_xls.datemode)
                    except Exception:
                        pass
                xls_norm = None if xls_val == "" else xls_val
                xlsx_norm = None if out.value == "" else out.value
                if xls_norm != xlsx_norm:
                    diffs.append(f"{loc} 値: {xls_val!r} → {out.value!r}")

                xf = wb_xls.xf_list[cell.xf_index]
                f = wb_xls.font_list[xf.font_index]
                of = out.font

                if abs(f.height / 20 - (of.size or 11.0)) >= 0.5:
                    diffs.append(f"{loc} フォントサイズ: {f.height/20}pt → {of.size}pt")
                if bool(f.bold) != bool(of.bold):
                    diffs.append(f"{loc} 太字: {f.bold} → {of.bold}")
                if bool(f.italic) != bool(of.italic):
                    diffs.append(f"{loc} 斜体: {f.italic} → {of.italic}")
                if f.name and of.name and f.name != of.name:
                    diffs.append(f"{loc} フォント名: {f.name!r} → {of.name!r}")
                if _UNDERLINE.get(f.underline_type) != of.underline:
                    diffs.append(f"{loc} 下線: {_UNDERLINE.get(f.underline_type)!r} → {of.underline!r}")
                if bool(f.struck_out) != bool(of.strike):
                    diffs.append(f"{loc} 取り消し線: {f.struck_out} → {of.strike}")

                a = xf.alignment
                oa = out.alignment
                if _HOR_ALIGN.get(a.hor_align) != oa.horizontal:
                    diffs.append(f"{loc} 水平位置: {_HOR_ALIGN.get(a.hor_align)!r} → {oa.horizontal!r}")
                if bool(a.text_wrapped) != bool(oa.wrap_text):
                    diffs.append(f"{loc} 折り返し: {a.text_wrapped} → {oa.wrap_text}")

                xls_fmt = _num_format(wb_xls, cell) or "General"
                xlsx_fmt = out.number_format or "General"
                if xls_fmt != xlsx_fmt:
                    diffs.append(f"{loc} 数値書式: {xls_fmt!r} → {xlsx_fmt!r}")

        for col_idx, col_info in ws_xls.colinfo_map.items():
            if col_info.width == 0:
                continue
            expected = round(col_info.width / 256, 2)
            letter = openpyxl.utils.get_column_letter(col_idx + 1)
            actual = round(ws_xlsx.column_dimensions[letter].width, 2)
            if abs(expected - actual) >= 0.1:
                diffs.append(f"シート「{sheet_name}」列{letter} 列幅: {expected} → {actual}")

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
