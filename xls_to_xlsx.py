#!/usr/bin/env python3
"""XLS → XLSX 変換ツール"""

import sys
import os
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.styles.colors import Color

# --- 定数マッピング ---

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


# --- ヘルパー関数 ---

def _rgb(colour_map: dict, index: int) -> str | None:
    """xlrd カラーインデックス → openpyxl ARGB 16進文字列。変換不能なら None。"""
    if index is None or index == 0x7FFF:
        return None
    rgb = colour_map.get(index)
    if rgb is None:
        return None
    return f"FF{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def _make_font(wb_xls, cell) -> Font:
    xf = wb_xls.xf_list[cell.xf_index]
    f = wb_xls.font_list[xf.font_index]
    colour = _rgb(wb_xls.colour_map, f.colour_index)
    return Font(
        name=f.name,
        size=f.height / 20,
        bold=bool(f.bold),
        italic=bool(f.italic),
        underline=_UNDERLINE.get(f.underline_type),
        strike=bool(f.struck_out),
        color=Color(rgb=colour) if colour else None,
    )


def _make_alignment(wb_xls, cell) -> Alignment:
    xf = wb_xls.xf_list[cell.xf_index]
    a = xf.alignment
    return Alignment(
        horizontal=_HOR_ALIGN.get(a.hor_align, "general"),
        vertical=_VERT_ALIGN.get(a.vert_align, "bottom"),
        wrap_text=bool(a.text_wrapped),
    )


def _make_fill(wb_xls, cell) -> PatternFill | None:
    xf = wb_xls.xf_list[cell.xf_index]
    bg = xf.background
    if bg.fill_pattern == 0:
        return None
    colour = _rgb(wb_xls.colour_map, bg.pattern_colour_index)
    if not colour:
        return None
    return PatternFill(fill_type="solid", fgColor=colour)


def _make_side(colour_map: dict, line_type: int, colour_index: int) -> Side:
    style = _BORDER_STYLE.get(line_type)
    if not style:
        return Side()
    colour = _rgb(colour_map, colour_index)
    return Side(style=style, color=Color(rgb=colour) if colour else None)


def _make_border(wb_xls, cell) -> Border:
    xf = wb_xls.xf_list[cell.xf_index]
    b = xf.border
    cm = wb_xls.colour_map
    return Border(
        left=_make_side(cm, b.left_line_style,   b.left_colour_index),
        right=_make_side(cm, b.right_line_style,  b.right_colour_index),
        top=_make_side(cm, b.top_line_style,    b.top_colour_index),
        bottom=_make_side(cm, b.bottom_line_style, b.bottom_colour_index),
    )


def _num_format(wb_xls, cell) -> str | None:
    xf = wb_xls.xf_list[cell.xf_index]
    fmt = wb_xls.format_map.get(xf.format_key)
    if fmt and fmt.format_str and fmt.format_str != "General":
        return fmt.format_str
    return None


# --- 変換 ---

def convert(src: str, dst: str | None = None) -> str:
    if not src.lower().endswith(".xls"):
        raise ValueError(f"入力ファイルは .xls である必要があります: {src}")

    if dst is None:
        dst = os.path.splitext(src)[0] + ".xlsx"

    wb_xls = xlrd.open_workbook(src, formatting_info=True)
    wb_xlsx = openpyxl.Workbook()
    wb_xlsx.remove(wb_xlsx.active)

    for sheet_name in wb_xls.sheet_names():
        ws_xls = wb_xls.sheet_by_name(sheet_name)
        ws_xlsx = wb_xlsx.create_sheet(title=sheet_name)

        # 列幅（XLS: 1/256文字幅 → XLSX: 文字幅）
        for col_idx, col_info in ws_xls.colinfo_map.items():
            if col_info.width > 0:
                letter = openpyxl.utils.get_column_letter(col_idx + 1)
                ws_xlsx.column_dimensions[letter].width = col_info.width / 256

        # 行高（XLS: 1/20pt → XLSX: pt）
        for row_idx, row_info in ws_xls.rowinfo_map.items():
            if row_info.height > 0:
                ws_xlsx.row_dimensions[row_idx + 1].height = row_info.height / 20

        # 結合セル
        for rlo, rhi, clo, chi in ws_xls.merged_cells:
            ws_xlsx.merge_cells(
                start_row=rlo + 1, end_row=rhi,
                start_column=clo + 1, end_column=chi,
            )

        # 結合セルの非左上セル集合（書き込みスキップ用）
        merged_skip = set()
        for rlo, rhi, clo, chi in ws_xls.merged_cells:
            for r in range(rlo, rhi):
                for c in range(clo, chi):
                    if r != rlo or c != clo:
                        merged_skip.add((r, c))

        # セルデータ＋書式
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


# --- 検証 ---

def verify(src: str, dst: str) -> list[str]:
    """変換前後のファイルを比較し、差異のリストを返す。空リストなら一致。"""
    diffs = []

    wb_xls = xlrd.open_workbook(src, formatting_info=True)
    wb_xlsx = openpyxl.load_workbook(dst)

    if wb_xls.sheet_names() != wb_xlsx.sheetnames:
        diffs.append(f"シート構成: {wb_xls.sheet_names()} → {wb_xlsx.sheetnames}")
        return diffs

    for sheet_name in wb_xls.sheet_names():
        ws_xls = wb_xls.sheet_by_name(sheet_name)
        ws_xlsx = wb_xlsx[sheet_name]

        # 結合セルの検証
        xls_merges = {(rlo, rhi, clo, chi) for rlo, rhi, clo, chi in ws_xls.merged_cells}
        xlsx_merges = set()
        for rng in ws_xlsx.merged_cells.ranges:
            xlsx_merges.add((rng.min_row - 1, rng.max_row, rng.min_col - 1, rng.max_col))
        if xls_merges != xlsx_merges:
            diffs.append(f"シート「{sheet_name}」結合セルが異なります")

        # 結合セルの非左上セル（値・書式の検証をスキップ）
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
                loc = f"シート「{sheet_name}」行{row_idx+1} 列{col_idx+1}"
                cell = ws_xls.cell(row_idx, col_idx)
                out = ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1)

                # 値
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

                # フォント
                xf = wb_xls.xf_list[cell.xf_index]
                f = wb_xls.font_list[xf.font_index]
                of = out.font

                xls_size = f.height / 20
                xlsx_size = of.size or 11.0
                if abs(xls_size - xlsx_size) >= 0.5:
                    diffs.append(f"{loc} フォントサイズ: {xls_size}pt → {xlsx_size}pt")

                if bool(f.bold) != bool(of.bold):
                    diffs.append(f"{loc} 太字: {f.bold} → {of.bold}")
                if bool(f.italic) != bool(of.italic):
                    diffs.append(f"{loc} 斜体: {f.italic} → {of.italic}")
                if f.name and of.name and f.name != of.name:
                    diffs.append(f"{loc} フォント名: {f.name!r} → {of.name!r}")

                expected_ul = _UNDERLINE.get(f.underline_type)
                if expected_ul != of.underline:
                    diffs.append(f"{loc} 下線: {expected_ul!r} → {of.underline!r}")

                if bool(f.struck_out) != bool(of.strike):
                    diffs.append(f"{loc} 取り消し線: {f.struck_out} → {of.strike}")

                # アライメント
                a = xf.alignment
                oa = out.alignment
                if _HOR_ALIGN.get(a.hor_align) != oa.horizontal:
                    diffs.append(f"{loc} 水平位置: {_HOR_ALIGN.get(a.hor_align)!r} → {oa.horizontal!r}")
                if bool(a.text_wrapped) != bool(oa.wrap_text):
                    diffs.append(f"{loc} 折り返し: {a.text_wrapped} → {oa.wrap_text}")

                # 数値フォーマット
                xls_fmt = _num_format(wb_xls, cell) or "General"
                xlsx_fmt = out.number_format or "General"
                if xls_fmt != xlsx_fmt:
                    diffs.append(f"{loc} 数値書式: {xls_fmt!r} → {xlsx_fmt!r}")

        # 列幅の検証
        for col_idx, col_info in ws_xls.colinfo_map.items():
            if col_info.width == 0:
                continue
            expected = round(col_info.width / 256, 2)
            letter = openpyxl.utils.get_column_letter(col_idx + 1)
            actual = round(ws_xlsx.column_dimensions[letter].width, 2)
            if abs(expected - actual) >= 0.1:
                diffs.append(f"シート「{sheet_name}」列{letter} 列幅: {expected} → {actual}")

    return diffs


# --- CLI ---

def main():
    args = sys.argv[1:]
    if not args:
        print("使い方: python3 xls_to_xlsx.py <ファイル.xls> [出力先.xlsx]")
        print("例:     python3 xls_to_xlsx.py data.xls")
        print("        python3 xls_to_xlsx.py data.xls output.xlsx")
        sys.exit(1)

    src = args[0]
    dst = args[1] if len(args) >= 2 else None

    if not os.path.exists(src):
        print(f"エラー: ファイルが見つかりません: {src}")
        sys.exit(1)

    output = convert(src, dst)
    print(f"変換完了: {src} → {output}")

    print("差異チェック中...")
    diffs = verify(src, output)
    if diffs:
        print(f"[警告] {len(diffs)} 件の差異が見つかりました:")
        for d in diffs:
            print(f"  - {d}")
        sys.exit(2)
    else:
        print("差異なし: 変換前後のデータは完全に一致しています。")


if __name__ == "__main__":
    main()
