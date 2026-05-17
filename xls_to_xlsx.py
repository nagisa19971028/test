#!/usr/bin/env python3
"""XLS → XLSX 変換ツール"""

import sys
import os
import xlrd
import openpyxl
from openpyxl.styles import Font


def _xls_font(wb_xls, cell) -> tuple:
    """xlrdのセルからフォント情報 (size_pt, bold, italic) を返す。"""
    xf = wb_xls.xf_list[cell.xf_index]
    font = wb_xls.font_list[xf.font_index]
    return (font.height / 20, bool(font.bold), bool(font.italic))


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

        # 列幅をコピー（XLS単位: 1/256文字幅 → XLSX単位: 文字幅）
        for col_idx, col_info in ws_xls.colinfo_map.items():
            if col_info.width > 0:
                col_letter = openpyxl.utils.get_column_letter(col_idx + 1)
                ws_xlsx.column_dimensions[col_letter].width = col_info.width / 256

        # 行高をコピー（XLS単位: 1/20ポイント → XLSX単位: ポイント）
        for row_idx, row_info in ws_xls.rowinfo_map.items():
            if row_info.height > 0:
                ws_xlsx.row_dimensions[row_idx + 1].height = row_info.height / 20

        for row_idx in range(ws_xls.nrows):
            for col_idx in range(ws_xls.ncols):
                cell = ws_xls.cell(row_idx, col_idx)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        dt = xlrd.xldate_as_datetime(value, wb_xls.datemode)
                        value = dt
                    except Exception:
                        pass

                out_cell = ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1, value=value)

                # フォント（サイズ・太字・斜体）をコピー
                size, bold, italic = _xls_font(wb_xls, cell)
                out_cell.font = Font(size=size, bold=bold, italic=italic)

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
        diffs.append(f"シート構成が異なります: {xls_sheets} → {xlsx_sheets}")
        return diffs

    for sheet_name in xls_sheets:
        ws_xls = wb_xls.sheet_by_name(sheet_name)
        ws_xlsx = wb_xlsx[sheet_name]

        for row_idx in range(ws_xls.nrows):
            for col_idx in range(ws_xls.ncols):
                loc = f"シート「{sheet_name}」行{row_idx+1} 列{col_idx+1}"
                cell = ws_xls.cell(row_idx, col_idx)
                out_cell = ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1)

                # 値の比較
                xls_val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        xls_val = xlrd.xldate_as_datetime(xls_val, wb_xls.datemode)
                    except Exception:
                        pass
                xlsx_val = out_cell.value
                xls_norm = None if xls_val == "" else xls_val
                xlsx_norm = None if xlsx_val == "" else xlsx_val
                if xls_norm != xlsx_norm:
                    diffs.append(f"{loc} 値: 変換前={xls_val!r} / 変換後={xlsx_val!r}")

                # フォントの比較
                xls_size, xls_bold, xls_italic = _xls_font(wb_xls, cell)
                xlsx_font = out_cell.font
                xlsx_size = xlsx_font.size or 11.0
                xlsx_bold = bool(xlsx_font.bold)
                xlsx_italic = bool(xlsx_font.italic)

                if abs(xls_size - xlsx_size) >= 0.5:
                    diffs.append(f"{loc} フォントサイズ: 変換前={xls_size}pt / 変換後={xlsx_size}pt")
                if xls_bold != xlsx_bold:
                    diffs.append(f"{loc} 太字: 変換前={xls_bold} / 変換後={xlsx_bold}")
                if xls_italic != xlsx_italic:
                    diffs.append(f"{loc} 斜体: 変換前={xls_italic} / 変換後={xlsx_italic}")

        # 列幅の検証
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
