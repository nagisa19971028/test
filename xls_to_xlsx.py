#!/usr/bin/env python3
"""XLS → XLSX 変換ツール"""

import sys
import os
import xlrd
import openpyxl


def convert(src: str, dst: str | None = None) -> str:
    if not src.lower().endswith(".xls"):
        raise ValueError(f"入力ファイルは .xls である必要があります: {src}")

    if dst is None:
        dst = os.path.splitext(src)[0] + ".xlsx"

    wb_xls = xlrd.open_workbook(src, formatting_info=True)
    wb_xlsx = openpyxl.Workbook()
    wb_xlsx.remove(wb_xlsx.active)  # デフォルトシートを削除

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

                # 日付型 (xlrd type 3) を文字列に変換
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        dt = xlrd.xldate_as_datetime(value, wb_xls.datemode)
                        value = dt
                    except Exception:
                        pass

                ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1, value=value)

    wb_xlsx.save(dst)
    return dst


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


if __name__ == "__main__":
    main()
