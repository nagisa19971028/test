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

    wb_xls = xlrd.open_workbook(src)
    wb_xlsx = openpyxl.Workbook()
    wb_xlsx.remove(wb_xlsx.active)  # デフォルトシートを削除

    for sheet_name in wb_xls.sheet_names():
        ws_xls = wb_xls.sheet_by_name(sheet_name)
        ws_xlsx = wb_xlsx.create_sheet(title=sheet_name)

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
