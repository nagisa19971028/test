#!/usr/bin/env python3
"""テスト用の複雑な帳票XLSファイルを生成する"""

import xlwt
from datetime import date

wb = xlwt.Workbook(encoding="utf-8")

# ===== シート1: 請求書 =====
ws1 = wb.add_sheet("請求書")

# タイトル
ws1.write(0, 0, "請求書")
ws1.write(1, 0, "請求番号")
ws1.write(1, 1, "INV-2024-00123")
ws1.write(2, 0, "発行日")
ws1.write(2, 1, "2024/03/31")
ws1.write(3, 0, "支払期限")
ws1.write(3, 1, "2024/04/30")

# 請求先
ws1.write(5, 0, "請求先")
ws1.write(6, 0, "会社名")
ws1.write(6, 1, "株式会社サンプル商事")
ws1.write(7, 0, "担当者")
ws1.write(7, 1, "山田 太郎 様")
ws1.write(8, 0, "住所")
ws1.write(8, 1, "〒100-0001 東京都千代田区千代田1-1-1")

# 請求元
ws1.write(5, 3, "請求元")
ws1.write(6, 3, "会社名")
ws1.write(6, 4, "株式会社テスト開発")
ws1.write(7, 3, "担当者")
ws1.write(7, 4, "鈴木 花子")
ws1.write(8, 3, "電話")
ws1.write(8, 4, "03-1234-5678")
ws1.write(9, 3, "メール")
ws1.write(9, 4, "suzuki@test-dev.co.jp")

# 明細ヘッダー
ws1.write(11, 0, "No.")
ws1.write(11, 1, "品目")
ws1.write(11, 2, "数量")
ws1.write(11, 3, "単位")
ws1.write(11, 4, "単価（円）")
ws1.write(11, 5, "金額（円）")

# 明細データ
items = [
    (1, "Webシステム開発（基本設計）", 1, "式", 500000, 500000),
    (2, "Webシステム開発（詳細設計）", 1, "式", 400000, 400000),
    (3, "Webシステム開発（実装）",     2, "人月", 600000, 1200000),
    (4, "Webシステム開発（テスト）",   1, "人月", 600000, 600000),
    (5, "サーバー構築・設定",          1, "式", 150000, 150000),
    (6, "ドキュメント作成",            1, "式",  80000,  80000),
    (7, "保守サポート（3ヶ月分）",     3, "ヶ月", 50000, 150000),
]

for i, (no, name, qty, unit, price, amount) in enumerate(items):
    row = 12 + i
    ws1.write(row, 0, no)
    ws1.write(row, 1, name)
    ws1.write(row, 2, qty)
    ws1.write(row, 3, unit)
    ws1.write(row, 4, price)
    ws1.write(row, 5, amount)

# 合計
subtotal = sum(x[5] for x in items)
tax = int(subtotal * 0.10)
total = subtotal + tax

ws1.write(20, 4, "小計")
ws1.write(20, 5, subtotal)
ws1.write(21, 4, "消費税（10%）")
ws1.write(21, 5, tax)
ws1.write(22, 4, "合計金額")
ws1.write(22, 5, total)

# 備考
ws1.write(24, 0, "備考")
ws1.write(25, 0, "・お振込先：〇〇銀行 △△支店 普通 1234567 カ）テストカイハツ")
ws1.write(26, 0, "・振込手数料はご負担ください。")
ws1.write(27, 0, "・本請求書に関するお問い合わせは担当者までご連絡ください。")

# ===== シート2: 月次売上明細 =====
ws2 = wb.add_sheet("月次売上明細")

ws2.write(0, 0, "月次売上明細レポート")
ws2.write(1, 0, "対象期間")
ws2.write(1, 1, "2024年1月〜3月")
ws2.write(2, 0, "作成日")
ws2.write(2, 1, "2024/03/31")

# ヘッダー
headers = ["月", "担当者", "顧客名", "案件名", "売上金額", "原価", "粗利", "粗利率(%)"]
for col, h in enumerate(headers):
    ws2.write(4, col, h)

# データ
records = [
    ("2024/01", "田中 一郎", "A社", "システム改修",    800000, 480000, 320000, 40.0),
    ("2024/01", "佐藤 二郎", "B社", "コンサルティング",350000, 140000, 210000, 60.0),
    ("2024/01", "鈴木 三郎", "C社", "保守運用",        200000, 100000, 100000, 50.0),
    ("2024/02", "田中 一郎", "D社", "新規開発",       1200000, 720000, 480000, 40.0),
    ("2024/02", "佐藤 二郎", "A社", "追加開発",        450000, 225000, 225000, 50.0),
    ("2024/02", "山本 四郎", "E社", "システム導入",    600000, 300000, 300000, 50.0),
    ("2024/03", "田中 一郎", "F社", "要件定義",        300000, 150000, 150000, 50.0),
    ("2024/03", "鈴木 三郎", "B社", "テスト支援",      250000, 125000, 125000, 50.0),
    ("2024/03", "山本 四郎", "G社", "インフラ構築",    500000, 250000, 250000, 50.0),
    ("2024/03", "佐藤 二郎", "H社", "運用設計",        380000, 190000, 190000, 50.0),
]

for i, row_data in enumerate(records):
    for col, val in enumerate(row_data):
        ws2.write(5 + i, col, val)

# 合計行
total_row = 5 + len(records)
ws2.write(total_row, 0, "合計")
ws2.write(total_row, 4, sum(r[4] for r in records))
ws2.write(total_row, 5, sum(r[5] for r in records))
ws2.write(total_row, 6, sum(r[6] for r in records))
avg_margin = sum(r[7] for r in records) / len(records)
ws2.write(total_row, 7, round(avg_margin, 1))

# 担当者別集計
ws2.write(total_row + 2, 0, "担当者別集計")
ws2.write(total_row + 3, 0, "担当者")
ws2.write(total_row + 3, 1, "売上合計")
ws2.write(total_row + 3, 2, "粗利合計")

persons = {}
for r in records:
    name = r[1]
    persons.setdefault(name, [0, 0])
    persons[name][0] += r[4]
    persons[name][1] += r[6]

for j, (name, (sales, profit)) in enumerate(sorted(persons.items())):
    ws2.write(total_row + 4 + j, 0, name)
    ws2.write(total_row + 4 + j, 1, sales)
    ws2.write(total_row + 4 + j, 2, profit)

# ===== シート3: 経費精算書 =====
ws3 = wb.add_sheet("経費精算書")

ws3.write(0, 0, "経費精算書")
ws3.write(1, 0, "申請者")
ws3.write(1, 1, "田中 一郎")
ws3.write(2, 0, "所属")
ws3.write(2, 1, "開発部 第1グループ")
ws3.write(3, 0, "申請日")
ws3.write(3, 1, "2024/03/31")
ws3.write(4, 0, "精算対象月")
ws3.write(4, 1, "2024年3月")

exp_headers = ["日付", "費目", "内容", "支払先", "金額（円）", "領収書"]
for col, h in enumerate(exp_headers):
    ws3.write(6, col, h)

expenses = [
    ("2024/03/05", "交通費",   "客先訪問（A社往復）",         "交通機関",     1640, "有"),
    ("2024/03/07", "交通費",   "客先訪問（B社往復）",         "交通機関",     2100, "有"),
    ("2024/03/12", "接待費",   "顧客との打合せランチ（3名）", "〇〇レストラン", 8500, "有"),
    ("2024/03/14", "通信費",   "携帯電話料金（業務分）",      "通信会社",     3000, "有"),
    ("2024/03/19", "交通費",   "セミナー参加（往復）",        "交通機関",     980, "有"),
    ("2024/03/19", "研修費",   "技術セミナー参加費",          "△△協会",     15000, "有"),
    ("2024/03/22", "消耗品費", "業務用文具・用紙",            "事務用品店",   2350, "有"),
    ("2024/03/28", "交通費",   "月末客先訪問（C社往復）",     "交通機関",     1820, "有"),
]

for i, row_data in enumerate(expenses):
    for col, val in enumerate(row_data):
        ws3.write(7 + i, col, val)

exp_total = sum(e[4] for e in expenses)
ws3.write(7 + len(expenses), 3, "合計")
ws3.write(7 + len(expenses), 4, exp_total)

ws3.write(7 + len(expenses) + 2, 0, "承認欄")
ws3.write(7 + len(expenses) + 3, 0, "担当")
ws3.write(7 + len(expenses) + 3, 1, "")
ws3.write(7 + len(expenses) + 3, 2, "課長")
ws3.write(7 + len(expenses) + 3, 3, "")
ws3.write(7 + len(expenses) + 3, 4, "部長")
ws3.write(7 + len(expenses) + 3, 5, "")

wb.save("sample_complex.xls")
print("sample_complex.xls を作成しました")
print(f"  シート1: 請求書（合計 {total:,}円）")
print(f"  シート2: 月次売上明細（{len(records)}件）")
print(f"  シート3: 経費精算書（合計 {exp_total:,}円）")
