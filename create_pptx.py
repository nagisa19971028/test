from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# Color palette
COLOR_DARK_BLUE = RGBColor(0x1F, 0x35, 0x64)   # Dark blue (title bg)
COLOR_MID_BLUE  = RGBColor(0x2E, 0x75, 0xB6)   # Mid blue (accent)
COLOR_LIGHT_BLUE= RGBColor(0xBD, 0xD7, 0xEE)   # Light blue (table header)
COLOR_ORANGE    = RGBColor(0xED, 0x7D, 0x31)   # Orange (warning/note)
COLOR_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_BLACK     = RGBColor(0x00, 0x00, 0x00)
COLOR_GRAY_BG   = RGBColor(0xF2, 0xF2, 0xF2)
COLOR_GREEN     = RGBColor(0x70, 0xAD, 0x47)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]  # completely blank


def add_rect(slide, l, t, w, h, fill=None, line=None):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, text, l, t, w, h,
                font_size=14, bold=False, color=COLOR_BLACK,
                align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txb


def slide_header(slide, title, subtitle=None):
    """Dark blue top bar with white title text."""
    add_rect(slide, 0, 0, 13.33, 1.2, fill=COLOR_DARK_BLUE)
    add_textbox(slide, title, 0.4, 0.1, 10, 0.8,
                font_size=28, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_textbox(slide, subtitle, 0.4, 0.82, 10, 0.4,
                    font_size=13, color=COLOR_LIGHT_BLUE, align=PP_ALIGN.LEFT)
    # bottom accent line
    add_rect(slide, 0, 1.2, 13.33, 0.04, fill=COLOR_MID_BLUE)


def slide_footer(slide, page, total):
    add_rect(slide, 0, 7.1, 13.33, 0.4, fill=COLOR_DARK_BLUE)
    add_textbox(slide, f"Excel拡張子問題 対応方針  |  {page} / {total}",
                0.3, 7.12, 12, 0.3, font_size=10, color=COLOR_WHITE)


# ── Slide 1: Title ────────────────────────────────────────────────────────────
sl1 = prs.slides.add_slide(BLANK)
add_rect(sl1, 0, 0, 13.33, 7.5, fill=COLOR_DARK_BLUE)
add_rect(sl1, 0, 2.8, 13.33, 0.06, fill=COLOR_MID_BLUE)
add_rect(sl1, 0, 4.8, 13.33, 0.06, fill=COLOR_MID_BLUE)

add_textbox(sl1, "Excel 拡張子問題", 1.0, 1.6, 11, 1.0,
            font_size=40, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
add_textbox(sl1, "対応方針のご説明", 1.0, 2.9, 11, 0.8,
            font_size=24, color=COLOR_LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_textbox(sl1, "xls → xlsx 変換対応について", 1.0, 3.8, 11, 0.6,
            font_size=16, color=COLOR_LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_textbox(sl1, "2026年5月20日", 1.0, 6.5, 11, 0.5,
            font_size=13, color=COLOR_LIGHT_BLUE, align=PP_ALIGN.CENTER)


# ── Slide 2: Agenda ───────────────────────────────────────────────────────────
sl2 = prs.slides.add_slide(BLANK)
slide_header(sl2, "本日のご説明内容")
slide_footer(sl2, 1, 5)

items = [
    ("01", "現状と課題",       "一部帳票がxls形式で出力される背景"),
    ("02", "対応策 A　RPA活用",  "RPA による自動変換（2つの実現方式）"),
    ("03", "対応策 B　帳票仕様見直し", "テンプレート変更による xlsx 対応"),
    ("04", "対応策の比較・推奨方針", "メリット・デメリット・推奨フロー"),
    ("05", "次のアクション",    "お客様への確認事項"),
]

for i, (num, title, desc) in enumerate(items):
    y = 1.5 + i * 1.0
    add_rect(sl2, 0.5, y, 1.0, 0.75, fill=COLOR_MID_BLUE)
    add_textbox(sl2, num, 0.5, y + 0.1, 1.0, 0.55,
                font_size=18, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_rect(sl2, 1.6, y, 10.8, 0.75, fill=COLOR_GRAY_BG, line=COLOR_LIGHT_BLUE)
    add_textbox(sl2, title, 1.8, y + 0.02, 4.5, 0.4,
                font_size=15, bold=True, color=COLOR_DARK_BLUE)
    add_textbox(sl2, desc, 1.8, y + 0.38, 10, 0.35,
                font_size=12, color=RGBColor(0x40, 0x40, 0x40))


# ── Slide 3: 現状と課題 ────────────────────────────────────────────────────────
sl3 = prs.slides.add_slide(BLANK)
slide_header(sl3, "現状と課題", "なぜ xls のまま出力されているのか")
slide_footer(sl3, 2, 5)

# Problem box
add_rect(sl3, 0.4, 1.4, 12.5, 1.5, fill=RGBColor(0xFF, 0xEB, 0xEB), line=RGBColor(0xC0, 0x00, 0x00))
add_textbox(sl3, "⚠  課題", 0.6, 1.45, 3, 0.4,
            font_size=13, bold=True, color=RGBColor(0xC0, 0x00, 0x00))
add_textbox(sl3,
    "一部帳票が .xls 形式で出力されており、.xlsx での出力が望ましい状況です。",
    0.6, 1.82, 12, 0.45, font_size=13, color=COLOR_BLACK)

# Cause box
add_rect(sl3, 0.4, 3.05, 12.5, 1.7, fill=COLOR_GRAY_BG, line=COLOR_LIGHT_BLUE)
add_textbox(sl3, "原因", 0.6, 3.1, 3, 0.4,
            font_size=13, bold=True, color=COLOR_DARK_BLUE)

causes = [
    "帳票の項目数が非常に多く、xlsx テンプレートへの直接変換が困難",
    "現状の仕様・構成のまま xls → xlsx へ変換することは難しい状況",
]
for j, c in enumerate(causes):
    add_textbox(sl3, f"●  {c}", 0.7, 3.5 + j * 0.5, 11.8, 0.45, font_size=13)

# Arrow
add_textbox(sl3, "▼  対応策を検討します", 4.5, 4.9, 6, 0.5,
            font_size=14, bold=True, color=COLOR_MID_BLUE, align=PP_ALIGN.CENTER)


# ── Slide 4: 対応策 A（RPA）────────────────────────────────────────────────────
sl4 = prs.slides.add_slide(BLANK)
slide_header(sl4, "対応策 A　RPA による自動変換", "担当者への要件説明が前提となります")
slide_footer(sl4, 3, 5)

add_textbox(sl4, "RPA を利用して xls → xlsx の自動変換を実現します。いずれの方式も、お客様から担当者への要件説明が必要です。",
            0.4, 1.35, 12.5, 0.5, font_size=12, color=RGBColor(0x40, 0x40, 0x40))

# A-1
add_rect(sl4, 0.4, 1.95, 5.9, 3.5, fill=COLOR_GRAY_BG, line=COLOR_MID_BLUE)
add_rect(sl4, 0.4, 1.95, 5.9, 0.5, fill=COLOR_MID_BLUE)
add_textbox(sl4, "A-1　担当者が作成", 0.5, 1.98, 5.5, 0.45,
            font_size=15, bold=True, color=COLOR_WHITE)

a1 = [
    ("メリット", "お客様側の工数が不要"),
    ("デメリット", "順番待ちがあり、着手・完了時期が未定"),
    ("リスク", "いつ実現できるか見通しが立てにくい"),
]
for k, (lbl, txt) in enumerate(a1):
    y = 2.58 + k * 0.85
    add_rect(sl4, 0.5, y, 1.4, 0.4, fill=COLOR_DARK_BLUE)
    add_textbox(sl4, lbl, 0.5, y + 0.05, 1.4, 0.35,
                font_size=11, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(sl4, txt, 2.0, y + 0.05, 4.1, 0.7, font_size=12)

# A-2
add_rect(sl4, 7.0, 1.95, 5.9, 3.5, fill=COLOR_GRAY_BG, line=COLOR_GREEN)
add_rect(sl4, 7.0, 1.95, 5.9, 0.5, fill=COLOR_GREEN)
add_textbox(sl4, "A-2　お客様が作成", 7.1, 1.98, 5.5, 0.45,
            font_size=15, bold=True, color=COLOR_WHITE)

a2 = [
    ("メリット", "調整次第ですぐ着手可能"),
    ("デメリット", "お客様側の工数が必要"),
    ("確認", "対応リソースの確保が前提条件"),
]
for k, (lbl, txt) in enumerate(a2):
    y = 2.58 + k * 0.85
    add_rect(sl4, 7.1, y, 1.4, 0.4, fill=COLOR_DARK_BLUE)
    add_textbox(sl4, lbl, 7.1, y + 0.05, 1.4, 0.35,
                font_size=11, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(sl4, txt, 8.6, y + 0.05, 4.1, 0.7, font_size=12)

add_textbox(sl4, "★ 推奨：A-2 を優先しつつ、A-1 への申請も並行して進めることで早期実現の可能性を高められます",
            0.4, 5.6, 12.5, 0.55, font_size=12, bold=True, color=COLOR_ORANGE)


# ── Slide 5: 対応策 B（帳票仕様見直し）─────────────────────────────────────────
sl5 = prs.slides.add_slide(BLANK)
slide_header(sl5, "対応策 B　帳票仕様の見直し", "テンプレート変更による xlsx 対応")
slide_footer(sl5, 4, 5)

add_textbox(sl5, "帳票テンプレート自体を変更し、xlsx に対応させる方法です。",
            0.4, 1.35, 12.5, 0.4, font_size=12, color=RGBColor(0x40, 0x40, 0x40))

methods = [
    ("B-1", "帳票を分割", "項目数を減らし、複数シート・複数ファイルに分割する"),
    ("B-2", "CSV等に変更", "帳票形式をCSV等に変更する（見た目・レイアウトが変わります）"),
]

for m, (code, title, desc) in enumerate(methods):
    y = 1.85 + m * 1.1
    add_rect(sl5, 0.4, y, 12.5, 0.95, fill=COLOR_GRAY_BG, line=COLOR_LIGHT_BLUE)
    add_rect(sl5, 0.4, y, 0.8, 0.95, fill=COLOR_MID_BLUE)
    add_textbox(sl5, code, 0.4, y + 0.2, 0.8, 0.55,
                font_size=12, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(sl5, title, 1.35, y + 0.08, 3.5, 0.4,
                font_size=14, bold=True, color=COLOR_DARK_BLUE)
    add_textbox(sl5, desc, 1.35, y + 0.48, 11, 0.4, font_size=12)

# Notes
add_rect(sl5, 0.4, 4.15, 12.5, 2.05, fill=RGBColor(0xFF, 0xF4, 0xE0), line=COLOR_ORANGE)
add_textbox(sl5, "⚠  注意事項", 0.6, 4.2, 4, 0.4,
            font_size=13, bold=True, color=COLOR_ORANGE)
notes = [
    "エンドユーザーが現行帳票のレイアウトにこだわりがある場合、調整・合意が必要",
    "帳票の在り方（様式・項目）に変更が生じるため、関係者との合意形成が前提",
    "修正工数については別途見積もりをご提示します（現時点では未確定）",
]
for n, note in enumerate(notes):
    add_textbox(sl5, f"●  {note}", 0.6, 4.65 + n * 0.46, 12.0, 0.44, font_size=12)


# ── Slide 6: 比較＆次のアクション ──────────────────────────────────────────────
sl6 = prs.slides.add_slide(BLANK)
slide_header(sl6, "次のアクション　／　確認事項", "お客様にご確認いただきたい事項")
slide_footer(sl6, 5, 5)

checks = [
    ("RPA 対応リソース確認",
     "A-2（お客様側RPA作成）について、対応可能なリソース・担当者はいますか？",
     COLOR_GREEN),
    ("帳票レイアウト変更の可否",
     "B案採用の場合、エンドユーザーとのレイアウト変更に関する調整は可能ですか？",
     COLOR_MID_BLUE),
    ("見積もり着手の承認",
     "B案を検討する場合、修正工数の見積もり作業に着手してよいですか？",
     COLOR_ORANGE),
    ("A-1 申請の並行実施",
     "A-2 を進めながら、担当者への A-1 申請も並行して行ってよいですか？",
     COLOR_MID_BLUE),
]

for i, (title, body, color) in enumerate(checks):
    y = 1.45 + i * 1.3 * 0.95
    add_rect(sl6, 0.4, y, 12.5, 1.05, fill=COLOR_GRAY_BG, line=color)
    add_rect(sl6, 0.4, y, 0.55, 1.05, fill=color)
    add_textbox(sl6, str(i + 1), 0.4, y + 0.28, 0.55, 0.5,
                font_size=16, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(sl6, f"□  {title}", 1.08, y + 0.05, 11.5, 0.42,
                font_size=13, bold=True, color=COLOR_DARK_BLUE)
    add_textbox(sl6, body, 1.08, y + 0.5, 11.5, 0.5, font_size=12)


out = "/home/user/test/Excel拡張子問題_対応方針.pptx"
prs.save(out)
print(f"Saved: {out}")
