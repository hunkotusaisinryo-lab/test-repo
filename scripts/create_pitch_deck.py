"""
肉酒場 然（ぜん）事業計画 ピッチデック生成スクリプト
python-pptx を使用
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import os
from datetime import datetime

# === カラーパレット ===
C_BG       = RGBColor(0x1A, 0x12, 0x0B)   # 深い焦げ茶（和の闇）
C_ACCENT   = RGBColor(0xC8, 0x96, 0x3C)   # 金色（麹・稲穂）
C_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT    = RGBColor(0xF0, 0xE6, 0xD3)   # 生成り（和紙）
C_GRAY     = RGBColor(0x8A, 0x7A, 0x6A)   # 灰茶
C_RED      = RGBColor(0xC0, 0x39, 0x2B)   # 朱色
C_GREEN    = RGBColor(0x27, 0x7A, 0x5A)   # 緑

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def new_prs():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_slide(prs):
    layout = prs.slide_layouts[6]  # completely blank
    return prs.slides.add_slide(layout)


def bg(slide, color=C_BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, text, left, top, width, height,
                size=18, bold=False, color=C_WHITE,
                align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_line(slide, left, top, width, color=C_ACCENT, height=Pt(1.5)):
    rect = slide.shapes.add_shape(1, left, top, width, int(height))
    rect.fill.solid()
    rect.fill.fore_color.rgb = color
    rect.line.fill.background()
    return rect


def slide_number(slide, num, total=12):
    add_textbox(slide, f"{num} / {total}",
                Inches(12.5), Inches(7.1), Inches(0.8), Inches(0.3),
                size=9, color=C_GRAY, align=PP_ALIGN.RIGHT)


# ─────────────────────────────────────────
# スライド定義
# ─────────────────────────────────────────

def s01_title(prs):
    """タイトル"""
    sl = blank_slide(prs)
    bg(sl)

    # 左帯
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    # ロゴタイプ
    add_textbox(sl, "肉酒場", Inches(1.0), Inches(1.0), Inches(6), Inches(1.0),
                size=36, bold=True, color=C_ACCENT)
    add_textbox(sl, "然（ぜん）", Inches(1.0), Inches(1.9), Inches(6), Inches(1.4),
                size=72, bold=True, color=C_WHITE)

    # キャッチコピー
    add_line(sl, Inches(1.0), Inches(3.5), Inches(5), color=C_ACCENT)
    add_textbox(sl, "糀漬け肉の定食と、和の一杯。",
                Inches(1.0), Inches(3.65), Inches(7), Inches(0.6),
                size=22, italic=True, color=C_LIGHT)

    # サブ
    add_textbox(sl, "事業計画書　2026年6月",
                Inches(1.0), Inches(4.5), Inches(5), Inches(0.5),
                size=14, color=C_GRAY)

    # 右側 大きな漢字装飾
    add_textbox(sl, "然", Inches(9.5), Inches(0.8), Inches(3.5), Inches(5.5),
                size=220, bold=True, color=RGBColor(0x2A, 0x1E, 0x12),
                align=PP_ALIGN.CENTER)

    slide_number(sl, 1)
    return sl


def s02_vision(prs):
    """ビジョン"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "VISION", Inches(0.6), Inches(0.4), Inches(4), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "私たちのビジョン", Inches(0.6), Inches(0.75), Inches(8), Inches(0.7),
                size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    # 大きなビジョン文
    add_textbox(sl, "7年で100店舗。",
                Inches(0.6), Inches(2.0), Inches(12), Inches(1.2),
                size=54, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)

    add_textbox(sl,
                "発酵と肉の掛け合わせで、\n日本の食文化に新しい定番をつくる。",
                Inches(0.6), Inches(3.4), Inches(12), Inches(1.2),
                size=24, color=C_LIGHT, align=PP_ALIGN.CENTER)

    add_textbox(sl,
                "居酒屋×発酵×定食　という空白地帯に参入し、\n"
                "糀漬け肉料理を「日本の新定番」として全国に広める。",
                Inches(1.5), Inches(5.0), Inches(10), Inches(1.2),
                size=16, color=C_GRAY, align=PP_ALIGN.CENTER, italic=True)

    slide_number(sl, 2)
    return sl


def s03_why_now(prs):
    """なぜ今か — 市場環境"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "MARKET", Inches(0.6), Inches(0.4), Inches(4), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "なぜ今か　――　市場の追い風", Inches(0.6), Inches(0.75),
                Inches(9), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    # 3ボックス
    boxes = [
        ("外食産業", "35.7兆円", "2025年市場規模\n（前年比＋3.5%）"),
        ("居酒屋業態", "48ヶ月連続", "売上前年比プラス\nコロナ前超え達成"),
        ("発酵食品", "SNS急拡大", "健康志向×体験価値\nZ世代・40代双方に刺さる"),
    ]
    for i, (title, big, sub) in enumerate(boxes):
        x = Inches(0.6 + i * 4.1)
        add_rect(sl, x, Inches(2.0), Inches(3.8), Inches(4.2),
                 RGBColor(0x2A, 0x1E, 0x12))
        add_textbox(sl, title, x + Inches(0.15), Inches(2.15),
                    Inches(3.5), Inches(0.5), size=14, color=C_ACCENT, bold=True)
        add_textbox(sl, big, x + Inches(0.1), Inches(2.7),
                    Inches(3.6), Inches(0.9), size=32, bold=True, color=C_WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(sl, sub, x + Inches(0.1), Inches(3.7),
                    Inches(3.6), Inches(0.8), size=13, color=C_GRAY,
                    align=PP_ALIGN.CENTER)

    add_textbox(sl,
                "→ 出店タイミングは今が最適。居酒屋×発酵×定食の空白地帯は誰もいない。",
                Inches(0.6), Inches(6.4), Inches(12), Inches(0.6),
                size=16, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)

    slide_number(sl, 3)
    return sl


def s04_concept(prs):
    """コンセプト"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "CONCEPT", Inches(0.6), Inches(0.4), Inches(4), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "肉酒場 然のコンセプト", Inches(0.6), Inches(0.75),
                Inches(9), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    # 左：コンセプト説明
    add_textbox(sl, "業態定義", Inches(0.6), Inches(1.8), Inches(5.5), Inches(0.4),
                size=12, bold=True, color=C_ACCENT)
    add_textbox(sl,
                "昼：糀漬け肉定食の専門店\n夜：発酵×肉の創作酒場",
                Inches(0.6), Inches(2.2), Inches(5.5), Inches(0.9),
                size=20, bold=True, color=C_WHITE)

    add_textbox(sl, "三つの独自性", Inches(0.6), Inches(3.3), Inches(5.5), Inches(0.4),
                size=12, bold=True, color=C_ACCENT)

    points = [
        ("糀漬け×低温調理", "前日仕込みで品質均一化。FC展開時の再現性を担保。"),
        ("昼夜二毛作モデル", "ランチ定食で固定売上 ＋ ディナー酒場でドリンク収益。"),
        ("石板焼き体験", "煙・音・ロゼ色の断面がSNS映え。来店動機を設計する。"),
    ]
    for i, (head, body) in enumerate(points):
        y = Inches(3.75 + i * 0.95)
        add_rect(sl, Inches(0.6), y, Inches(0.06), Inches(0.65), C_ACCENT)
        add_textbox(sl, head, Inches(0.8), y, Inches(2.0), Inches(0.35),
                    size=14, bold=True, color=C_WHITE)
        add_textbox(sl, body, Inches(0.8), y + Inches(0.35), Inches(5.0), Inches(0.4),
                    size=12, color=C_GRAY)

    # 右：ポジショニング図（テキストで表現）
    add_rect(sl, Inches(7.0), Inches(1.8), Inches(5.8), Inches(5.0),
             RGBColor(0x2A, 0x1E, 0x12))
    add_textbox(sl, "ポジショニングマップ",
                Inches(7.1), Inches(1.9), Inches(5.6), Inches(0.4),
                size=12, bold=True, color=C_ACCENT)

    map_text = (
        "       健康・発酵（高）\n"
        "            ↑\n"
        "   びゃく   │\n"
        "            │   ★肉酒場 然\n"
        "            │  （空白地帯）\n"
        "大衆──────┼──────── こだわり\n"
        " ひまり堂   │ なかよし\n"
        " つむぎ堂   │\n"
        "            ↓\n"
        "       健康・発酵（低）"
    )
    add_textbox(sl, map_text, Inches(7.1), Inches(2.4), Inches(5.6), Inches(3.8),
                size=11, color=C_LIGHT)

    add_textbox(sl, "「健康・発酵 × 肉 × ランチ＆ディナー」\nは競合ゼロの完全空白地帯",
                Inches(7.1), Inches(6.1), Inches(5.6), Inches(0.7),
                size=12, bold=True, color=C_ACCENT)

    slide_number(sl, 4)
    return sl


def s05_products(prs):
    """看板3商品"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "PRODUCT", Inches(0.6), Inches(0.4), Inches(4), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "看板3品　――　これを食べに来る", Inches(0.6), Inches(0.75),
                Inches(9), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    products = [
        {
            "no": "01",
            "name": "糀漬けサガリの\n石板焼き定食",
            "price": "¥1,980",
            "desc": "北海道産サガリを糀で48時間熟成。\n石板で焼き上げる煙と音が体験価値に。",
            "cost": "原価率 34.5%",
            "hook": "SNS映え No.1",
        },
        {
            "no": "02",
            "name": "糀漬け豚の\n厚切りカツ定食",
            "price": "¥1,780",
            "desc": "糀漬け豚ロースを低温調理後に揚げる。\nロゼ色の断面が口コミの起点。",
            "cost": "原価率 32.1%",
            "hook": "定食の王道×発酵",
        },
        {
            "no": "03",
            "name": "糀漬け鶏の\nロースト定食",
            "price": "¥1,680",
            "desc": "鶏もも肉を糀×ハーブで低温ロースト。\n鶏の旨みが最大限に引き出される。",
            "cost": "原価率 35.6%",
            "hook": "ヘルシー需要取込み",
        },
    ]

    for i, p in enumerate(products):
        x = Inches(0.5 + i * 4.2)
        # カード背景
        add_rect(sl, x, Inches(1.8), Inches(4.0), Inches(5.2),
                 RGBColor(0x2A, 0x1E, 0x12))
        # No
        add_textbox(sl, p["no"], x + Inches(0.15), Inches(1.95),
                    Inches(0.6), Inches(0.4), size=11, bold=True, color=C_ACCENT)
        # フック
        add_rect(sl, x + Inches(2.0), Inches(1.95), Inches(1.8), Inches(0.32),
                 C_ACCENT)
        add_textbox(sl, p["hook"], x + Inches(2.0), Inches(1.95),
                    Inches(1.8), Inches(0.32), size=10, bold=True,
                    color=C_BG, align=PP_ALIGN.CENTER)
        # 商品名
        add_textbox(sl, p["name"], x + Inches(0.15), Inches(2.4),
                    Inches(3.7), Inches(0.8), size=18, bold=True, color=C_WHITE)
        # 価格
        add_textbox(sl, p["price"], x + Inches(0.15), Inches(3.3),
                    Inches(3.7), Inches(0.55), size=28, bold=True, color=C_ACCENT)
        # 説明
        add_textbox(sl, p["desc"], x + Inches(0.15), Inches(3.95),
                    Inches(3.7), Inches(0.9), size=12, color=C_LIGHT)
        # 原価率
        add_textbox(sl, p["cost"], x + Inches(0.15), Inches(5.0),
                    Inches(3.7), Inches(0.35), size=11, color=C_GRAY)

    add_textbox(sl,
                "全3品にライス・味噌汁・小鉢2品・漬物が付く。「体にいい定食」を徹底追求。",
                Inches(0.5), Inches(7.0), Inches(12), Inches(0.35),
                size=12, color=C_GRAY, align=PP_ALIGN.CENTER, italic=True)

    slide_number(sl, 5)
    return sl


def s06_competitor(prs):
    """競合分析"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "COMPETITION", Inches(0.6), Inches(0.4), Inches(5), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "競合5店　繁盛の法則と然の優位性", Inches(0.6), Inches(0.75),
                Inches(10), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    # テーブルヘッダー
    headers = ["店舗", "月商", "ランチ", "発酵", "肉特化", "FC展開"]
    col_w   = [2.5, 1.5, 1.2, 1.2, 1.2, 1.2]
    col_x   = [0.6]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w)

    y_h = Inches(1.75)
    add_rect(sl, Inches(0.5), y_h, Inches(12.3), Inches(0.38), C_ACCENT)
    for i, h in enumerate(headers):
        add_textbox(sl, h, Inches(col_x[i] + 0.05), y_h + Inches(0.04),
                    Inches(col_w[i]), Inches(0.3),
                    size=11, bold=True, color=C_BG)

    rows = [
        ["なかよし（恵比寿）",  "非公開", "◎", "△", "✗", "◎"],
        ["つむぎ堂（新宿）",    "1,500万", "✗", "✗", "△", "◎"],
        ["びゃく（学芸大学）",  "非公開", "✗", "◎", "✗", "△"],
        ["ひまり堂（恵比寿）",  "850万", "✗", "✗", "△", "◎"],
        ["ひまり商店（新橋）",  "1,700万", "✗", "✗", "✗", "◎"],
        ["★ 肉酒場 然",         "720万目標", "◎", "◎", "◎", "◎"],
    ]

    for r, row in enumerate(rows):
        y_r = Inches(2.15 + r * 0.64)
        bg_c = RGBColor(0x2A, 0x1E, 0x12) if r < 5 else RGBColor(0x3A, 0x2A, 0x08)
        add_rect(sl, Inches(0.5), y_r, Inches(12.3), Inches(0.6), bg_c)
        for c, cell in enumerate(row):
            color = C_ACCENT if r == 5 else (C_WHITE if c == 0 else C_LIGHT)
            add_textbox(sl, cell,
                        Inches(col_x[c] + 0.05), y_r + Inches(0.12),
                        Inches(col_w[c]), Inches(0.38),
                        size=12, bold=(r == 5), color=color)

    add_textbox(sl,
                "「健康・発酵 × 肉 × ランチ×ディナー × FC」を同時に満たすのは然だけ",
                Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.35),
                size=13, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)

    slide_number(sl, 6)
    return sl


def s07_profile(prs):
    """代表者プロフィール"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "FOUNDER", Inches(0.6), Inches(0.4), Inches(4), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "代表者プロフィール", Inches(0.6), Inches(0.75),
                Inches(9), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    # 経歴
    add_textbox(sl, "経歴", Inches(0.6), Inches(1.75), Inches(5.5), Inches(0.4),
                size=13, bold=True, color=C_ACCENT)
    careers = [
        ("前職", "都内約25店舗展開　焼肉レストランブランド　運営本部長"),
        ("現 職", "やきとり家すみれ　運営本部長（SV体制構築・販促・業績管理）"),
        ("兼 業", "EC通販事業 経営メンバー（販売戦略・商品企画・数値分析）"),
    ]
    for i, (label, text) in enumerate(careers):
        y = Inches(2.2 + i * 0.65)
        add_rect(sl, Inches(0.6), y + Inches(0.05), Inches(0.06), Inches(0.42), C_ACCENT)
        add_textbox(sl, label, Inches(0.8), y, Inches(1.0), Inches(0.35),
                    size=11, color=C_ACCENT, bold=True)
        add_textbox(sl, text, Inches(1.9), y, Inches(4.2), Inches(0.55),
                    size=13, color=C_WHITE)

    # 5つの強み
    add_textbox(sl, "5つの強み", Inches(7.0), Inches(1.75), Inches(5.8), Inches(0.4),
                size=13, bold=True, color=C_ACCENT)
    strengths = [
        "複数店舗の運営経験（本部長として売上・利益・人材を総合管理）",
        "現場と経営の両立（現場課題と経営数値の両方を深く理解）",
        "数字に基づく改善力（客数・客単価・原価・人件費を分解改善）",
        "商品企画・販促経験（SNS・LINE・キャンペーンによる集客実績）",
        "人材育成と仕組み化（属人化しない標準化・教育体制の構築）",
    ]
    for i, s in enumerate(strengths):
        y = Inches(2.2 + i * 0.72)
        add_rect(sl, Inches(7.0), y + Inches(0.1), Inches(5.8), Inches(0.52),
                 RGBColor(0x2A, 0x1E, 0x12))
        add_textbox(sl, f"0{i+1}  {s}", Inches(7.15), y + Inches(0.1),
                    Inches(5.5), Inches(0.5), size=12, color=C_LIGHT)

    # ひとこと
    add_rect(sl, Inches(0.5), Inches(5.9), Inches(12.3), Inches(1.0),
             RGBColor(0x2A, 0x1E, 0x12))
    add_textbox(sl,
                "「飲食の現場と数字を両方知っている人間が、\n"
                "　ゼロから作るブランドだから強い。」",
                Inches(0.8), Inches(5.95), Inches(11.7), Inches(0.9),
                size=16, italic=True, color=C_ACCENT, align=PP_ALIGN.CENTER)

    slide_number(sl, 7)
    return sl


def s08_pl(prs):
    """収支計画"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "FINANCIALS", Inches(0.6), Inches(0.4), Inches(5), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "収支計画　――　堅実な数字設計", Inches(0.6), Inches(0.75),
                Inches(10), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    # 月次売上目標
    add_textbox(sl, "月次売上目標（安定期・30席）",
                Inches(0.6), Inches(1.7), Inches(5.5), Inches(0.4),
                size=13, bold=True, color=C_ACCENT)

    sales_rows = [
        ("ランチ", "2,000円", "30席", "1.8回転", "25日", "270万円"),
        ("ディナー", "5,000円", "30席", "1.2回転", "25日", "450万円"),
        ("合計", "", "", "", "", "720万円"),
    ]
    s_headers = ["", "客単価", "席数", "回転", "営業日", "売上"]
    s_col_w = [1.1, 1.0, 0.8, 0.9, 0.8, 1.0]
    s_col_x = [0.6]
    for w in s_col_w[:-1]:
        s_col_x.append(s_col_x[-1] + w)

    add_rect(sl, Inches(0.5), Inches(2.1), Inches(5.8), Inches(0.32), C_ACCENT)
    for i, h in enumerate(s_headers):
        add_textbox(sl, h, Inches(s_col_x[i]+0.03), Inches(2.12),
                    Inches(s_col_w[i]), Inches(0.28), size=10, bold=True, color=C_BG)

    for r, row in enumerate(sales_rows):
        y = Inches(2.44 + r * 0.48)
        bg_c = RGBColor(0x3A, 0x2A, 0x08) if r == 2 else RGBColor(0x2A, 0x1E, 0x12)
        add_rect(sl, Inches(0.5), y, Inches(5.8), Inches(0.46), bg_c)
        for c, cell in enumerate(row):
            col = C_ACCENT if (r == 2 and c == 5) else C_WHITE
            add_textbox(sl, cell, Inches(s_col_x[c]+0.03), y + Inches(0.07),
                        Inches(s_col_w[c]), Inches(0.35), size=12,
                        bold=(r == 2), color=col)

    # フェーズ別
    add_textbox(sl, "フェーズ別収支",
                Inches(0.6), Inches(3.9), Inches(5.5), Inches(0.4),
                size=13, bold=True, color=C_ACCENT)
    phases = [
        ("立ち上げ期", "0〜3ヶ月", "360万", "▲72万", "C_RED"),
        ("成長期",     "3〜6ヶ月", "504万", "＋14万", "C_GRAY"),
        ("安定期",     "6ヶ月〜",  "648万", "＋142万", "C_GREEN"),
    ]
    for i, (ph, period, sales, profit, cc) in enumerate(phases):
        y = Inches(4.35 + i * 0.62)
        pcolor = C_RED if cc == "C_RED" else (C_GREEN if cc == "C_GREEN" else C_GRAY)
        add_rect(sl, Inches(0.5), y, Inches(5.8), Inches(0.58),
                 RGBColor(0x2A, 0x1E, 0x12))
        add_textbox(sl, ph, Inches(0.65), y + Inches(0.1),
                    Inches(1.5), Inches(0.38), size=13, bold=True, color=C_WHITE)
        add_textbox(sl, period, Inches(2.1), y + Inches(0.1),
                    Inches(1.2), Inches(0.38), size=12, color=C_GRAY)
        add_textbox(sl, sales, Inches(3.3), y + Inches(0.1),
                    Inches(1.3), Inches(0.38), size=13, color=C_LIGHT)
        add_textbox(sl, profit, Inches(4.6), y + Inches(0.1),
                    Inches(1.4), Inches(0.38), size=14, bold=True, color=pcolor)

    # 右：損益分岐
    add_rect(sl, Inches(7.0), Inches(1.7), Inches(5.8), Inches(5.0),
             RGBColor(0x2A, 0x1E, 0x12))
    add_textbox(sl, "損益分岐点分析",
                Inches(7.1), Inches(1.8), Inches(5.6), Inches(0.4),
                size=13, bold=True, color=C_ACCENT)

    bep_items = [
        ("損益分岐点売上", "460万円/月"),
        ("目標売上", "720万円/月"),
        ("目標対比", "64%で黒字転換"),
        ("固定費合計", "約175万円/月"),
        ("FL比率（安定期）", "48.6%"),
        ("安定期営業利益", "約142万円/月"),
        ("月次返済（友人＋公庫）", "約20万円/月"),
        ("返済後手残り（安定期）", "約122万円/月"),
    ]
    for i, (k, v) in enumerate(bep_items):
        y = Inches(2.3 + i * 0.54)
        add_rect(sl, Inches(7.0), y, Inches(5.8), Inches(0.5),
                 RGBColor(0x22, 0x18, 0x0E) if i % 2 == 0 else RGBColor(0x2A, 0x1E, 0x12))
        add_textbox(sl, k, Inches(7.15), y + Inches(0.08),
                    Inches(3.2), Inches(0.38), size=12, color=C_GRAY)
        vc = C_ACCENT if "利益" in k or "手残り" in k else C_WHITE
        add_textbox(sl, v, Inches(10.3), y + Inches(0.08),
                    Inches(2.3), Inches(0.38), size=13, bold=True,
                    color=vc, align=PP_ALIGN.RIGHT)

    slide_number(sl, 8)
    return sl


def s09_finance(prs):
    """資金計画"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "FUNDING", Inches(0.6), Inches(0.4), Inches(5), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "資金調達計画", Inches(0.6), Inches(0.75),
                Inches(9), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    # 調達サマリー
    funding = [
        ("代表 自己資金", "630万円", "90%持分（株式）"),
        ("友人 出資", "70万円", "10%持分（株式）"),
        ("友人 株主借入金", "1,430万円", "金利1%・10年返済"),
        ("日本政策金融公庫", "500万円", "金利2%・7年返済"),
    ]
    add_textbox(sl, "調達構成",
                Inches(0.6), Inches(1.75), Inches(5.5), Inches(0.4),
                size=13, bold=True, color=C_ACCENT)
    total = 0
    for i, (name, amount, note) in enumerate(funding):
        y = Inches(2.2 + i * 0.72)
        add_rect(sl, Inches(0.5), y, Inches(5.8), Inches(0.65),
                 RGBColor(0x2A, 0x1E, 0x12))
        add_textbox(sl, name, Inches(0.65), y + Inches(0.05),
                    Inches(2.2), Inches(0.55), size=13, color=C_LIGHT)
        add_textbox(sl, amount, Inches(2.8), y + Inches(0.05),
                    Inches(1.5), Inches(0.55), size=16, bold=True, color=C_ACCENT)
        add_textbox(sl, note, Inches(4.2), y + Inches(0.08),
                    Inches(2.0), Inches(0.5), size=11, color=C_GRAY)

    # 合計
    add_rect(sl, Inches(0.5), Inches(5.1), Inches(5.8), Inches(0.65),
             RGBColor(0x3A, 0x2A, 0x08))
    add_textbox(sl, "調達総額", Inches(0.65), Inches(5.17),
                Inches(2.0), Inches(0.52), size=14, bold=True, color=C_WHITE)
    add_textbox(sl, "2,630万円", Inches(2.8), Inches(5.17),
                Inches(2.5), Inches(0.52), size=20, bold=True, color=C_ACCENT)

    # 使途
    add_textbox(sl, "資金使途",
                Inches(7.0), Inches(1.75), Inches(5.5), Inches(0.4),
                size=13, bold=True, color=C_ACCENT)
    uses = [
        ("物件取得費（保証金・礼金）", "350万円"),
        ("内装・設備工事", "870万円"),
        ("厨房機器・備品", "200万円"),
        ("仕入・食材初回発注", "50万円"),
        ("採用・オープン前研修", "80万円"),
        ("マーケティング（3ヶ月分）", "45万円"),
        ("運転資金（3ヶ月分）", "405万円"),
        ("🟡 手元残金（緊急予備）", "630万円"),
    ]
    for i, (use, amt) in enumerate(uses):
        y = Inches(2.2 + i * 0.57)
        bg_c = RGBColor(0x3A, 0x2A, 0x08) if "手元" in use else RGBColor(0x2A, 0x1E, 0x12)
        add_rect(sl, Inches(7.0), y, Inches(5.8), Inches(0.53), bg_c)
        col = C_ACCENT if "手元" in use else C_LIGHT
        add_textbox(sl, use.replace("🟡 ", ""), Inches(7.15), y + Inches(0.07),
                    Inches(3.8), Inches(0.4), size=12, color=col)
        add_textbox(sl, amt, Inches(10.8), y + Inches(0.07),
                    Inches(1.8), Inches(0.4), size=13, bold=True,
                    color=col, align=PP_ALIGN.RIGHT)

    add_textbox(sl,
                "手元残金630万円は立ち上げ期の赤字（▲216万）を十分カバー。財務安全性は高い。",
                Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.35),
                size=12, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)

    slide_number(sl, 9)
    return sl


def s10_roadmap(prs):
    """展開ロードマップ"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "ROADMAP", Inches(0.6), Inches(0.4), Inches(5), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "展開ロードマップ　7年で100店舗へ", Inches(0.6), Inches(0.75),
                Inches(11), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    phases = [
        {
            "phase": "実証期",
            "period": "1〜2年目",
            "stores": "1〜2店舗",
            "action": "直営1号店で型をつくる\nQSC・レシピ・教育マニュアル完成\nインフルエンサー戦略で認知獲得",
            "kpi": "月商黒字化\nリピート率60%以上",
        },
        {
            "phase": "横展開期",
            "period": "3年目",
            "stores": "2〜5店舗",
            "action": "直営2〜3号店 ＋ FC募集開始\nFC加盟オーナー教育体制構築\n三軒茶屋・目黒・自由が丘エリアへ",
            "kpi": "FC1号店開業\nグループ月商2,000万超",
        },
        {
            "phase": "加速期",
            "period": "4〜5年目",
            "stores": "20〜50店舗",
            "action": "FC中心に年15〜20店舗ペース\n魚酒場 然・鶏酒場 然も展開開始\nSV体制・本部機能を整備",
            "kpi": "FC比率70%以上\n年商10億超",
        },
        {
            "phase": "確立期",
            "period": "6〜7年目",
            "stores": "100店舗",
            "action": "全国展開・上場検討\nブランド価値の確立\n「然」を食文化のアイコンへ",
            "kpi": "100店舗達成\n上場準備",
        },
    ]

    for i, p in enumerate(phases):
        x = Inches(0.5 + i * 3.2)
        # フェーズ帯
        add_rect(sl, x, Inches(1.75), Inches(3.0), Inches(0.45), C_ACCENT)
        add_textbox(sl, f"{p['phase']}　{p['period']}",
                    x + Inches(0.1), Inches(1.78), Inches(2.8), Inches(0.38),
                    size=13, bold=True, color=C_BG)
        # 店舗数
        add_rect(sl, x, Inches(2.2), Inches(3.0), Inches(5.15),
                 RGBColor(0x2A, 0x1E, 0x12))
        add_textbox(sl, p["stores"],
                    x + Inches(0.1), Inches(2.3), Inches(2.8), Inches(0.6),
                    size=26, bold=True, color=C_ACCENT if i == 3 else C_WHITE)
        add_line(sl, x + Inches(0.1), Inches(3.0), Inches(2.7), color=C_ACCENT,
                 height=Pt(0.8))
        add_textbox(sl, p["action"],
                    x + Inches(0.1), Inches(3.1), Inches(2.8), Inches(2.3),
                    size=11, color=C_LIGHT)
        # KPI
        add_rect(sl, x, Inches(5.35), Inches(3.0), Inches(0.8),
                 RGBColor(0x1A, 0x12, 0x05))
        add_textbox(sl, "KPI: " + p["kpi"],
                    x + Inches(0.1), Inches(5.4), Inches(2.8), Inches(0.68),
                    size=11, color=C_ACCENT)

    # 矢印（テキストで代用）
    for i in range(3):
        add_textbox(sl, "→", Inches(3.3 + i * 3.2), Inches(3.8),
                    Inches(0.25), Inches(0.4), size=20, color=C_ACCENT,
                    align=PP_ALIGN.CENTER)

    add_textbox(sl, "FC展開により資本効率を最大化。オーナーのノウハウを型にして横展開する。",
                Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.35),
                size=12, color=C_GRAY, align=PP_ALIGN.CENTER, italic=True)

    slide_number(sl, 10)
    return sl


def s11_strategy(prs):
    """集客・マーケ戦略"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    add_textbox(sl, "MARKETING", Inches(0.6), Inches(0.4), Inches(5), Inches(0.5),
                size=11, bold=True, color=C_ACCENT)
    add_textbox(sl, "集客・マーケティング戦略", Inches(0.6), Inches(0.75),
                Inches(10), Inches(0.7), size=28, bold=True, color=C_WHITE)
    add_line(sl, Inches(0.6), Inches(1.55), Inches(11.5))

    strategies = [
        {
            "title": "① インフルエンサー戦略",
            "ref": "ひまり商店モデル（初月1,700万円）",
            "body": "開業3ヶ月前からマイクロ〜メガ\nインフルエンサーをリスト化。\n無料招待→動画・リール投稿で\nオープン前から認知を爆発させる。",
        },
        {
            "title": "② Instagram中心のSNS設計",
            "ref": "石板の煙・音・断面美を動画化",
            "body": "撮られることを前提にした\n商品・空間・BGMの設計。\n「映え」ではなく「体験動画」\nとして拡散される仕掛けを作る。",
        },
        {
            "title": "③ MEO（Googleマップ）対策",
            "ref": "「肉酒場 三軒茶屋」で1位獲得",
            "body": "業態説明に「糀漬け肉定食居酒屋」\nを明記。写真・口コミ・投稿頻度\nで上位表示を維持する。\nAI検索（AIO）にも対応した記述。",
        },
        {
            "title": "④ ランチ×ディナー来店設計",
            "ref": "一人の客が昼も夜も来る店に",
            "body": "ランチで日常客を獲得し、\n夜の酒場に引き上げる。\nLINE公式・スタンプカードで\nリピートを仕組み化する。",
        },
    ]

    for i, s in enumerate(strategies):
        col = i % 2
        row = i // 2
        x = Inches(0.5 + col * 6.4)
        y = Inches(1.75 + row * 2.65)
        add_rect(sl, x, y, Inches(6.1), Inches(2.5), RGBColor(0x2A, 0x1E, 0x12))
        add_textbox(sl, s["title"], x + Inches(0.15), y + Inches(0.1),
                    Inches(5.8), Inches(0.4), size=15, bold=True, color=C_WHITE)
        add_textbox(sl, s["ref"], x + Inches(0.15), y + Inches(0.52),
                    Inches(5.8), Inches(0.3), size=11, color=C_ACCENT, italic=True)
        add_line(sl, x + Inches(0.15), y + Inches(0.85), Inches(5.7),
                 color=C_GRAY, height=Pt(0.5))
        add_textbox(sl, s["body"], x + Inches(0.15), y + Inches(0.95),
                    Inches(5.8), Inches(1.3), size=12, color=C_LIGHT)

    slide_number(sl, 11)
    return sl


def s12_closing(prs):
    """クロージング"""
    sl = blank_slide(prs)
    bg(sl)
    add_rect(sl, 0, 0, Inches(0.25), SLIDE_H, C_ACCENT)

    # 大きな然
    add_textbox(sl, "然", Inches(7.5), Inches(0.2), Inches(5.5), Inches(6.5),
                size=280, bold=True, color=RGBColor(0x2A, 0x1E, 0x12),
                align=PP_ALIGN.CENTER)

    add_textbox(sl, "私たちが目指す未来",
                Inches(0.8), Inches(1.2), Inches(6.5), Inches(0.6),
                size=16, bold=True, color=C_ACCENT)
    add_line(sl, Inches(0.8), Inches(1.85), Inches(5.5))

    add_textbox(sl,
                "糀漬け肉の定食と、和の一杯。",
                Inches(0.8), Inches(2.0), Inches(7.0), Inches(0.7),
                size=26, bold=True, color=C_WHITE)

    add_textbox(sl,
                "日本全国に「然」の灯りを。\n7年で100店舗。\n発酵と肉で、新しい定番をつくる。",
                Inches(0.8), Inches(2.9), Inches(6.5), Inches(1.6),
                size=20, color=C_LIGHT)

    add_textbox(sl,
                "肉酒場 然は、単なる飲食店ではありません。\n"
                "糀という日本の発酵文化と、現代人の「おいしくて、体にいい」\n"
                "という本質的なニーズを結びつけた、新しいブランドです。\n\n"
                "直営で型をつくり、FCで全国へ。\n"
                "その先には、日本の食文化に残る一皿があります。",
                Inches(0.8), Inches(4.7), Inches(6.5), Inches(2.0),
                size=13, color=C_GRAY, italic=True)

    slide_number(sl, 12)
    return sl


def main():
    prs = new_prs()

    s01_title(prs)
    s02_vision(prs)
    s03_why_now(prs)
    s04_concept(prs)
    s05_products(prs)
    s06_competitor(prs)
    s07_profile(prs)
    s08_pl(prs)
    s09_finance(prs)
    s10_roadmap(prs)
    s11_strategy(prs)
    s12_closing(prs)

    os.makedirs("reports", exist_ok=True)
    out = f"reports/肉酒場然_事業計画書_{datetime.now().strftime('%Y%m%d')}.pptx"
    prs.save(out)
    print(f"✅ 保存完了: {out}")


if __name__ == "__main__":
    main()
