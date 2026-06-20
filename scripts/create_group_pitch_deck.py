#!/usr/bin/env python3
"""
然グループ 統合ピッチデッキ生成スクリプト
"""

import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─── 定数 ───────────────────────────────────────────────
OUTPUT_DIR = Path("/home/user/test-repo/zen/07_融資・投資家資料")
OUTPUT_FILE = OUTPUT_DIR / "然グループ_統合ピッチデッキ.pptx"

FONT_NAME = "IPAGothic"
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

BG_COLOR   = RGBColor(0xFA, 0xFA, 0xF8)
ACCENT     = RGBColor(0x2C, 0x4A, 0x3E)
GOLD       = RGBColor(0x8B, 0x69, 0x14)
TEXT_COLOR = RGBColor(0x1A, 0x1A, 0x1A)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

HEADER_H = Inches(1.0)

# ─── ヘルパー関数 ────────────────────────────────────────

def set_bg(slide, prs):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR


def add_header_band(slide):
    shape = slide.shapes.add_shape(
        1,
        Inches(0), Inches(0),
        SLIDE_W, HEADER_H
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = ACCENT
    shape.line.fill.background()
    return shape


def add_title_in_header(slide, title_text, font_size=28):
    txBox = slide.shapes.add_textbox(
        Inches(0.3), Inches(0.1),
        Inches(12.5), HEADER_H - Inches(0.1)
    )
    tf = txBox.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.name = FONT_NAME
    run.font.size = Pt(font_size)
    run.font.color.rgb = WHITE
    run.font.bold = True
    p.alignment = PP_ALIGN.LEFT
    return txBox


def add_textbox(slide, left, top, width, height, text, font_size=14,
                color=None, bold=False, align=PP_ALIGN.LEFT, wrap=True):
    if color is None:
        color = TEXT_COLOR
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.name = FONT_NAME
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    p.alignment = align
    return txBox


def add_multiline_textbox(slide, left, top, width, height, lines,
                           font_size=14, color=None, bold=False,
                           align=PP_ALIGN.LEFT):
    if color is None:
        color = TEXT_COLOR
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.name = FONT_NAME
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        run.font.bold = bold
        p.alignment = align
    return txBox


def add_simple_table(slide, left, top, width, rows_data, col_widths=None,
                     header_bg=None, font_size=12):
    if header_bg is None:
        header_bg = ACCENT
    n_rows = len(rows_data)
    n_cols = len(rows_data[0]) if rows_data else 1
    row_h = Inches(0.45)
    table = slide.shapes.add_table(n_rows, n_cols, left, top,
                                   width, row_h * n_rows).table
    if col_widths:
        for ci, cw in enumerate(col_widths):
            table.columns[ci].width = cw

    for ri, row in enumerate(rows_data):
        for ci, cell_text in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = str(cell_text)
            tf = cell.text_frame
            for para in tf.paragraphs:
                for run in para.runs:
                    run.font.name = FONT_NAME
                    run.font.size = Pt(font_size)
                    if ri == 0:
                        run.font.color.rgb = WHITE
                        run.font.bold = True
                    else:
                        run.font.color.rgb = TEXT_COLOR
                para.alignment = PP_ALIGN.CENTER
            fill = cell.fill
            fill.solid()
            if ri == 0:
                fill.fore_color.rgb = header_bg
            elif ri % 2 == 0:
                fill.fore_color.rgb = RGBColor(0xF0, 0xF0, 0xED)
            else:
                fill.fore_color.rgb = WHITE
    return table


def slide_header(slide, prs, title):
    set_bg(slide, prs)
    add_header_band(slide)
    add_title_in_header(slide, title)


# ─── スライド生成 ────────────────────────────────────────

def make_slide_01_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, prs)

    shape = slide.shapes.add_shape(
        1, Inches(0), Inches(6.5), SLIDE_W, Inches(1.0)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = ACCENT
    shape.line.fill.background()

    add_textbox(slide, Inches(0.5), Inches(1.0), Inches(4), Inches(4),
                "然", font_size=200, color=ACCENT, bold=True,
                align=PP_ALIGN.CENTER)

    add_textbox(slide, Inches(4.5), Inches(1.8), Inches(8.5), Inches(1.2),
                "然グループ", font_size=40, color=ACCENT, bold=True)

    add_textbox(slide, Inches(4.5), Inches(2.8), Inches(8.5), Inches(1.0),
                "事業説明資料", font_size=32, color=TEXT_COLOR, bold=True)

    add_textbox(slide, Inches(4.5), Inches(3.9), Inches(8.5), Inches(0.6),
                "2026年6月", font_size=18, color=GOLD)

    add_textbox(slide, Inches(4.5), Inches(4.5), Inches(8.5), Inches(0.5),
                "然グループ 事業企画室", font_size=14, color=TEXT_COLOR)

    add_textbox(slide, Inches(0.5), Inches(6.55), Inches(12), Inches(0.8),
                "Confidential  --  然グループ Internal Document",
                font_size=11, color=WHITE, align=PP_ALIGN.CENTER)


def make_slide_02_vision(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "グループビジョン")

    add_textbox(slide, Inches(1), Inches(1.2), Inches(11), Inches(1.0),
                "食の全シーンに、然を。",
                font_size=36, color=ACCENT, bold=True, align=PP_ALIGN.CENTER)

    add_textbox(slide, Inches(1.5), Inches(2.3), Inches(10), Inches(0.6),
                "日本の「だし・発酵・素材」文化を現代の食卓とビジネスに再定義する",
                font_size=16, color=GOLD, align=PP_ALIGN.CENTER)

    icons = [
        ("[ 食体験 ]", "本物のだし・発酵・肉文化を\n直営店舗で提供"),
        ("[ フランチャイズ ]", "蕎麦然・だし然を軸とした\n全国展開モデル"),
        ("[ SNS・コンサル ]", "飲食特化のSNS支援で\nブランドと収益を拡大"),
    ]
    for i, (title, body) in enumerate(icons):
        left = Inches(0.5 + i * 4.3)
        box = slide.shapes.add_shape(
            1, left, Inches(3.2), Inches(3.9), Inches(3.2)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0xE8, 0xF0, 0xED)
        box.line.color.rgb = ACCENT

        add_textbox(slide, left + Inches(0.1), Inches(3.3),
                    Inches(3.7), Inches(0.6),
                    title, font_size=16, color=ACCENT, bold=True,
                    align=PP_ALIGN.CENTER)
        add_multiline_textbox(slide, left + Inches(0.2), Inches(4.0),
                              Inches(3.5), Inches(2.0),
                              body.split("\n"), font_size=13,
                              align=PP_ALIGN.CENTER)


def make_slide_03_market(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "市場機会")

    nums = [
        ("25兆円", "国内外食産業市場規模", Inches(0.4)),
        ("4.3兆円", "インバウンド消費（2025年）", Inches(4.5)),
        ("急拡大", "健康・発酵食品トレンド", Inches(8.6)),
    ]
    for num, label, left in nums:
        add_textbox(slide, left, Inches(1.2), Inches(3.8), Inches(1.2),
                    num, font_size=40, color=ACCENT, bold=True,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, left, Inches(2.3), Inches(3.8), Inches(0.5),
                    label, font_size=12, color=GOLD, align=PP_ALIGN.CENTER)

    bullets = [
        "● 国内外食産業は約25兆円規模（2025年）。コロナ禍後の回復が続き、都市部を中心に拡大中",
        "● 訪日外国人消費は2025年に4.3兆円超。和食・発酵食品への関心が急増",
        "● 健康志向・発酵ブームにより、だし・糀・熟成肉カテゴリの成長率は年15〜20%超",
        "● 飲食FC市場も拡大中。低投資・高収益モデルへのニーズが加盟希望者を集める",
        "● SNS（TikTok・Instagram）を軸とした飲食プロモーション需要が急速に拡大",
    ]
    add_multiline_textbox(slide, Inches(0.5), Inches(3.1), Inches(12.3), Inches(4.0),
                          bullets, font_size=14)


def make_slide_04_overview(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "グループ事業概要")

    add_textbox(slide, Inches(0.5), Inches(1.1), Inches(12), Inches(0.4),
                "4つの事業が連携し、食の全シーンをカバーします",
                font_size=14, color=GOLD)

    rows = [
        ["事業", "ブランド", "業態", "客単価", "ターゲット"],
        ["飲食①", "肉酒場 然", "居酒屋・焼肉（直営）", "4,000〜6,000円", "30〜50代 男女"],
        ["飲食②", "だし 然", "だし料理・ランチ（FC）", "1,500〜2,500円", "20〜40代 女性・観光客"],
        ["飲食③", "蕎麦 然", "立ち食い蕎麦（FC）", "700〜1,200円", "全年代・ビジネス層"],
        ["デジタル", "SNS事業部", "飲食特化コンサル", "月9.8〜39.8万円", "中小飲食事業者"],
    ]
    col_widths = [Inches(1.4), Inches(2.0), Inches(3.0), Inches(2.8), Inches(3.5)]
    add_simple_table(slide, Inches(0.3), Inches(1.7), Inches(12.7), rows,
                     col_widths=col_widths, font_size=12)


def make_slide_05_nikunomise(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "肉酒場 然  --  フラッグシップ直営店")

    sections = [
        ("コンセプト",
         "「発酵と肉の融合」をテーマに、糀・塩麹・味噌を活かした熟成肉料理を提供。\n"
         "日本酒・クラフトビールとのペアリングで、大人の食体験を演出。"),
        ("差別化ポイント",
         "● 自社仕込みの発酵ダレ・漬けダレで他社との差別化\n"
         "● 食材ロスを削減する全頭利用メニュー構成\n"
         "● SNS映えする盛り付けとストーリーテリング"),
        ("現状実績",
         "● 現直営店舗：月商 約450万円（ランチ＋ディナー）\n"
         "● 客単価：平均 5,200円  /  FL比率 58%\n"
         "● リピート率：約40%（LINEミニアプリ計測）"),
        ("月次利益目標",
         "売上450万円  →  営業利益 約45〜60万円（利益率10〜13%）\n"
         "旗艦店の知名度を活かし、FC事業・SNS事業へ送客"),
    ]
    top = Inches(1.2)
    for title, body in sections:
        add_textbox(slide, Inches(0.5), top, Inches(12), Inches(0.35),
                    title, font_size=14, color=ACCENT, bold=True)
        top += Inches(0.35)
        add_multiline_textbox(slide, Inches(0.7), top, Inches(12), Inches(0.8),
                              body.split("\n"), font_size=13)
        top += Inches(0.9)


def make_slide_06_dashi(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "だし 然  --  だし体験特化型FC")

    add_textbox(slide, Inches(8.5), Inches(1.2), Inches(4.5), Inches(0.8),
                "月次利益目標: 330万円",
                font_size=20, color=GOLD, bold=True, align=PP_ALIGN.RIGHT)
    add_textbox(slide, Inches(8.5), Inches(2.0), Inches(4.5), Inches(0.5),
                "投資回収: 15〜18ヶ月",
                font_size=16, color=GOLD, align=PP_ALIGN.RIGHT)

    sections = [
        ("コンセプト",
         "本枯れ節・昆布・煮干しを軸とした「だし体験」専門店。\n"
         "飲み比べセット・だし茶漬け・だし鍋など、だしを主役にしたメニュー展開。"),
        ("ランチ/ディナー二毛作",
         "● ランチ：だし茶漬け・定食（客単価1,500〜2,000円）\n"
         "● ディナー：だし鍋・日本酒コース（客単価3,000〜4,000円）\n"
         "● 1坪あたり売上効率を最大化する設計"),
        ("FCモデル概要",
         "● 加盟金：300万円  /  保証金：50万円\n"
         "● ロイヤリティ：売上の5%\n"
         "● 初期投資総額目安：1,500〜2,000万円（物件除く）\n"
         "● 損益分岐点：月商 約210万円（目標月商350万円）"),
    ]
    top = Inches(2.6)
    for title, body in sections:
        add_textbox(slide, Inches(0.5), top, Inches(8), Inches(0.35),
                    title, font_size=14, color=ACCENT, bold=True)
        top += Inches(0.35)
        add_multiline_textbox(slide, Inches(0.7), top, Inches(12), Inches(0.85),
                              body.split("\n"), font_size=13)
        top += Inches(0.95)


def make_slide_07_soba(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "蕎麦 然  --  立ち食い蕎麦FC")

    metrics = [
        ("月次利益目標", "132万円"),
        ("投資回収", "8〜10ヶ月"),
        ("1日BEP", "43名"),
    ]
    for i, (label, val) in enumerate(metrics):
        left = Inches(0.5 + i * 4.2)
        box = slide.shapes.add_shape(
            1, left, Inches(1.2), Inches(3.8), Inches(1.1)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0xE8, 0xF0, 0xED)
        box.line.color.rgb = ACCENT
        add_textbox(slide, left, Inches(1.25), Inches(3.8), Inches(0.5),
                    label, font_size=12, color=GOLD, align=PP_ALIGN.CENTER)
        add_textbox(slide, left, Inches(1.65), Inches(3.8), Inches(0.5),
                    val, font_size=24, color=ACCENT, bold=True,
                    align=PP_ALIGN.CENTER)

    sections = [
        ("コンセプト",
         "石臼挽き十割蕎麦を立ち食いスタイルで提供。\n"
         "「本物を、早く、安く」をモットーに、駅前・商業施設での展開を想定。"),
        ("低投資・高回転FCモデル",
         "● 加盟金：150万円  /  保証金：30万円\n"
         "● ロイヤリティ：売上の3%\n"
         "● 初期投資目安：600〜900万円（物件除く）\n"
         "● 目標月商：220万円  /  客単価：850円  /  1日目標客数：85名"),
    ]
    top = Inches(2.5)
    for title, body in sections:
        add_textbox(slide, Inches(0.5), top, Inches(12), Inches(0.35),
                    title, font_size=14, color=ACCENT, bold=True)
        top += Inches(0.35)
        add_multiline_textbox(slide, Inches(0.7), top, Inches(12), Inches(0.85),
                              body.split("\n"), font_size=13)
        top += Inches(0.95)


def make_slide_08_sns(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "SNS事業部  --  飲食特化デジタルコンサルティング")

    add_textbox(slide, Inches(0.5), Inches(1.1), Inches(12), Inches(0.4),
                "Year3 売上目標: 1,000万円  /  自社ブランドのSNS戦略も内製化",
                font_size=15, color=GOLD, bold=True)

    rows = [
        ["プラン", "月額料金", "主なサービス内容", "想定クライアント"],
        ["ライト", "9.8万円",
         "Instagram投稿代行（週3回）ハッシュタグ戦略・月次レポート",
         "個人店・スタートアップ"],
        ["スタンダード", "19.8万円",
         "Instagram＋TikTok運用 リール制作（月4本）・広告運用支援",
         "中規模チェーン・FC加盟店"],
        ["プレミアム", "39.8万円",
         "全SNS統合運用＋PR企画 インフルエンサー連携・取材誘致支援",
         "多店舗展開企業・ブランド強化"],
    ]
    col_widths = [Inches(1.8), Inches(1.8), Inches(5.5), Inches(3.4)]
    add_simple_table(slide, Inches(0.3), Inches(1.7), Inches(12.7), rows,
                     col_widths=col_widths, font_size=12)

    add_multiline_textbox(slide, Inches(0.5), Inches(5.3), Inches(12), Inches(2.0),
                          [
                              "● 然グループ直営店・FC店を実績事例として活用し、新規クライアント獲得",
                              "● 飲食×SNSに特化した専門チームが運営。一般デジタルエージェンシーとの差別化を実現",
                              "● 将来的にはFC加盟店向けSNS支援パッケージとして標準化し、ロイヤリティ収入に組み込む予定",
                          ], font_size=13)


def make_slide_09_synergy(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "グループシナジー")

    add_textbox(slide, Inches(0.5), Inches(1.1), Inches(12), Inches(0.4),
                "4事業が相互に強化し合う「然エコシステム」",
                font_size=16, color=ACCENT, bold=True, align=PP_ALIGN.CENTER)

    synergies = [
        ("共同仕入れ（だし原材料）",
         "然グループ全店で本枯れ節・昆布などを一括購入。\n"
         "スケールメリットでコスト削減＋品質の統一。"),
        ("然カード（顧客囲い込み）",
         "グループ共通ポイント「然カード」で肉酒場・だし・蕎麦の\n"
         "クロスユースを促進。LTV向上と来店頻度アップ。"),
        ("内製SNS支援",
         "SNS事業部が全直営店・FC店の情報発信を担当。\n"
         "コスト削減と一貫したブランドイメージを実現。"),
        ("FCクロスセル",
         "蕎麦然オーナーへだし然の複数ブランド加盟を提案。\n"
         "1人のFCオーナーがグループ内複数ブランドを展開可能。"),
    ]
    for i, (title, body) in enumerate(synergies):
        row, col = divmod(i, 2)
        left = Inches(0.4 + col * 6.5)
        top = Inches(1.8 + row * 2.5)
        box = slide.shapes.add_shape(
            1, left, top, Inches(6.1), Inches(2.2)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0xE8, 0xF0, 0xED)
        box.line.color.rgb = ACCENT
        add_textbox(slide, left + Inches(0.1), top + Inches(0.05),
                    Inches(5.9), Inches(0.4),
                    title, font_size=14, color=ACCENT, bold=True)
        add_multiline_textbox(slide,
                              left + Inches(0.15), top + Inches(0.5),
                              Inches(5.8), Inches(1.6),
                              body.split("\n"), font_size=12)


def make_slide_10_financials(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "5ヵ年財務計画")

    add_textbox(slide, Inches(0.5), Inches(1.1), Inches(12), Inches(0.4),
                "Year5 グループ年商目標：20億円（直営＋FC＋SNS）",
                font_size=18, color=GOLD, bold=True, align=PP_ALIGN.CENTER)

    rows = [
        ["年度", "直営店舗数", "FC店舗数（累計）", "月商合計（万円）", "年商合計（億円）"],
        ["Year 1（2026）", "1", "2",  "800",    "0.96"],
        ["Year 2（2027）", "2", "8",  "2,200",  "2.64"],
        ["Year 3（2028）", "3", "20", "5,000",  "6.00"],
        ["Year 4（2029）", "4", "40", "10,000", "12.00"],
        ["Year 5（2030）", "5", "70", "17,000", "20.40"],
    ]
    col_widths = [Inches(2.2), Inches(2.0), Inches(2.5), Inches(2.5), Inches(2.5)]
    add_simple_table(slide, Inches(0.65), Inches(1.7), Inches(11.7), rows,
                     col_widths=col_widths, font_size=13)

    add_textbox(slide, Inches(0.5), Inches(6.4), Inches(12), Inches(0.7),
                "※ FC店舗数はだし然＋蕎麦然の合算。月商合計はロイヤリティ収入を含むグループ連結ベース。",
                font_size=11, color=GOLD)


def make_slide_11_expansion(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "店舗展開フェーズ")

    phases = [
        ("Phase 1（2026〜2027）\n首都圏集中展開",
         "● 東京都内（渋谷・新宿・品川・上野）を優先エリアに設定\n"
         "● 直営2店舗 ＋ FC試験導入8店舗\n"
         "● 成功モデルの標準化とマニュアル整備を完了"),
        ("Phase 2（2028〜2029）\n関西・主要都市展開",
         "● 大阪・京都・名古屋・福岡への出店\n"
         "● 地域FCパートナー制度を導入し加速展開\n"
         "● 40店舗体制でブランド認知を全国区へ"),
        ("Phase 3（2030〜）\n海外展開",
         "● 訪日客が多いアジア主要都市（バンコク・台北・シンガポール）を検討\n"
         "● 海外マスターFC方式での展開\n"
         "● 「JAPAN DASHI」として海外ブランド化"),
    ]
    for i, (title, body) in enumerate(phases):
        left = Inches(0.3 + i * 4.3)
        box = slide.shapes.add_shape(
            1, left, Inches(1.2), Inches(4.1), Inches(5.8)
        )
        box.fill.solid()
        if i == 0:
            box.fill.fore_color.rgb = RGBColor(0xD0, 0xE8, 0xD8)
        elif i == 1:
            box.fill.fore_color.rgb = RGBColor(0xE8, 0xF0, 0xED)
        else:
            box.fill.fore_color.rgb = RGBColor(0xF5, 0xF0, 0xE0)
        box.line.color.rgb = ACCENT

        add_multiline_textbox(slide,
                              left + Inches(0.1), Inches(1.3),
                              Inches(3.9), Inches(0.9),
                              title.split("\n"),
                              font_size=13, color=ACCENT, bold=True)
        add_multiline_textbox(slide,
                              left + Inches(0.15), Inches(2.3),
                              Inches(3.8), Inches(4.0),
                              body.split("\n"), font_size=12)


def make_slide_12_revenue(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "収益モデルサマリー")

    add_textbox(slide, Inches(0.5), Inches(1.1), Inches(12), Inches(0.4),
                "3ブランドの月次利益合計（単店ベース）とロイヤリティ収入スケール",
                font_size=14, color=GOLD)

    rows_profit = [
        ["ブランド", "目標月商", "月次利益（単店）", "利益率"],
        ["肉酒場 然", "450万円", "45〜60万円", "10〜13%"],
        ["だし 然（FC）", "350万円", "330万円※", "〜94%※"],
        ["蕎麦 然（FC）", "220万円", "132万円※", "〜60%※"],
    ]
    add_simple_table(slide, Inches(0.3), Inches(1.7), Inches(7.0), rows_profit,
                     col_widths=[Inches(2.0), Inches(1.8), Inches(1.8), Inches(1.4)],
                     font_size=12)

    add_textbox(slide, Inches(0.3), Inches(4.0), Inches(7.0), Inches(0.4),
                "※ FCはロイヤリティ収入ベース（加盟店売上×ロイヤリティ率）",
                font_size=10, color=GOLD)

    rows_royalty = [
        ["FC店舗数", "月次ロイヤリティ収入（目安）"],
        ["10店舗", "約100〜175万円/月"],
        ["20店舗", "約200〜350万円/月"],
        ["40店舗", "約400〜700万円/月"],
        ["70店舗", "約700〜1,225万円/月"],
    ]
    add_simple_table(slide, Inches(7.5), Inches(1.7), Inches(5.5), rows_royalty,
                     col_widths=[Inches(2.2), Inches(3.3)], font_size=12)

    add_textbox(slide, Inches(7.5), Inches(4.8), Inches(5.5), Inches(0.4),
                "※ だし然5%＋蕎麦然3%の平均から試算",
                font_size=10, color=GOLD)

    add_multiline_textbox(slide, Inches(0.5), Inches(5.0), Inches(12.3), Inches(2.0),
                          [
                              "● SNS事業部の月次収益（Year3）：約830万円（25クライアント×平均33万円）",
                              "● グループ合計（Year3見込み）：月次1,500〜2,000万円  /  年間1.8〜2.4億円",
                          ], font_size=13)


def make_slide_13_fundraising(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "資金調達計画")

    add_textbox(slide, Inches(0.5), Inches(1.1), Inches(12), Inches(0.6),
                "ラウンド1：5,000万円",
                font_size=36, color=ACCENT, bold=True, align=PP_ALIGN.CENTER)

    rows = [
        ["用途", "金額", "内訳・目的"],
        ["だし 然 出店準備", "3,000万円",
         "1〜2号店の内装・厨房設備・FCモデル確立費用"],
        ["蕎麦 然 出店準備", "930万円",
         "試験店舗（2店）の設備・マニュアル整備費用"],
        ["運転資金", "1,070万円",
         "開業後6ヶ月分の人件費・家賃・仕入れ資金"],
        ["合計", "5,000万円", "―"],
    ]
    col_widths = [Inches(2.5), Inches(2.0), Inches(8.2)]
    add_simple_table(slide, Inches(0.3), Inches(2.0), Inches(12.7), rows,
                     col_widths=col_widths, font_size=13)

    add_multiline_textbox(slide, Inches(0.5), Inches(5.2), Inches(12), Inches(2.0),
                          [
                              "【調達方法の想定】エクイティ（株式）またはコンバーティブルノートによる第三者割当増資",
                              "【バリュエーション】Year2売上見込みを基準に協議（詳細は別紙）",
                              "【ラウンド2以降】Year2末（2027年末）に追加調達を検討。関西展開・海外FC準備費用として2〜3億円を想定",
                          ], font_size=13)


def make_slide_14_risks(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, prs, "リスクと対策")

    rows = [
        ["リスク項目", "影響度", "対策"],
        ["食材原価高騰（円安・インフレ）",
         "高",
         "共同仕入れ・産地直送契約・メニュー価格の段階的見直し"],
        ["FC加盟者の品質管理",
         "中〜高",
         "SV制度・月次研修・ブランド監査で基準遵守を担保"],
        ["競合他社の類似コンセプト参入",
         "中",
         "然ブランドの商標登録・先行者優位・SNSによる認知拡大"],
        ["人材不足（調理・マネジメント）",
         "中",
         "パート比率最適化・グループ内人材ローテーション・採用専任担当設置"],
        ["SNS事業のクライアント離脱",
         "低〜中",
         "年間契約による収益安定化・飲食FC加盟店を優先クライアントに"],
    ]
    col_widths = [Inches(3.5), Inches(1.5), Inches(7.7)]
    add_simple_table(slide, Inches(0.3), Inches(1.3), Inches(12.7), rows,
                     col_widths=col_widths, font_size=12)


def make_slide_15_contact(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, prs)

    shape = slide.shapes.add_shape(
        1, Inches(0), Inches(0), SLIDE_W, Inches(2.0)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = ACCENT
    shape.line.fill.background()

    add_textbox(slide, Inches(1), Inches(0.3), Inches(11), Inches(0.8),
                "然グループ 事業企画室",
                font_size=32, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(1.1), Inches(11), Inches(0.6),
                "お問い合わせ・投資家向けお申し込み",
                font_size=16, color=RGBColor(0xCC, 0xDD, 0xCC),
                align=PP_ALIGN.CENTER)

    contact_items = [
        ("Email",  "（担当者よりご連絡いたします）"),
        ("Tel",    "（担当よりご案内いたします）"),
        ("Web",    "https://zen-group.jp  （準備中）"),
        ("所在地", "東京都（詳細はお問い合わせください）"),
    ]
    top = Inches(2.5)
    for label, value in contact_items:
        add_textbox(slide, Inches(1.5), top, Inches(2.5), Inches(0.5),
                    label + "：", font_size=14, color=GOLD, bold=True)
        add_textbox(slide, Inches(4.0), top, Inches(9.0), Inches(0.5),
                    value, font_size=14, color=TEXT_COLOR)
        top += Inches(0.7)

    add_textbox(slide, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.9),
                "本資料の内容は機密情報です。無断複製・転用・開示はお断りいたします。"
                "  /  (c) 2026 然グループ. All Rights Reserved.",
                font_size=11, color=GOLD, align=PP_ALIGN.CENTER)


# ─── メイン ─────────────────────────────────────────────

def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    make_slide_01_cover(prs)
    make_slide_02_vision(prs)
    make_slide_03_market(prs)
    make_slide_04_overview(prs)
    make_slide_05_nikunomise(prs)
    make_slide_06_dashi(prs)
    make_slide_07_soba(prs)
    make_slide_08_sns(prs)
    make_slide_09_synergy(prs)
    make_slide_10_financials(prs)
    make_slide_11_expansion(prs)
    make_slide_12_revenue(prs)
    make_slide_13_fundraising(prs)
    make_slide_14_risks(prs)
    make_slide_15_contact(prs)

    prs.save(str(OUTPUT_FILE))
    print("Saved: " + str(OUTPUT_FILE))
    print("Slides: " + str(len(prs.slides)))


if __name__ == "__main__":
    main()
