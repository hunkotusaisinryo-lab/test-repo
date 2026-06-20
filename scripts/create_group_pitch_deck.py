#!/usr/bin/env python3
"""
然グループ 統合ピッチデッキ生成スクリプト
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# カラー定義
C_BG = RGBColor(250, 250, 248)       # オフホワイト
C_ACCENT = RGBColor(44, 74, 62)      # 深緑
C_GOLD = RGBColor(139, 105, 20)      # ゴールド
C_TEXT = RGBColor(26, 26, 26)        # テキスト
C_WHITE = RGBColor(255, 255, 255)
C_LIGHT = RGBColor(220, 230, 225)    # 薄緑

FONT = "IPAGothic"
W = Inches(13.33)
H = Inches(7.5)

def new_prs():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    return prs

def blank_slide(prs):
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)
    # 背景
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = C_BG
    return slide

def add_rect(slide, l, t, w, h, color):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h, size=18, bold=False, color=None, align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(l, t, w, h)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color if color else C_TEXT
    return txBox

def header(slide, title):
    add_rect(slide, 0, 0, W, Inches(0.9), C_ACCENT)
    add_text(slide, title, Inches(0.4), Inches(0.1), Inches(12), Inches(0.7),
             size=24, bold=True, color=C_WHITE)

def footer(slide):
    add_rect(slide, 0, H - Inches(0.35), W, Inches(0.35), C_ACCENT)
    add_text(slide, "然グループ 事業説明資料  |  Confidential", Inches(0.4), H - Inches(0.32),
             Inches(10), Inches(0.3), size=9, color=C_WHITE)

def add_table(slide, headers, rows, l, t, w, h, header_color=None):
    if header_color is None:
        header_color = C_ACCENT
    col_count = len(headers)
    row_count = len(rows) + 1
    table = slide.shapes.add_table(row_count, col_count, l, t, w, h).table
    col_w = w // col_count
    for i in range(col_count):
        table.columns[i].width = col_w

    # ヘッダー行
    for ci, hdr in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = hdr
        run.font.name = FONT
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = C_WHITE

    # データ行
    for ri, row in enumerate(rows):
        bg = C_LIGHT if ri % 2 == 0 else C_WHITE
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(val)
            run.font.name = FONT
            run.font.size = Pt(10)
            run.font.color.rgb = C_TEXT
    return table

def bullet_block(slide, items, l, t, w, h, title=None, title_color=None):
    if title_color is None:
        title_color = C_ACCENT
    y = t
    if title:
        add_rect(slide, l, y, w, Inches(0.35), title_color)
        add_text(slide, title, l + Inches(0.1), y + Inches(0.03), w - Inches(0.2), Inches(0.3),
                 size=12, bold=True, color=C_WHITE)
        y += Inches(0.38)
    for item in items:
        add_text(slide, "▶ " + item, l + Inches(0.1), y, w - Inches(0.2), Inches(0.38),
                 size=11, color=C_TEXT)
        y += Inches(0.38)

def big_number(slide, number, label, l, t, w, h):
    add_rect(slide, l, t, w, h, C_ACCENT)
    add_text(slide, number, l, t + Inches(0.1), w, Inches(0.6),
             size=32, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)
    add_text(slide, label, l, t + Inches(0.65), w, Inches(0.35),
             size=11, color=C_WHITE, align=PP_ALIGN.CENTER)


# ── スライド生成 ──────────────────────────────

def slide_cover(prs):
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, W, H, C_ACCENT)
    # 大きい「然」
    add_text(slide, "然", Inches(0.5), Inches(0.8), Inches(3), Inches(3.5),
             size=180, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)
    # タイトル
    add_text(slide, "然グループ", Inches(4), Inches(1.5), Inches(8.5), Inches(1.2),
             size=48, bold=True, color=C_WHITE)
    add_text(slide, "事業説明資料", Inches(4), Inches(2.7), Inches(8.5), Inches(0.9),
             size=36, bold=False, color=C_LIGHT)
    add_text(slide, "食の全シーンに、然を。", Inches(4), Inches(3.8), Inches(8.5), Inches(0.7),
             size=20, bold=False, color=C_GOLD)
    add_text(slide, "2026年6月  |  Confidential", Inches(4), Inches(5.5), Inches(8.5), Inches(0.5),
             size=13, color=C_LIGHT)


def slide_vision(prs):
    slide = blank_slide(prs)
    header(slide, "グループビジョン")
    add_text(slide, "「食の全シーンに、然を。」",
             Inches(1), Inches(1.2), Inches(11), Inches(1.1),
             size=34, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)
    add_text(slide, "日本の食文化の本質（出汁・発酵・素材）を核に、\nランチから夜まで・スタンドから体験型まで・国内から海外まで。\n飲食業界の「新しいスタンダード」を創るグループ。",
             Inches(1.5), Inches(2.4), Inches(10), Inches(1.5),
             size=16, color=C_TEXT, align=PP_ALIGN.CENTER)

    brands = [
        ("🍜\n蕎麦 然", "立ち食い蕎麦FC\n客単価 1,500円"),
        ("🍲\nだし 然", "出汁体験レストラン\n客単価 1,800〜9,000円"),
        ("🥩\n肉酒場 然", "発酵×肉の居酒屋\n客単価 4,000〜6,000円"),
        ("📱\nSNS事業部", "飲食特化コンサル\n月額 9.8〜39.8万円"),
    ]
    bw = Inches(2.8)
    for i, (name, desc) in enumerate(brands):
        lx = Inches(0.5) + i * Inches(3.2)
        add_rect(slide, lx, Inches(4.2), bw, Inches(2.5), C_ACCENT)
        add_text(slide, name, lx, Inches(4.3), bw, Inches(1.0),
                 size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, desc, lx, Inches(5.3), bw, Inches(1.2),
                 size=12, color=C_LIGHT, align=PP_ALIGN.CENTER)
    footer(slide)


def slide_market(prs):
    slide = blank_slide(prs)
    header(slide, "市場機会")

    nums = [
        ("25兆円", "国内外食産業市場規模"),
        ("3,188万人", "2023年訪日外国人数\n（コロナ前超えの勢い）"),
        ("68%", "健康志向食品市場\n過去5年成長率"),
    ]
    for i, (n, l) in enumerate(nums):
        lx = Inches(0.5) + i * Inches(4.2)
        big_number(slide, n, l, lx, Inches(1.1), Inches(3.8), Inches(1.3))

    bullets = [
        "外食産業は2030年に向けて回復・成長トレンド継続",
        "インバウンド旅行者の「本物の日本食体験」需要が急増",
        "発酵食品・出汁・和食の健康志向は国内外で高まる一方",
        "立ち食い・スタンド業態はコスパ・時短ニーズで拡大中",
        "飲食FC市場：参入障壁が低く加盟希望者が増加傾向",
        "SNSによる集客格差が飲食業界で顕在化（勝者総取り化）",
    ]
    bullet_block(slide, bullets, Inches(0.5), Inches(3.0), Inches(12), Inches(4.0),
                 title="市場トレンドサマリー")
    footer(slide)


def slide_overview(prs):
    slide = blank_slide(prs)
    header(slide, "然グループ 全体像")
    headers = ["ブランド", "業態", "客単価", "ターゲット", "展開モデル", "月次利益目標"]
    rows = [
        ["蕎麦 然", "立ち食い蕎麦", "1,500円", "ビジネスパーソン（ランチ）", "FC特化", "132万円/店"],
        ["だし 然", "出汁定食・体験ディナー", "1,800〜9,000円", "ワーカー+インバウンド", "直営+FC", "330万円/店"],
        ["肉酒場 然", "発酵×肉の居酒屋", "4,000〜6,000円", "20〜40代・飲み会需要", "直営中心", "200万円/店"],
        ["SNS事業部", "飲食特化SNSコンサル", "月9.8〜39.8万円", "飲食店オーナー", "コンサル", "83万円/月"],
    ]
    add_table(slide, headers, rows, Inches(0.3), Inches(1.1), Inches(12.7), Inches(2.8))

    add_text(slide, "食の時間帯カバレッジ", Inches(0.5), Inches(4.2), Inches(4), Inches(0.4),
             size=13, bold=True, color=C_ACCENT)
    timeline = [
        ("蕎麦 然", "ランチ 11:00〜15:00", Inches(0.5), Inches(3.5)),
        ("だし 然", "ランチ〜ディナー 11:30〜22:00", Inches(3.0), Inches(3.5)),
        ("肉酒場 然", "ディナー 17:00〜23:00", Inches(7.0), Inches(3.5)),
    ]
    for name, time, lx, _ in timeline:
        add_rect(slide, lx, Inches(4.7), Inches(2.5), Inches(0.5), C_ACCENT)
        add_text(slide, name, lx, Inches(4.72), Inches(2.5), Inches(0.45),
                 size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, time, lx, Inches(5.25), Inches(2.5), Inches(0.35),
                 size=9, color=C_TEXT, align=PP_ALIGN.CENTER)

    add_text(slide, "→ グループ全体で食の全時間帯を網羅。顧客をグループ内で回遊させる仕組み。",
             Inches(0.5), Inches(5.8), Inches(12), Inches(0.5),
             size=12, bold=True, color=C_GOLD)
    footer(slide)


def slide_niku(prs):
    slide = blank_slide(prs)
    header(slide, "肉酒場 然  ─  発酵×肉の居酒屋")

    add_text(slide, "「発酵の香りと肉煙が満ちる、大人の隠れ酒場」",
             Inches(0.5), Inches(1.1), Inches(12), Inches(0.6),
             size=20, bold=True, color=C_ACCENT)

    left_items = [
        "発酵肉（麹漬け・味噌漬け）を核心に据えた独自メニュー",
        "炭火グリルをオープンキッチンで見せる体験型設計",
        "日本酒・焼酎との発酵ペアリングでドリンク単価向上",
        "22〜25坪・30〜35席・客単価 4,000〜6,000円",
    ]
    bullet_block(slide, left_items, Inches(0.5), Inches(1.9), Inches(5.8), Inches(3.5),
                 title="コンセプト・特徴")

    right_items = [
        "然グループの「顔」かつ 発酵文化の発信基地",
        "ディナー特化・予約制で客単価を安定させる",
        "SNS映えする炭火・発酵器具がコンテンツになる",
        "常連顧客育成で口コミによる自然集客を実現",
    ]
    bullet_block(slide, right_items, Inches(6.7), Inches(1.9), Inches(5.8), Inches(3.5),
                 title="戦略的位置づけ")

    kpis = [("月次売上目標", "600万円"), ("月次営業利益", "200万円"), ("利益率", "33%")]
    for i, (label, val) in enumerate(kpis):
        lx = Inches(0.5) + i * Inches(4.2)
        add_rect(slide, lx, Inches(5.8), Inches(3.8), Inches(1.0), C_ACCENT)
        add_text(slide, val, lx, Inches(5.85), Inches(3.8), Inches(0.55),
                 size=28, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)
        add_text(slide, label, lx, Inches(6.4), Inches(3.8), Inches(0.35),
                 size=11, color=C_WHITE, align=PP_ALIGN.CENTER)
    footer(slide)


def slide_dashi(prs):
    slide = blank_slide(prs)
    header(slide, "だし 然  ─  出汁体験レストラン")

    add_text(slide, "「一杯の出汁が、今日を整える。」",
             Inches(0.5), Inches(1.1), Inches(12), Inches(0.6),
             size=20, bold=True, color=C_ACCENT)

    left_items = [
        "昼：出汁定食（1,600〜1,900円）地元ワーカー向け",
        "夜：出汁しゃぶ・体験コース（8,000〜12,000円）インバウンド向け",
        "昆布×本枯れ節×麹発酵のオリジナルブレンド出汁",
        "25坪・34席（カウンター10席）・多言語対応",
    ]
    bullet_block(slide, left_items, Inches(0.5), Inches(1.9), Inches(5.8), Inches(3.0),
                 title="ビジネスモデル（昼夜二毛作）")

    right_items = [
        "FC展開：加盟金300万円・ロイヤルティ3%",
        "Phase1：東京直営（浅草・銀座・上野）",
        "Phase2：関西FC（京都・大阪）",
        "Phase3：20店舗（直営5+FC15）",
        "Phase4：海外（台湾・シンガポール）",
    ]
    bullet_block(slide, right_items, Inches(6.7), Inches(1.9), Inches(5.8), Inches(3.0),
                 title="展開計画")

    kpis = [("月次売上", "960万円"), ("月次営業利益", "330万円"), ("投資回収", "15〜18ヶ月")]
    for i, (label, val) in enumerate(kpis):
        lx = Inches(0.5) + i * Inches(4.2)
        add_rect(slide, lx, Inches(5.5), Inches(3.8), Inches(1.0), C_ACCENT)
        add_text(slide, val, lx, Inches(5.55), Inches(3.8), Inches(0.55),
                 size=24, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)
        add_text(slide, label, lx, Inches(6.1), Inches(3.8), Inches(0.35),
                 size=11, color=C_WHITE, align=PP_ALIGN.CENTER)
    footer(slide)


def slide_soba(prs):
    slide = blank_slide(prs)
    header(slide, "蕎麦 然  ─  立ち食い蕎麦FC")

    add_text(slide, "「2分で届く、本物の蕎麦。」",
             Inches(0.5), Inches(1.1), Inches(12), Inches(0.6),
             size=20, bold=True, color=C_ACCENT)

    left_items = [
        "10坪・12スタンド・2人オペレーション",
        "客単価1,500円・回転時間12分・稼働率80%",
        "揚げたて天ぷら・北海道幌加内産蕎麦を使用",
        "QR前払い・キャッシュレス100%でオペレーション効率化",
    ]
    bullet_block(slide, left_items, Inches(0.5), Inches(1.9), Inches(5.8), Inches(3.0),
                 title="業態設計")

    right_items = [
        "FC加盟金200万円・ロイヤルティ4%",
        "標準内装パッケージで施工期間2.5ヶ月",
        "首都圏駅前・オフィス街から展開",
        "目標：30店舗（直営3+FC27）",
    ]
    bullet_block(slide, right_items, Inches(6.7), Inches(1.9), Inches(5.8), Inches(3.0),
                 title="FC展開モデル")

    kpis = [("月次売上", "348万円"), ("月次営業利益", "132万円"), ("BEP", "1日43人・投資回収8〜10ヶ月")]
    for i, (label, val) in enumerate(kpis):
        lx = Inches(0.5) + i * Inches(4.2)
        add_rect(slide, lx, Inches(5.5), Inches(3.8), Inches(1.0), C_ACCENT)
        add_text(slide, val, lx, Inches(5.55), Inches(3.8), Inches(0.55),
                 size=20, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)
        add_text(slide, label, lx, Inches(6.1), Inches(3.8), Inches(0.35),
                 size=11, color=C_WHITE, align=PP_ALIGN.CENTER)
    footer(slide)


def slide_sns(prs):
    slide = blank_slide(prs)
    header(slide, "SNS事業部  ─  飲食特化SNSコンサル")

    add_text(slide, "「飲食店のSNSを、売上に変える。」",
             Inches(0.5), Inches(1.1), Inches(12), Inches(0.6),
             size=20, bold=True, color=C_ACCENT)

    plan_headers = ["プラン", "月額", "主な内容", "対象"]
    plan_rows = [
        ["スターター", "9.8万円", "Instagram月12投稿＋リール4本", "単店舗・立ち上げ期"],
        ["スタンダード", "19.8万円", "Instagram+TikTok月20投稿＋撮影月1回", "複数店舗・グロース期"],
        ["フルサポート", "39.8万円〜", "全SNS＋専属担当＋ブランド戦略", "FC本部・大手チェーン"],
    ]
    add_table(slide, plan_headers, plan_rows, Inches(0.5), Inches(1.9), Inches(12), Inches(1.8))

    right_items = [
        "然グループ自身が証明した「SNSで集客する飲食店」ノウハウを外販",
        "毎日の飲食トレンドリサーチをクライアントのコンテンツに即反映",
        "FC加盟検討者へのクロスセルで受注確率向上",
        "然グループ全店のSNSを内製→広告費年間300万円削減",
    ]
    bullet_block(slide, right_items, Inches(0.5), Inches(3.9), Inches(12), Inches(2.5),
                 title="競合優位・シナジー")

    kpis = [("Year3月次売上", "168万円"), ("Year3月次利益", "83万円"), ("年次収益（Year3）", "1,000万円")]
    for i, (label, val) in enumerate(kpis):
        lx = Inches(0.5) + i * Inches(4.2)
        add_rect(slide, lx, Inches(6.3), Inches(3.8), Inches(0.9), C_ACCENT)
        add_text(slide, val, lx, Inches(6.32), Inches(3.8), Inches(0.5),
                 size=20, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)
        add_text(slide, label, lx, Inches(6.82), Inches(3.8), Inches(0.35),
                 size=10, color=C_WHITE, align=PP_ALIGN.CENTER)
    footer(slide)


def slide_synergy(prs):
    slide = blank_slide(prs)
    header(slide, "グループシナジー")

    synergies = [
        ("出汁素材の共通仕入れ", "だし然・蕎麦然・肉酒場然が同一出汁ベースを使用\n→ バイイングパワーで原価率1〜2%改善"),
        ("然カード（共通会員）", "全ブランドで使えるポイントカード\n→ 蕎麦然ランチ客→肉酒場然ディナーへの回遊促進"),
        ("SNS内製（コスト0）", "SNS事業部が然グループ全店のSNSを担当\n→ 広告費 年間300万円削減 + ブランド統一"),
        ("FC加盟者へのクロスセル", "蕎麦然・だし然FC加盟者にSNSコンサルを提案\n→ 自然な受注機会の創出"),
        ("セントラルキッチン（将来）", "将来的に共通仕込み場を設置\n→ 人件費・フードロス削減・品質均一化"),
    ]
    for i, (title, desc) in enumerate(synergies):
        lx = Inches(0.4) if i < 3 else Inches(0.4) + Inches(6.5) * (i - 3)
        ty = Inches(1.2) + (i % 3) * Inches(1.85) if i < 3 else Inches(1.2) + Inches(1.85)
        if i >= 3:
            lx = Inches(0.4) + Inches(6.5) * (i - 3)
            ty = Inches(1.2) + Inches(1.85)

        add_rect(slide, lx, ty, Inches(5.8), Inches(1.6), C_ACCENT)
        add_text(slide, "▶ " + title, lx + Inches(0.1), ty + Inches(0.05),
                 Inches(5.6), Inches(0.4), size=13, bold=True, color=C_GOLD)
        add_text(slide, desc, lx + Inches(0.1), ty + Inches(0.45),
                 Inches(5.6), Inches(1.1), size=11, color=C_WHITE)

    footer(slide)


def slide_5year(prs):
    slide = blank_slide(prs)
    header(slide, "5カ年数値計画")

    hdrs = ["年度", "総店舗数", "グループ月商", "年商", "EBITDA（推定）", "主要マイルストーン"]
    rows = [
        ["Year 1\n(2026)", "3店", "1,200万円", "1.4億円", "2,000万円", "だし然・蕎麦然1号店開業"],
        ["Year 2\n(2027)", "9店", "3,000万円", "3.6億円", "7,000万円", "FC展開開始・肉酒場然2号店"],
        ["Year 3\n(2028)", "22店", "6,000万円", "7.2億円", "1.8億円", "関西進出・SNS事業部開始"],
        ["Year 4\n(2029)", "35店", "10,000万円", "12億円", "3.5億円", "持株会社設立・海外準備"],
        ["Year 5\n(2030)", "55店", "17,000万円", "20億円", "6億円", "上場・EXIT準備"],
    ]
    add_table(slide, hdrs, rows, Inches(0.3), Inches(1.1), Inches(12.7), Inches(4.0))

    add_rect(slide, Inches(0.3), Inches(5.4), Inches(12.7), Inches(0.8), C_GOLD)
    add_text(slide, "Year 5 目標：年商 20億円  /  累計店舗 55店  /  EBITDA 6億円",
             Inches(0.5), Inches(5.45), Inches(12.3), Inches(0.7),
             size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    footer(slide)


def slide_expansion(prs):
    slide = blank_slide(prs)
    header(slide, "店舗展開マップ")

    phases = [
        ("Phase 1\n2026〜2027", "東京・首都圏",
         ["だし 然：浅草・銀座・上野・渋谷（直営4店）",
          "蕎麦 然：新橋・神田・大手町（直営+FC計5店）",
          "肉酒場 然：2号店展開"],
         Inches(0.3)),
        ("Phase 2\n2027〜2028", "関西進出",
         ["だし 然 FC：京都祇園・大阪道頓堀",
          "蕎麦 然 FC：大阪・神戸の駅前",
          "インバウンド需要の高いエリア優先"],
         Inches(4.5)),
        ("Phase 3\n2029〜2030", "全国・海外",
         ["蕎麦 然：30店舗（直営3+FC27）",
          "だし 然：20店舗（直営5+FC15）",
          "海外：台湾・シンガポール（DASHI ZEN）"],
         Inches(8.7)),
    ]

    for title, area, items, lx in phases:
        add_rect(slide, lx, Inches(1.1), Inches(3.9), Inches(5.8), C_ACCENT)
        add_text(slide, title, lx + Inches(0.1), Inches(1.15), Inches(3.7), Inches(0.7),
                 size=16, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)
        add_rect(slide, lx, Inches(1.85), Inches(3.9), Inches(0.4), C_GOLD)
        add_text(slide, area, lx + Inches(0.1), Inches(1.87), Inches(3.7), Inches(0.36),
                 size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        for j, item in enumerate(items):
            add_text(slide, "▶ " + item, lx + Inches(0.15), Inches(2.35) + j * Inches(0.55),
                     Inches(3.6), Inches(0.5), size=11, color=C_WHITE)
    footer(slide)


def slide_revenue(prs):
    slide = blank_slide(prs)
    header(slide, "収益モデルサマリー")

    hdrs = ["ブランド", "1店舗月次利益", "目標店舗数", "グループ月次利益"]
    rows = [
        ["肉酒場 然", "200万円", "5店", "1,000万円"],
        ["だし 然（直営+FC）", "330万円/店 + FC収入", "20店", "2,200万円+"],
        ["蕎麦 然（直営+FC）", "132万円/店 + FC収入", "30店", "1,500万円+"],
        ["SNS事業部", "83万円", "─", "83万円+"],
        ["合計（安定期）", "─", "55店", "4,783万円〜"],
    ]
    add_table(slide, hdrs, rows, Inches(0.5), Inches(1.1), Inches(12), Inches(3.2))

    hdrs2 = ["FC店舗数", "月次ロイヤルティ収入（だし然3%）", "月次ロイヤルティ収入（蕎麦然4%）", "合計"]
    rows2 = [
        ["FC 10店", "288万円", "139万円", "427万円"],
        ["FC 25店", "720万円", "348万円", "1,068万円"],
        ["FC 40店", "1,152万円", "557万円", "1,709万円"],
    ]
    add_table(slide, hdrs2, rows2, Inches(0.5), Inches(4.5), Inches(12), Inches(2.0))

    add_text(slide, "FCロイヤルティ収入スケール（ストック収益）",
             Inches(0.5), Inches(4.3), Inches(8), Inches(0.3),
             size=12, bold=True, color=C_ACCENT)
    footer(slide)


def slide_funding(prs):
    slide = blank_slide(prs)
    header(slide, "資金調達計画")

    add_rect(slide, Inches(0.5), Inches(1.1), Inches(5.5), Inches(1.2), C_ACCENT)
    add_text(slide, "第1回調達目標", Inches(0.6), Inches(1.15), Inches(5.3), Inches(0.4),
             size=13, color=C_WHITE)
    add_text(slide, "5,000万円", Inches(0.6), Inches(1.55), Inches(5.3), Inches(0.65),
             size=36, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)

    add_rect(slide, Inches(6.5), Inches(1.1), Inches(6.3), Inches(1.2), C_ACCENT)
    add_text(slide, "調達方法", Inches(6.6), Inches(1.15), Inches(6.1), Inches(0.4),
             size=13, color=C_WHITE)
    add_text(slide, "日本政策金融公庫 + エンジェル投資家", Inches(6.6), Inches(1.55),
             Inches(6.1), Inches(0.65), size=16, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)

    hdrs = ["用途", "金額", "内訳"]
    rows = [
        ["だし 然 1号店 開業費", "3,000万円", "内装1,375万+厨房450万+敷金保証金+運転資金"],
        ["蕎麦 然 1号店 開業費", "930万円", "内装400万+設備200万+保証金+運転資金"],
        ["人材採用・研修費", "500万円", "店長・料理長採用・研修プログラム構築"],
        ["運転資金（6ヶ月）", "570万円", "グループ管理費・マーケティング費用"],
        ["合計", "5,000万円", ""],
    ]
    add_table(slide, hdrs, rows, Inches(0.5), Inches(2.6), Inches(12), Inches(3.3))

    add_text(slide, "第2回調達（Year 2〜3）：2〜3億円  /  用途：関西展開・FC加速  /  VCまたは事業会社との提携",
             Inches(0.5), Inches(6.2), Inches(12), Inches(0.4),
             size=11, bold=True, color=C_ACCENT)
    footer(slide)


def slide_risk(prs):
    slide = blank_slide(prs)
    header(slide, "リスクと対策")

    hdrs = ["リスク", "影響度", "対策"]
    rows = [
        ["物件取得競争", "高", "不動産仲介との早期リレーション構築・複数候補を常時確保"],
        ["FC加盟店の品質低下", "高", "SV（スーパーバイザー）制度・月次研修・本部仕入れで品質統一"],
        ["インバウンド需要の変動", "中", "ランチ（国内需要）でベース売上を確保・円安依存を排除"],
        ["人材採用難", "中", "SNSブランディングで「働きたい会社」化・待遇改善（利益還元）"],
        ["出汁素材の調達リスク", "低", "複数サプライヤーと年間契約・輸入品との組み合わせで安定化"],
    ]
    add_table(slide, hdrs, rows, Inches(0.3), Inches(1.1), Inches(12.7), Inches(5.0))
    footer(slide)


def slide_contact(prs):
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, W, H, C_ACCENT)
    add_text(slide, "然グループ", Inches(1), Inches(1.5), Inches(11), Inches(1.5),
             size=48, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, "事業説明資料に関するお問い合わせ",
             Inches(1), Inches(3.2), Inches(11), Inches(0.7),
             size=20, color=C_LIGHT, align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(3), Inches(4.1), Inches(7), Inches(1.8), RGBColor(30, 55, 45))
    add_text(slide, "然グループ 事業企画室",
             Inches(3.2), Inches(4.2), Inches(6.6), Inches(0.5),
             size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, "Email：＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
             Inches(3.2), Inches(4.75), Inches(6.6), Inches(0.5),
             size=16, color=C_LIGHT, align=PP_ALIGN.CENTER)
    add_text(slide, "Tel：＿＿＿＿＿＿＿＿",
             Inches(3.2), Inches(5.25), Inches(6.6), Inches(0.5),
             size=16, color=C_LIGHT, align=PP_ALIGN.CENTER)
    add_text(slide, "本資料は機密情報を含みます。無断転載・配布を禁じます。",
             Inches(1), Inches(6.5), Inches(11), Inches(0.5),
             size=10, color=C_LIGHT, align=PP_ALIGN.CENTER)


def main():
    prs = new_prs()

    slide_cover(prs)
    slide_vision(prs)
    slide_market(prs)
    slide_overview(prs)
    slide_niku(prs)
    slide_dashi(prs)
    slide_soba(prs)
    slide_sns(prs)
    slide_synergy(prs)
    slide_5year(prs)
    slide_expansion(prs)
    slide_revenue(prs)
    slide_funding(prs)
    slide_risk(prs)
    slide_contact(prs)

    out_dir = "zen/07_融資・投資家資料"
    os.makedirs(out_dir, exist_ok=True)
    out_path = f"{out_dir}/然グループ_統合ピッチデッキ.pptx"
    prs.save(out_path)
    print(f"保存完了: {out_path}")


if __name__ == "__main__":
    main()
