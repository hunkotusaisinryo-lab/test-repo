#!/usr/bin/env python3
"""
週次トレンドリサーチレポート生成スクリプト
毎週月曜日にGitHub Actionsから実行される
前週（月〜日）のresearch/*.mdをPowerPointにまとめる
"""

import os
import re
import datetime
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# カラー定義
COLOR_BG = RGBColor(0x1A, 0x1A, 0x2E)        # 濃紺（背景）
COLOR_ACCENT = RGBColor(0xE8, 0x9C, 0x45)     # 金（アクセント）
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)       # 白
COLOR_LIGHT = RGBColor(0xCC, 0xCC, 0xCC)      # ライトグレー
COLOR_SECTION = RGBColor(0x2D, 0x2D, 0x44)    # セクション背景


def get_last_week_dates():
    """前週（月〜日）の日付リストを返す"""
    today = datetime.date.today()
    last_monday = today - datetime.timedelta(days=today.weekday() + 7)
    return [last_monday + datetime.timedelta(days=i) for i in range(7)]


def parse_markdown(filepath):
    """markdownファイルをセクションごとに解析する"""
    with open(filepath, encoding="utf-8") as f:
        content = f.read()

    sections = {}
    current_section = None
    current_lines = []

    for line in content.split("\n"):
        if line.startswith("## "):
            if current_section:
                sections[current_section] = "\n".join(current_lines).strip()
            current_section = line[3:].strip()
            current_lines = []
        elif line.startswith("# "):
            sections["title"] = line[2:].strip()
        elif current_section:
            current_lines.append(line)

    if current_section:
        sections[current_section] = "\n".join(current_lines).strip()

    return sections


def extract_bullets(text, max_items=5):
    """テキストから箇条書き項目を抽出する"""
    bullets = []
    for line in text.split("\n"):
        line = line.strip()
        if line.startswith("- ") or line.startswith("・"):
            item = line.lstrip("- ").lstrip("・").strip()
            if item:
                bullets.append(item)
        elif line.startswith("**") and line.endswith("**"):
            bullets.append(line.strip("*"))
    return bullets[:max_items]


def set_slide_background(slide, prs, color):
    """スライド背景色を設定する"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, text, left, top, width, height,
                font_size=18, bold=False, color=None, align=PP_ALIGN.LEFT, wrap=True):
    """テキストボックスを追加する"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return txBox


def add_title_slide(prs, week_start, week_end, total_days):
    """表紙スライドを作成する"""
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, prs, COLOR_BG)

    W = prs.slide_width
    H = prs.slide_height

    # アクセントライン（上部）
    from pptx.util import Pt as PtUtil
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        0, 0, W, Inches(0.08)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_ACCENT
    shape.line.fill.background()

    # メインタイトル
    add_textbox(slide, "飲食トレンド週次レポート",
                Inches(0.8), Inches(2.2), Inches(8.4), Inches(1.2),
                font_size=36, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

    # 期間
    period = f"{week_start.strftime('%Y年%m月%d日')}（月）〜{week_end.strftime('%m月%d日')}（日）"
    add_textbox(slide, period,
                Inches(0.8), Inches(3.5), Inches(8.4), Inches(0.7),
                font_size=18, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)

    # サブテキスト
    add_textbox(slide, f"対象日数：{total_days}日分　｜　自動生成レポート",
                Inches(0.8), Inches(4.2), Inches(8.4), Inches(0.5),
                font_size=13, color=COLOR_LIGHT, align=PP_ALIGN.CENTER)

    # アクセントライン（下部）
    shape2 = slide.shapes.add_shape(
        1, 0, H - Inches(0.08), W, Inches(0.08)
    )
    shape2.fill.solid()
    shape2.fill.fore_color.rgb = COLOR_ACCENT
    shape2.line.fill.background()

    return slide


def add_daily_slide(prs, date, sections):
    """1日分のスライドを作成する"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, prs, COLOR_BG)

    W = prs.slide_width

    # 上部アクセントバー
    shape = slide.shapes.add_shape(1, 0, 0, W, Inches(0.06))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_ACCENT
    shape.line.fill.background()

    # 日付ヘッダー背景
    header_bg = slide.shapes.add_shape(1, 0, Inches(0.06), W, Inches(0.9))
    header_bg.fill.solid()
    header_bg.fill.fore_color.rgb = COLOR_SECTION
    header_bg.line.fill.background()

    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    weekday = weekdays[date.weekday()]
    date_str = f"{date.strftime('%Y年%m月%d日')}（{weekday}）"
    add_textbox(slide, date_str,
                Inches(0.4), Inches(0.12), Inches(5), Inches(0.7),
                font_size=20, bold=True, color=COLOR_WHITE)

    # ハイライト（左カラム）
    col1_left = Inches(0.4)
    col1_width = Inches(4.4)
    col2_left = Inches(5.1)
    col2_width = Inches(4.4)
    content_top = Inches(1.15)
    section_gap = Inches(0.3)

    y = content_top

    # 今日のハイライト
    add_textbox(slide, "▍ 今日のハイライト",
                col1_left, y, col1_width, Inches(0.35),
                font_size=12, bold=True, color=COLOR_ACCENT)
    y += Inches(0.35)

    highlight_text = sections.get("今日のハイライト", "（データなし）")
    highlight_clean = re.sub(r'\*+', '', highlight_text).strip()
    add_textbox(slide, highlight_clean[:200],
                col1_left, y, col1_width, Inches(1.0),
                font_size=10, color=COLOR_LIGHT)
    y += Inches(1.05) + section_gap

    # 話題の料理・食材
    add_textbox(slide, "▍ 話題の料理・食材",
                col1_left, y, col1_width, Inches(0.35),
                font_size=12, bold=True, color=COLOR_ACCENT)
    y += Inches(0.35)

    food_bullets = extract_bullets(sections.get("話題の料理・食材", ""), max_items=4)
    food_text = "\n".join(f"• {b}" for b in food_bullets) if food_bullets else "（データなし）"
    add_textbox(slide, food_text,
                col1_left, y, col1_width, Inches(1.2),
                font_size=10, color=COLOR_LIGHT)
    y += Inches(1.25) + section_gap

    # SNSトレンド
    add_textbox(slide, "▍ SNSトレンド",
                col1_left, y, col1_width, Inches(0.35),
                font_size=12, bold=True, color=COLOR_ACCENT)
    y += Inches(0.35)

    sns_bullets = extract_bullets(sections.get("SNSトレンド", ""), max_items=3)
    sns_text = "\n".join(f"• {b}" for b in sns_bullets) if sns_bullets else "（データなし）"
    add_textbox(slide, sns_text,
                col1_left, y, col1_width, Inches(0.9),
                font_size=10, color=COLOR_LIGHT)

    # 右カラム
    y2 = content_top

    # 注目の業態
    add_textbox(slide, "▍ 注目の業態・新店舗",
                col2_left, y2, col2_width, Inches(0.35),
                font_size=12, bold=True, color=COLOR_ACCENT)
    y2 += Inches(0.35)

    store_bullets = extract_bullets(sections.get("注目の業態・新店舗", ""), max_items=3)
    store_text = "\n".join(f"• {b}" for b in store_bullets) if store_bullets else "（データなし）"
    add_textbox(slide, store_text,
                col2_left, y2, col2_width, Inches(0.9),
                font_size=10, color=COLOR_LIGHT)
    y2 += Inches(0.95) + section_gap

    # 発酵・健康・肉料理
    add_textbox(slide, "▍ 発酵・健康・肉料理",
                col2_left, y2, col2_width, Inches(0.35),
                font_size=12, bold=True, color=COLOR_ACCENT)
    y2 += Inches(0.35)

    ferment_bullets = extract_bullets(sections.get("発酵・健康・肉料理関連", ""), max_items=3)
    ferment_text = "\n".join(f"• {b}" for b in ferment_bullets) if ferment_bullets else "（データなし）"
    add_textbox(slide, ferment_text,
                col2_left, y2, col2_width, Inches(0.9),
                font_size=10, color=COLOR_LIGHT)
    y2 += Inches(0.95) + section_gap

    # 事業アイデアメモ（ハイライト枠）
    idea_bg = slide.shapes.add_shape(
        1, col2_left, y2, col2_width, Inches(1.4)
    )
    idea_bg.fill.solid()
    idea_bg.fill.fore_color.rgb = RGBColor(0x2A, 0x1F, 0x0A)
    idea_bg.line.color.rgb = COLOR_ACCENT
    idea_bg.line.width = Pt(1)

    add_textbox(slide, "💡 事業へのアイデアメモ",
                col2_left + Inches(0.15), y2 + Inches(0.08), col2_width - Inches(0.3), Inches(0.3),
                font_size=11, bold=True, color=COLOR_ACCENT)

    idea_bullets = extract_bullets(sections.get("事業へのアイデアメモ", ""), max_items=2)
    idea_text = "\n".join(f"• {b}" for b in idea_bullets) if idea_bullets else "（データなし）"
    add_textbox(slide, idea_text,
                col2_left + Inches(0.15), y2 + Inches(0.42), col2_width - Inches(0.3), Inches(0.9),
                font_size=9, color=COLOR_LIGHT)

    return slide


def add_summary_slide(prs, all_ideas, week_start, week_end):
    """週次まとめスライドを作成する"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, prs, COLOR_BG)

    W = prs.slide_width

    shape = slide.shapes.add_shape(1, 0, 0, W, Inches(0.06))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_ACCENT
    shape.line.fill.background()

    add_textbox(slide, "週次まとめ｜事業アイデア総覧",
                Inches(0.4), Inches(0.2), Inches(9.2), Inches(0.7),
                font_size=24, bold=True, color=COLOR_WHITE)

    period = f"{week_start.strftime('%m/%d')}〜{week_end.strftime('%m/%d')} の気づき"
    add_textbox(slide, period,
                Inches(0.4), Inches(0.9), Inches(9.2), Inches(0.4),
                font_size=13, color=COLOR_ACCENT)

    # アイデア一覧
    y = Inches(1.4)
    for i, (date, ideas) in enumerate(all_ideas.items()):
        if not ideas:
            continue
        weekdays = ["月", "火", "水", "木", "金", "土", "日"]
        day_label = f"{date.strftime('%m/%d')}（{weekdays[date.weekday()]}）"
        add_textbox(slide, day_label,
                    Inches(0.4), y, Inches(1.8), Inches(0.35),
                    font_size=11, bold=True, color=COLOR_ACCENT)

        idea_text = "　" + "　／　".join(ideas[:2])
        add_textbox(slide, idea_text,
                    Inches(2.3), y, Inches(7.2), Inches(0.35),
                    font_size=10, color=COLOR_LIGHT)
        y += Inches(0.42)

        if y > Inches(6.8):
            break

    return slide


def generate_report(research_dir="research", output_dir="reports"):
    dates = get_last_week_dates()
    week_start = dates[0]
    week_end = dates[-1]

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # 表紙
    available_dates = []
    all_sections = {}
    for date in dates:
        filepath = Path(research_dir) / f"{date}.md"
        if filepath.exists():
            available_dates.append(date)
            all_sections[date] = parse_markdown(filepath)

    add_title_slide(prs, week_start, week_end, len(available_dates))

    # 日別スライド
    all_ideas = {}
    for date in available_dates:
        sections = all_sections[date]
        add_daily_slide(prs, date, sections)
        idea_bullets = extract_bullets(sections.get("事業へのアイデアメモ", ""), max_items=2)
        all_ideas[date] = idea_bullets

    # まとめスライド
    if all_ideas:
        add_summary_slide(prs, all_ideas, week_start, week_end)

    # 保存
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{output_dir}/weekly_report_{week_start}_{week_end}.pptx"
    prs.save(filename)
    print(f"レポート生成完了: {filename}")
    return filename


if __name__ == "__main__":
    generate_report()
