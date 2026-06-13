#!/usr/bin/env python3
"""
毎日の飲食トレンドリサーチスクリプト
GitHub Actionsから実行される
"""

import anthropic
import datetime
import os
import sys

def research_trends():
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    today = datetime.date.today().strftime("%Y-%m-%d")

    prompt = f"""今日（{today}）の飲食・グルメ・外食産業のトレンドをリサーチしてください。

以下の観点で調べてください：
1. 今話題の料理・食材・業態（日本国内）
2. SNS（Instagram・TikTok）で拡散している飲食コンテンツ
3. 東京・都市部で注目されている新業態・新店舗
4. 発酵食品・健康志向・肉料理に関連するトレンド
5. 飲食業界のビジネスモデル・集客に関するニュース

調査結果を以下のMarkdown形式でまとめてください：

# 飲食トレンドリサーチ {today}

## 今日のハイライト
（最も注目すべきトレンドを3行以内で）

## 話題の料理・食材
（箇条書きで5〜8件）

## 注目の業態・新店舗
（箇条書きで3〜5件）

## 発酵・健康・肉料理関連
（箇条書きで3〜5件）

## SNSトレンド
（箇条書きで3〜5件）

## ビジネス・集客ニュース
（箇条書きで3〜5件）

## 事業へのアイデアメモ
（今日のトレンドから自分のビジネスに活かせそうなアイデアを2〜3件）

---
*自動生成 by Claude API*
"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 5}],
        messages=[{"role": "user", "content": prompt}]
    )

    result_text = ""
    for block in response.content:
        if hasattr(block, "text"):
            result_text += block.text

    return result_text, today


def save_result(content, date):
    filepath = f"research/{date}.md"
    os.makedirs("research", exist_ok=True)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"保存完了: {filepath}")
    return filepath


if __name__ == "__main__":
    print("飲食トレンドリサーチを開始します...")
    content, date = research_trends()
    filepath = save_result(content, date)
    print("完了")
