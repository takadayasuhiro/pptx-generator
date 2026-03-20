import json
import re
from openai import OpenAI
from app.config import GITHUB_TOKEN, AI_MODEL
from app.models import (
    GenerateRequest, PresentationData, SlideData,
    ChartData, StatItem,
)

client = OpenAI(
    base_url="https://models.github.ai/inference",
    api_key=GITHUB_TOKEN,
)

SYSTEM_PROMPT = """あなたはプレゼンテーション資料の構成専門家であり、データ分析の専門家でもあります。
指示に従い、純粋なJSON形式のみで出力してください。
コードブロック記法（```）は絶対に使わないでください。
CSVデータの分析結果が提供された場合は、その実データを正確に反映したチャートやstatsを作成してください。
chart の series.values には、提供された実データの数値をそのまま使ってください。架空の数値は使わないでください。
重要な指標にはstats（大きな数字の強調表示）を活用してください。
画像が効果的なスライドにはimage_keywordを必ず指定してください。"""

AVAILABLE_THEMES = [
    "ocean_blue", "midnight", "forest", "sunset",
    "lavender", "monochrome", "coral", "emerald",
]

AVAILABLE_LAYOUTS = [
    "title_slide", "section_header", "content",
    "two_column", "image_right", "chart", "stats", "closing",
]


def _build_source_context(req: GenerateRequest) -> str:
    sections = []

    if req.pdf_text:
        sections.append(f"""
■ 参照PDF文書の内容:
{req.pdf_text}
""")

    if req.csv_analysis:
        a = req.csv_analysis
        overview = a.get("overview", {})
        csv_text = f"""
■ CSVデータ分析結果:
- データ概要: {overview.get('rows', '?')}行 × {overview.get('columns', '?')}列
- 列名: {', '.join(overview.get('column_names', []))}
- 数値列: {', '.join(overview.get('numeric_columns', []))}
- カテゴリ列: {', '.join(overview.get('category_columns', []))}
"""

        if a.get("statistics"):
            csv_text += "\n- 基本統計量:\n"
            for col, stats in a["statistics"].items():
                mean = stats.get("mean", "?")
                std = stats.get("std", "?")
                min_v = stats.get("min", "?")
                max_v = stats.get("max", "?")
                csv_text += f"    {col}: 平均={mean}, 標準偏差={std}, 最小={min_v}, 最大={max_v}\n"

        if a.get("correlations"):
            csv_text += "\n- 注目すべき相関:\n"
            for c in a["correlations"][:5]:
                csv_text += f"    {c['col1']} ↔ {c['col2']} = {c['value']}\n"

        if a.get("trends"):
            csv_text += "\n- トレンド:\n"
            for col, trend in a["trends"].items():
                csv_text += f"    {col}: {trend}\n"

        if a.get("category_summary"):
            csv_text += "\n- カテゴリ別集計:\n"
            for col, counts in a["category_summary"].items():
                items = [f"{k}({v})" for k, v in list(counts.items())[:5]]
                csv_text += f"    {col}: {', '.join(items)}\n"

        if a.get("chart_ready_data"):
            csv_text += "\n- グラフ用データ（実データ・必ずこの数値を使うこと）:\n"
            for cd in a["chart_ready_data"]:
                csv_text += f"    タイプ: {cd['suggested_type']}, タイトル: {cd['title']}\n"
                csv_text += f"    categories: {json.dumps(cd['categories'], ensure_ascii=False)}\n"
                csv_text += f"    series: {json.dumps(cd['series'], ensure_ascii=False)}\n\n"

        if a.get("preview"):
            csv_text += f"\n- データプレビュー（先頭5行）:\n"
            csv_text += f"    {json.dumps(a['preview'], ensure_ascii=False, indent=2)}\n"

        sections.append(csv_text)

    return "\n".join(sections)


def _build_prompt(req: GenerateRequest) -> str:
    style_map = {
        "business": "ビジネス向けのプロフェッショナルなトーン",
        "casual": "カジュアルで親しみやすいトーン",
        "academic": "学術的で論理的なトーン",
    }
    style_desc = style_map.get(req.style, style_map["business"])
    lang = "すべて日本語で記述してください。" if req.language == "ja" else "Write everything in English."

    themes_str = ", ".join(AVAILABLE_THEMES)
    layouts_str = ", ".join(AVAILABLE_LAYOUTS)

    source_context = _build_source_context(req)

    has_csv = bool(req.csv_analysis)
    has_pdf = bool(req.pdf_text)

    source_instruction = ""
    if has_pdf and has_csv:
        source_instruction = "PDF文書の内容とCSVデータの分析結果の両方を基に、プレゼンテーションを構成してください。CSVの実データを使ったグラフを積極的に含めてください。"
    elif has_pdf:
        source_instruction = "PDF文書の内容を基に、要点をまとめたプレゼンテーションを構成してください。"
    elif has_csv:
        source_instruction = "CSVデータの分析結果を基に、データドリブンなプレゼンテーションを構成してください。chart_ready_dataに含まれる実データを使ったグラフを必ず含めてください。架空の数値を使わないでください。"

    chart_rules = ""
    if has_csv:
        chart_rules = f"""- CSVデータが提供されている場合、chart_ready_data の実データ（categories, series）をそのまま chart に使うこと
- CSVの分析結果に基づく chart スライドを少なくとも2枚含めること
- CSVの統計量に基づく stats スライドを少なくとも1枚含めること"""
    else:
        chart_rules = f"""- {req.num_slides}枚中、少なくとも1枚は layout: "chart" を含めること
- {req.num_slides}枚中、少なくとも1枚は layout: "stats" を含めること"""

    prompt = f"""以下の条件でプレゼンテーションスライドの内容を生成してください。

【トピック】{req.topic}
【スライド枚数】{req.num_slides}枚
【スタイル】{style_desc}
【言語】{lang}
{f"【追加指示】{req.additional_instructions}" if req.additional_instructions else ""}
{source_instruction}

{source_context}

■ 使用可能なテーマ: {themes_str}
トピックに最適なテーマを1つ選んでください。

■ 使用可能なレイアウト: {layouts_str}
各スライドに最適なレイアウトを選んでください。

以下のJSON形式で出力してください。

{{
  "title": "プレゼン全体のタイトル",
  "theme": "テーマ名",
  "slides": [
    {{
      "slide_number": 1,
      "layout": "title_slide",
      "title": "メインタイトル",
      "subtitle": "サブタイトル",
      "bullet_points": [],
      "image_keyword": "technology innovation",
      "chart": null,
      "stats": null,
      "speaker_notes": "スピーカーノート"
    }},
    {{
      "slide_number": 2,
      "layout": "stats",
      "title": "主要指標",
      "subtitle": "",
      "bullet_points": [],
      "image_keyword": "",
      "chart": null,
      "stats": [
        {{"value": "85%", "label": "顧客満足度"}},
        {{"value": "2.5M", "label": "アクティブユーザー"}}
      ],
      "speaker_notes": "スピーカーノート"
    }},
    {{
      "slide_number": 3,
      "layout": "chart",
      "title": "データの推移",
      "subtitle": "",
      "bullet_points": [],
      "image_keyword": "",
      "chart": {{
        "chart_type": "bar",
        "categories": ["A", "B", "C"],
        "series": [{{"name": "値", "values": [100, 200, 300]}}]
      }},
      "stats": null,
      "speaker_notes": "スピーカーノート"
    }}
  ]
}}

ルール:
- slide_number 1 は必ず layout: "title_slide" にし、image_keyword を付けること
- 最後のスライドは layout: "closing" にすること
- セクションの区切りには layout: "section_header" を使うこと
{chart_rules}
- {req.num_slides}枚中、少なくとも1枚は layout: "image_right" を含めること
- image_keyword は英語で指定すること（画像検索用）
- chart の chart_type は bar, line, pie, doughnut のいずれか
- chart の series の values は数値のリスト
- stats は2〜4個の項目が目安
- bullet_points は3〜5個が目安
- speaker_notes は各スライドに必ず付けること
- 純粋なJSONのみ出力すること
"""
    return prompt


def _parse_response(text: str) -> PresentationData:
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    cleaned = cleaned.strip()

    data = json.loads(cleaned)

    theme = data.get("theme", "ocean_blue")
    if theme not in AVAILABLE_THEMES:
        theme = "ocean_blue"

    slides = []
    for s in data["slides"]:
        chart_data = None
        if s.get("chart"):
            c = s["chart"]
            chart_data = ChartData(
                chart_type=c.get("chart_type", "bar"),
                categories=c.get("categories", []),
                series=c.get("series", []),
            )

        stats_data = []
        if s.get("stats"):
            for st in s["stats"]:
                stats_data.append(StatItem(value=st["value"], label=st["label"]))

        layout = s.get("layout", "content")
        if layout not in AVAILABLE_LAYOUTS:
            layout = "content"

        slides.append(SlideData(
            slide_number=s["slide_number"],
            layout=layout,
            title=s["title"],
            subtitle=s.get("subtitle", ""),
            bullet_points=s.get("bullet_points", []),
            image_keyword=s.get("image_keyword", ""),
            chart=chart_data,
            stats=stats_data,
            speaker_notes=s.get("speaker_notes", ""),
        ))

    return PresentationData(title=data["title"], theme=theme, slides=slides)


async def generate_presentation_content(req: GenerateRequest) -> PresentationData:
    prompt = _build_prompt(req)

    try:
        response = client.chat.completions.create(
            model=AI_MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt},
            ],
            temperature=0.7,
        )
    except Exception as e:
        err_str = str(e)
        if "429" in err_str or "rate" in err_str.lower():
            raise RuntimeError(
                "APIのレート制限に達しました。しばらく待ってから再度お試しください。"
            )
        raise RuntimeError(f"AI API エラー: {e}")

    content = response.choices[0].message.content
    if not content:
        raise RuntimeError("AI API から空のレスポンスが返されました。")

    try:
        return _parse_response(content)
    except (json.JSONDecodeError, KeyError) as e:
        raise RuntimeError(f"AIの応答をパースできませんでした: {e}")
