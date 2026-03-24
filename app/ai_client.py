import json
import re
from app.model_registry import get_client, get_available_models
from app.models import (
    GenerateRequest, PresentationData, SlideData,
    ChartData, StatItem,
)

SYSTEM_PROMPT = """あなたはプレゼンテーション資料の構成専門家であり、データ分析の専門家でもあります。
指示に従い、純粋なJSON形式のみで出力してください。
コードブロック記法（```）は絶対に使わないでください。
CSVまたはExcelデータの分析結果が提供された場合は、その実データを正確に反映したチャートやstatsを作成してください。
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


def _is_no_access_error(err: Exception) -> bool:
    s = str(err).lower()
    return "no_access" in s or "no access to model" in s


def _fallback_model_id_for(model_name: str) -> str | None:
    # GitHub Modelsで gpt-4o 権限がない環境向けに mini へ退避する。
    if model_name == "openai/gpt-4o":
        return "gpt-4o-mini"
    if model_name == "gpt-4o":
        return "openai-gpt-4o-mini"
    return None


def _build_fallback_model_ids(current_model_id: str | None = None) -> list[str]:
    available = [m for m in get_available_models() if m.get("available")]
    ids = [m["id"] for m in available if m.get("id")]

    preferred = [
        "gpt-4o-mini",
        "openai-gpt-4o-mini",
        "openai-gpt-4o",
        "gpt-4o",
    ]
    ollama_ids = [i for i in ids if i.startswith("ollama-")]
    # Qwen 3.5 を Ollama 内で優先（フォールバック時に最初に試す）
    preferred_ollama = ["ollama-qwen3.5"]
    other_ollama = [i for i in ollama_ids if i not in preferred_ollama]
    ollama_ordered = [i for i in preferred_ollama if i in ids] + other_ollama
    ordered = [i for i in preferred if i in ids] + ollama_ordered + [
        i for i in ids if i not in preferred and i not in ollama_ids
    ]
    if current_model_id:
        ordered = [i for i in ordered if i != current_model_id]
    return ordered


def _create_chat_with_fallback(
    model_id: str,
    messages: list[dict[str, str]],
    temperature: float,
):
    client, model_name, model_info = get_client(model_id)
    try:
        resp = client.chat.completions.create(
            model=model_name,
            messages=messages,
            temperature=temperature,
        )
        return resp, model_info.get("id", model_name)
    except Exception as e:
        if not _is_no_access_error(e):
            raise

        # まずはモデル名に応じた近い候補へ退避
        first_fb = _fallback_model_id_for(model_name)
        candidates: list[str] = []
        if first_fb:
            candidates.append(first_fb)
        candidates.extend(
            [m for m in _build_fallback_model_ids(model_info.get("id")) if m not in candidates]
        )

        last_err: Exception = e
        for fb_model_id in candidates:
            try:
                fb_client, fb_model_name, fb_info = get_client(fb_model_id)
                resp = fb_client.chat.completions.create(
                    model=fb_model_name,
                    messages=messages,
                    temperature=temperature,
                )
                return resp, fb_info.get("id", fb_model_name)
            except Exception as e2:
                last_err = e2
                if _is_no_access_error(e2):
                    continue
                raise

        raise RuntimeError(
            f"AI API エラー: 利用可能モデルにアクセスできません。最終エラー: {last_err}"
        )


def _build_source_context(req: GenerateRequest) -> str:
    sections = []

    if req.pdf_text:
        sections.append(f"""
■ 参照データ・文書の内容:
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


def _extract_first_json_object(text: str) -> str | None:
    """モデル応答から最初の JSON オブジェクト本文を抽出する。"""
    if not text:
        return None
    start = text.find("{")
    if start < 0:
        return None

    depth = 0
    in_str = False
    escaped = False
    for i in range(start, len(text)):
        ch = text[i]
        if in_str:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_str = False
            continue

        if ch == '"':
            in_str = True
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start:i + 1]
    return None


ANALYZE_SYSTEM_PROMPT = """あなたはデータ分析の専門家です。
提供されたデータ（CSV/Excel分析結果、PDF文書、メール、画像分析結果など）を読み解き、
ビジネスパーソンが理解しやすい日本語で、要点・洞察・提言を含む分析レポートを作成してください。
ユーザーが文字数・長さ・形式を指定した場合は、その指示を最優先で厳守してください。
指定がない場合は、マークダウンは使わずプレーンテキストで見やすく整形し、
セクションの区切りには「━━━」や「■」「▶」を使ってください。"""

GENERAL_SYSTEM_PROMPT = """あなたは知識豊富なビジネスアドバイザーであり、分析の専門家です。
ユーザーの質問や指示に対して、専門的かつ分かりやすい日本語で回答してください。
ユーザーが文字数・長さ・形式を指定した場合は、その指示を最優先で厳守してください。
指定がない場合は、マークダウンは使わずプレーンテキストで見やすく整形し、
セクションの区切りには「━━━」や「■」「▶」を使ってください。"""


def _build_analysis_context(csv_analysis: dict | None, pdf_text: str) -> str:
    sections = []
    if pdf_text:
        sections.append(f"■ 入力データの内容:\n{pdf_text}")

    if csv_analysis:
        a = csv_analysis
        overview = a.get("overview", {})
        text = f"■ CSVデータ分析結果:\n"
        text += f"データ概要: {overview.get('rows', '?')}行 × {overview.get('columns', '?')}列\n"
        text += f"列名: {', '.join(overview.get('column_names', []))}\n"

        if a.get("statistics"):
            text += "\n基本統計量:\n"
            for col, stats in a["statistics"].items():
                text += f"  {col}: 平均={stats.get('mean', '?')}, 標準偏差={stats.get('std', '?')}, 最小={stats.get('min', '?')}, 最大={stats.get('max', '?')}\n"

        if a.get("correlations"):
            text += "\n注目すべき相関:\n"
            for c in a["correlations"][:5]:
                text += f"  {c['col1']} ↔ {c['col2']} = {c['value']}\n"

        if a.get("trends"):
            text += "\nトレンド:\n"
            for col, trend in a["trends"].items():
                text += f"  {col}: {trend}\n"

        if a.get("category_summary"):
            text += "\nカテゴリ別集計:\n"
            for col, counts in a["category_summary"].items():
                items = [f"{k}({v}件)" for k, v in list(counts.items())[:8]]
                text += f"  {col}: {', '.join(items)}\n"

        if a.get("preview"):
            text += f"\nデータプレビュー（先頭5行）:\n"
            text += f"  {json.dumps(a['preview'], ensure_ascii=False, indent=2)}\n"

        sections.append(text)

    return "\n\n".join(sections)


async def generate_analysis_text(csv_analysis: dict | None = None,
                                  pdf_text: str = "",
                                  topic: str = "",
                                  model_id: str = "auto") -> tuple[str, str]:
    context = _build_analysis_context(csv_analysis, pdf_text)
    has_context = bool(context.strip())
    has_topic = bool(topic.strip())

    if not has_context and not has_topic:
        return "分析対象のデータがありません。", ""

    if has_context and has_topic:
        prompt = f"""以下のデータについて、ユーザーの指示に従って回答してください。

【ユーザーの指示】
{topic}

{context}"""
        system_prompt = ANALYZE_SYSTEM_PROMPT
    elif has_context:
        prompt = f"""以下のデータを分析し、包括的なレポートを作成してください。

{context}

以下の構成でレポートを作成してください:

1. データ概要（何のデータか、規模）
2. 主要な発見事項（3〜5点）
3. 数値の詳細分析（統計量、相関、トレンド）
4. カテゴリ別の特徴
5. 総合評価と提言（ビジネスへの示唆）

プレーンテキストで、読みやすく整形してください。"""
        system_prompt = ANALYZE_SYSTEM_PROMPT
    else:
        prompt = topic
        system_prompt = GENERAL_SYSTEM_PROMPT

    try:
        response, used_model = _create_chat_with_fallback(
            model_id=model_id,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
            temperature=0.5,
        )
    except Exception as e:
        err_str = str(e)
        if "429" in err_str or "rate" in err_str.lower():
            raise RuntimeError("APIのレート制限に達しました。しばらく待ってから再度お試しください。")
        raise RuntimeError(f"AI API エラー: {e}")

    content = response.choices[0].message.content
    if not content:
        raise RuntimeError(
            "AI API から空のレスポンスが返されました。"
            "Ollama等を使用している場合はモデルの応答が遅い可能性があります。"
            "しばらく待って再試行してください。"
        )

    return content, used_model


async def generate_presentation_content(req: GenerateRequest,
                                         model_id: str = "auto") -> tuple[PresentationData, str]:
    prompt = _build_prompt(req)
    try:
        response, used_model = _create_chat_with_fallback(
            model_id=model_id,
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
        raise RuntimeError(
            "AI API から空のレスポンスが返されました。"
            "Ollama等を使用している場合はモデルの応答が遅い可能性があります。"
            "スライド枚数を減らすか、しばらく待って再試行してください。"
        )

    try:
        return _parse_response(content), used_model
    except (json.JSONDecodeError, KeyError):
        extracted = _extract_first_json_object(content)
        if extracted:
            try:
                return _parse_response(extracted), used_model
            except (json.JSONDecodeError, KeyError) as e2:
                raise RuntimeError(f"AIの応答をパースできませんでした: {e2}")
        raise RuntimeError(
            "AIの応答をパースできませんでした: JSON形式で返っていない可能性があります。"
        )


# ---------------------------------------------------------------------------
# Excel generation
# ---------------------------------------------------------------------------

EXCEL_SYSTEM_PROMPT = """あなたはビジネスレポート作成の専門家です。
指示に従い、Excel形式のレポート構成を純粋なJSON形式のみで出力してください。
コードブロック記法（```）は絶対に使わないでください。
CSV/Excelデータの分析結果が提供された場合は、その実データを正確に反映したテーブルやチャートを作成してください。
chartのseriesのvaluesには、提供された実データの数値をそのまま使ってください。架空の数値は使わないでください。
tableのrowsの各要素では、数値はJSON数値型（引用符なし）で返してください。"""


def _build_excel_prompt(topic: str, source_context: str,
                        style: str, additional_instructions: str) -> str:
    style_map = {
        "business": "ビジネス向けのフォーマルなレポート",
        "casual": "読みやすくカジュアルなレポート",
        "academic": "学術的で詳細なレポート",
    }
    style_desc = style_map.get(style, style_map["business"])

    source_instruction = ""
    if source_context.strip():
        source_instruction = (
            "提供されたデータを正確に反映したテーブルとチャートを含めてください。"
            "chart_ready_data が含まれている場合はその実データを使ってください。"
        )

    return f"""以下の条件でExcelレポートの構成を生成してください。

【トピック】{topic}
【スタイル】{style_desc}
{f"【追加指示】{additional_instructions}" if additional_instructions else ""}
{source_instruction}

{source_context}

以下のJSON形式で出力してください。シートは1〜3枚が適切です。

{{
  "title": "レポートタイトル",
  "theme": "blue",
  "sheets": [
    {{
      "name": "シート名（最大31文字）",
      "sections": [
        {{"type": "title", "text": "レポートタイトル", "subtitle": "サブタイトル"}},
        {{"type": "heading", "text": "セクション見出し"}},
        {{"type": "text", "content": "テキスト内容（改行は\\nで）"}},
        {{"type": "kpi", "title": "主要指標", "items": [
          {{"label": "ラベル", "value": "¥1,234,567"}}
        ]}},
        {{"type": "table", "title": "表タイトル",
          "headers": ["列1", "列2", "列3"],
          "rows": [["文字列", 12345, 67.8]]
        }},
        {{"type": "chart", "chart_type": "bar",
          "title": "チャートタイトル",
          "categories": ["A", "B", "C"],
          "series": [{{"name": "系列名", "values": [100, 200, 300]}}]
        }}
      ]
    }}
  ]
}}

ルール:
- 最初のシートの最初のsectionは必ず type: "title" にすること
- theme は blue, green, orange, purple, gray のいずれか
- chart_type は bar, line, pie のいずれか
- table の rows の値は、数値はJSON数値型（引用符なし）で返すこと
- kpi の items は2〜5個が目安
- 読みやすく整理された構成にすること
- 純粋なJSONのみ出力すること
"""


def _parse_excel_response(text: str) -> dict:
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    return json.loads(cleaned.strip())


async def generate_excel_content(
    topic: str,
    csv_analysis: dict | None = None,
    pdf_text: str = "",
    additional_instructions: str = "",
    style: str = "business",
    model_id: str = "auto",
) -> tuple[dict, str]:
    source_context = _build_analysis_context(csv_analysis, pdf_text)
    prompt = _build_excel_prompt(topic, source_context,
                                 style, additional_instructions)

    try:
        response, used_model = _create_chat_with_fallback(
            model_id=model_id,
            messages=[
                {"role": "system", "content": EXCEL_SYSTEM_PROMPT},
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
        raise RuntimeError(
            "AI API から空のレスポンスが返されました。"
            "Ollama等を使用している場合はモデルの応答が遅い可能性があります。"
            "しばらく待って再試行してください。"
        )

    try:
        return _parse_excel_response(content), used_model
    except (json.JSONDecodeError, KeyError) as e:
        raise RuntimeError(f"AIの応答をパースできませんでした: {e}")
