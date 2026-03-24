import os
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from app.models import GenerateRequest
from app.ai_client import (
    generate_presentation_content, generate_analysis_text,
    generate_excel_content,
)
from app.pptx_builder import build_pptx, build_summary_data, get_available_pptx_templates
from app.excel_builder import build_excel
from app.file_processor import process_files
from app.model_registry import get_available_models, clear_ollama_cache

app = FastAPI(title="PPTX Auto Generator", version="4.0.0")

STATIC_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "static")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/models")
async def list_models(refresh: bool = False):
    if refresh:
        clear_ollama_cache()
    models = get_available_models()
    safe = [
        {
            "id": m["id"],
            "label": m["label"],
            "provider": m["provider"],
            "available": m["available"],
            "supports_vision": m.get("supports_vision", False),
            "description": m.get("description", ""),
        }
        for m in models
    ]
    return {"models": safe}


@app.get("/pptx-templates")
async def list_pptx_templates():
    templates = get_available_pptx_templates()
    return {"templates": templates}


@app.post("/generate")
async def generate_pptx(
    topic: str = Form(...),
    num_slides: int = Form(8),
    style: str = Form("business"),
    language: str = Form("ja"),
    additional_instructions: str = Form(""),
    include_summary: bool = Form(False),
    model: str = Form("auto"),
    template: str = Form(""),
    files: list[UploadFile] = File(default=[]),
):
    if not template:
        raise HTTPException(status_code=400, detail="PPTXテンプレートを選択してください。")

    processed = await process_files(files, model_id=model)

    all_text_parts = list(processed["texts"])
    if processed["image_descriptions"]:
        all_text_parts.append(
            "■ 画像分析結果:\n" + "\n\n".join(processed["image_descriptions"])
        )
    combined_text = "\n\n".join(all_text_parts)

    csv_analysis = None
    if processed["csv_analyses"]:
        csv_analysis = processed["csv_analyses"][0]["analysis"]

    has_data = bool(combined_text) or bool(csv_analysis)

    summary_data = None
    if include_summary and has_data:
        summary_data = build_summary_data(
            csv_analysis=csv_analysis, pdf_text=combined_text,
        )

    req = GenerateRequest(
        topic=topic,
        num_slides=num_slides,
        style=style,
        language=language,
        additional_instructions=additional_instructions,
        pdf_text=combined_text,
        csv_analysis=csv_analysis,
    )

    try:
        presentation_data, used_model = await generate_presentation_content(req, model_id=model)
    except RuntimeError as e:
        raise HTTPException(status_code=502, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"AI API エラー: {str(e)}")

    try:
        output_path = build_pptx(
            presentation_data, presentation_data.title,
            summary_data=summary_data,
            template_id=template,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPTX生成エラー: {str(e)}")

    response = FileResponse(
        path=output_path,
        filename=os.path.basename(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

    if processed["warnings"]:
        response.headers["X-Processing-Warnings"] = " | ".join(processed["warnings"])
    if used_model:
        response.headers["X-AI-Model-Used"] = used_model

    return response


@app.post("/generate-excel")
async def generate_excel_report(
    topic: str = Form(...),
    style: str = Form("business"),
    additional_instructions: str = Form(""),
    model: str = Form("auto"),
    files: list[UploadFile] = File(default=[]),
):
    processed = await process_files(files, model_id=model)

    all_text_parts = list(processed["texts"])
    if processed["image_descriptions"]:
        all_text_parts.append(
            "■ 画像分析結果:\n" + "\n\n".join(processed["image_descriptions"])
        )
    combined_text = "\n\n".join(all_text_parts)

    csv_analysis = None
    if processed["csv_analyses"]:
        csv_analysis = processed["csv_analyses"][0]["analysis"]

    try:
        excel_data, used_model = await generate_excel_content(
            topic=topic,
            csv_analysis=csv_analysis,
            pdf_text=combined_text,
            additional_instructions=additional_instructions,
            style=style,
            model_id=model,
        )
    except RuntimeError as e:
        raise HTTPException(status_code=502, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"AI API エラー: {str(e)}")

    try:
        output_path = build_excel(excel_data, excel_data.get("title", "report"))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel生成エラー: {str(e)}")

    response = FileResponse(
        path=output_path,
        filename=os.path.basename(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if processed["warnings"]:
        response.headers["X-Processing-Warnings"] = " | ".join(processed["warnings"])
    if used_model:
        response.headers["X-AI-Model-Used"] = used_model

    return response


@app.post("/analyze")
async def analyze_file(
    topic: str = Form(""),
    model: str = Form("auto"),
    files: list[UploadFile] = File(default=[]),
):
    processed = await process_files(files, model_id=model)

    all_text_parts = list(processed["texts"])
    if processed["image_descriptions"]:
        all_text_parts.append(
            "■ 画像分析結果:\n" + "\n\n".join(processed["image_descriptions"])
        )
    combined_text = "\n\n".join(all_text_parts)

    csv_analysis = None
    if processed["csv_analyses"]:
        csv_analysis = processed["csv_analyses"][0]["analysis"]

    has_file_data = bool(combined_text) or bool(csv_analysis)
    has_topic = bool(topic.strip())

    if not has_file_data and not has_topic:
        raise HTTPException(
            status_code=400,
            detail="テキストを入力するか、ファイルを添付してください。",
        )

    try:
        analysis, used_model = await generate_analysis_text(
            csv_analysis=csv_analysis,
            pdf_text=combined_text,
            topic=topic,
            model_id=model,
        )
    except RuntimeError as e:
        raise HTTPException(status_code=502, detail=str(e))

    result: dict = {"analysis": analysis}
    if processed["warnings"]:
        result["warnings"] = processed["warnings"]

    response = JSONResponse(content=result)
    if used_model:
        response.headers["X-AI-Model-Used"] = used_model
    return response
