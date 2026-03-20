import os
import json
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
from app.models import GenerateRequest
from app.ai_client import generate_presentation_content
from app.pptx_builder import build_pptx, build_summary_data
from app.pdf_service import extract_text as extract_pdf_text
from app.csv_service import analyze as analyze_csv

app = FastAPI(title="PPTX Auto Generator", version="2.1.0")

STATIC_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "static")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

ALLOWED_EXTENSIONS = {".pdf", ".csv"}


@app.get("/", response_class=HTMLResponse)
async def index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.post("/generate")
async def generate_pptx(
    topic: str = Form(...),
    num_slides: int = Form(8),
    style: str = Form("business"),
    language: str = Form("ja"),
    additional_instructions: str = Form(""),
    include_summary: bool = Form(False),
    pdf_file: UploadFile | None = File(None),
    csv_file: UploadFile | None = File(None),
):
    pdf_text = ""
    csv_analysis = None

    if pdf_file and pdf_file.filename:
        ext = os.path.splitext(pdf_file.filename)[1].lower()
        if ext != ".pdf":
            raise HTTPException(status_code=400, detail="PDFファイル以外はアップロードできません。")
        try:
            content = await pdf_file.read()
            pdf_text = extract_pdf_text(content)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))

    if csv_file and csv_file.filename:
        ext = os.path.splitext(csv_file.filename)[1].lower()
        if ext != ".csv":
            raise HTTPException(status_code=400, detail="CSVファイル以外はアップロードできません。")
        try:
            content = await csv_file.read()
            csv_analysis = analyze_csv(content)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))

    summary_data = None
    has_files = bool(pdf_text) or bool(csv_analysis)
    if include_summary and has_files:
        summary_data = build_summary_data(csv_analysis=csv_analysis, pdf_text=pdf_text)

    req = GenerateRequest(
        topic=topic,
        num_slides=num_slides,
        style=style,
        language=language,
        additional_instructions=additional_instructions,
        pdf_text=pdf_text,
        csv_analysis=csv_analysis,
    )

    try:
        presentation_data = await generate_presentation_content(req)
    except RuntimeError as e:
        raise HTTPException(status_code=502, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"AI API エラー: {str(e)}")

    try:
        output_path = build_pptx(presentation_data, presentation_data.title,
                                 summary_data=summary_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPTX生成エラー: {str(e)}")

    return FileResponse(
        path=output_path,
        filename=os.path.basename(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
