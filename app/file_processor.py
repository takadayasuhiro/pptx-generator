import os
import logging

from fastapi import UploadFile

from app.pdf_service import extract_text as extract_pdf_text
from app.csv_service import analyze as analyze_csv
from app.excel_service import analyze as analyze_excel
from app.msg_service import extract_content as extract_msg_content
from app.pptx_input_service import extract_text_from_pptx
from app.image_analysis import analyze_image, IMAGE_EXTENSIONS

logger = logging.getLogger(__name__)

DOCUMENT_EXT = {".pdf"}
PRESENTATION_EXT = {".pptx"}
SPREADSHEET_EXT = {".csv", ".xlsx", ".xls"}
EMAIL_EXT = {".msg"}
VECTOR_EXT = {".svg"}

ALL_EXTENSIONS = (
    DOCUMENT_EXT | PRESENTATION_EXT | SPREADSHEET_EXT | EMAIL_EXT
    | IMAGE_EXTENSIONS | VECTOR_EXT
)

MAX_SINGLE_FILE_SIZE = 50 * 1024 * 1024  # 50MB

FILE_TYPE_LABELS = {
    ".pdf": "PDF",
    ".pptx": "PPTX",
    ".csv": "CSV",
    ".xlsx": "Excel",
    ".xls": "Excel",
    ".msg": "メール",
    ".svg": "SVG",
}


async def process_files(files: list[UploadFile],
                        model_id: str = "auto") -> dict:
    """Process multiple uploaded files and return aggregated results."""
    result: dict = {
        "texts": [],
        "csv_analyses": [],
        "image_descriptions": [],
        "file_names": [],
        "warnings": [],
    }

    for file in files:
        if not file.filename:
            continue

        ext = os.path.splitext(file.filename)[1].lower()

        if ext not in ALL_EXTENSIONS:
            result["warnings"].append(f"未対応のファイル形式です: {file.filename}")
            continue

        content = await file.read()

        if len(content) > MAX_SINGLE_FILE_SIZE:
            size_mb = round(len(content) / (1024 * 1024), 1)
            result["warnings"].append(
                f"{file.filename}: ファイルサイズが上限を超えています"
                f"（{size_mb}MB / 最大50MB）"
            )
            continue

        if len(content) == 0:
            result["warnings"].append(f"{file.filename}: ファイルが空です")
            continue

        result["file_names"].append(file.filename)

        try:
            if ext in DOCUMENT_EXT:
                text = extract_pdf_text(content)
                result["texts"].append(f"【PDF: {file.filename}】\n{text}")

            elif ext in PRESENTATION_EXT:
                text = extract_text_from_pptx(content)
                result["texts"].append(f"【PPTX: {file.filename}】\n{text}")

            elif ext == ".csv":
                analysis = analyze_csv(content)
                result["csv_analyses"].append({
                    "filename": file.filename, "analysis": analysis,
                })

            elif ext in (".xlsx", ".xls"):
                analysis = analyze_excel(content, file.filename)
                result["csv_analyses"].append({
                    "filename": file.filename, "analysis": analysis,
                })

            elif ext in EMAIL_EXT:
                text = extract_msg_content(content)
                result["texts"].append(f"【メール: {file.filename}】\n{text}")

            elif ext in IMAGE_EXTENSIONS:
                desc = await analyze_image(content, file.filename,
                                           model_id=model_id)
                result["image_descriptions"].append(
                    f"【画像: {file.filename}】\n{desc}"
                )

            elif ext in VECTOR_EXT:
                svg_text = content.decode("utf-8", errors="replace")
                if len(svg_text) > 5000:
                    svg_text = svg_text[:5000] + "\n...（以下省略）"
                result["texts"].append(
                    f"【SVG: {file.filename}】\n{svg_text}"
                )

        except ValueError as e:
            result["warnings"].append(f"{file.filename}: {str(e)}")
        except Exception as e:
            logger.exception(f"ファイル処理エラー: {file.filename}")
            result["warnings"].append(
                f"{file.filename}: 処理中にエラーが発生しました - {str(e)}"
            )

    return result
