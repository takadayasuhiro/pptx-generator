import fitz  # PyMuPDF

MAX_PDF_SIZE = 10 * 1024 * 1024  # 10MB


def validate_pdf(content: bytes) -> None:
    if len(content) > MAX_PDF_SIZE:
        size_mb = round(len(content) / (1024 * 1024), 1)
        raise ValueError(
            f"PDFファイルサイズが上限を超えています（{size_mb}MB）。最大10MBまでアップロードできます。"
        )


def extract_text(content: bytes) -> str:
    validate_pdf(content)

    try:
        doc = fitz.open(stream=content, filetype="pdf")
    except Exception:
        raise ValueError("PDFファイルの読み込みに失敗しました。ファイルが破損している可能性があります。")

    pages = []
    for page in doc:
        text = page.get_text("text")
        if text.strip():
            pages.append(text.strip())
    doc.close()

    if not pages:
        raise ValueError("PDFからテキストを抽出できませんでした。スキャン画像のみのPDFには対応していません。")

    full_text = "\n\n".join(pages)

    max_chars = 12000
    if len(full_text) > max_chars:
        full_text = full_text[:max_chars] + "\n\n...（以下省略）"

    return full_text
