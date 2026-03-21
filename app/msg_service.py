import os
import tempfile

MAX_MSG_SIZE = 50 * 1024 * 1024  # 50MB


def extract_content(content: bytes) -> str:
    if len(content) > MAX_MSG_SIZE:
        raise ValueError("MSGファイルサイズが上限（50MB）を超えています。")

    try:
        import extract_msg
    except ImportError:
        raise ValueError(
            "MSGファイルの処理に必要なライブラリ(extract-msg)がインストールされていません。"
        )

    tmp = tempfile.NamedTemporaryFile(suffix=".msg", delete=False)
    try:
        tmp.write(content)
        tmp.close()

        msg = extract_msg.Message(tmp.name)
        parts: list[str] = []

        if msg.subject:
            parts.append(f"件名: {msg.subject}")
        if msg.sender:
            parts.append(f"差出人: {msg.sender}")
        if msg.date:
            parts.append(f"日時: {msg.date}")

        parts.append("")

        if msg.body:
            body = msg.body
            if len(body) > 10000:
                body = body[:10000] + "\n\n...（以下省略）"
            parts.append(f"本文:\n{body}")

        if msg.attachments:
            att_names = [
                getattr(a, "longFilename", None)
                or getattr(a, "shortFilename", None)
                or "不明"
                for a in msg.attachments
            ]
            parts.append(f"\n添付ファイル: {', '.join(att_names)}")

        msg.close()
        result = "\n".join(parts)

        if not result.strip():
            raise ValueError("MSGファイルからコンテンツを抽出できませんでした。")

        return result
    finally:
        os.unlink(tmp.name)
