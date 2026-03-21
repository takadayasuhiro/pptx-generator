import io
import os

import pandas as pd

from app.csv_service import analyze_dataframe

MAX_EXCEL_SIZE = 20 * 1024 * 1024  # 20MB
MAX_EXCEL_ROWS = 1000


def validate_and_read(content: bytes, filename: str = "") -> pd.DataFrame:
    if len(content) > MAX_EXCEL_SIZE:
        size_mb = round(len(content) / (1024 * 1024), 1)
        raise ValueError(
            f"Excelファイルサイズが上限を超えています（{size_mb}MB）。最大20MBまでアップロードできます。"
        )

    ext = os.path.splitext(filename)[1].lower() if filename else ".xlsx"

    try:
        if ext == ".xls":
            try:
                df = pd.read_excel(io.BytesIO(content), engine="xlrd")
            except ImportError:
                raise ValueError(
                    ".xls 形式を読み込むには xlrd が必要です。"
                    ".xlsx 形式に変換してからアップロードしてください。"
                )
        else:
            df = pd.read_excel(io.BytesIO(content), engine="openpyxl")
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"Excelファイルの読み込みに失敗しました: {e}")

    if len(df) > MAX_EXCEL_ROWS:
        raise ValueError(
            f"Excelファイルの行数が上限を超えています（{len(df):,}行）。"
            f"最大{MAX_EXCEL_ROWS:,}行までのファイルをアップロードしてください。"
        )

    if len(df) == 0:
        raise ValueError("Excelファイルにデータが含まれていません。")

    return df


def analyze(content: bytes, filename: str = "") -> dict:
    df = validate_and_read(content, filename)
    return analyze_dataframe(df)
