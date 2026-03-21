import io
import pandas as pd
import numpy as np

MAX_CSV_SIZE = 5 * 1024 * 1024  # 5MB
MAX_CSV_ROWS = 1000


def validate_and_read(content: bytes) -> pd.DataFrame:
    if len(content) > MAX_CSV_SIZE:
        size_mb = round(len(content) / (1024 * 1024), 1)
        raise ValueError(
            f"CSVファイルサイズが上限を超えています（{size_mb}MB）。最大5MBまでアップロードできます。"
        )

    for encoding in ("utf-8", "utf-8-sig", "shift_jis", "cp932"):
        try:
            df = pd.read_csv(io.BytesIO(content), encoding=encoding)
            break
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    else:
        raise ValueError(
            "CSVファイルの読み込みに失敗しました。UTF-8 または Shift-JIS で保存してください。"
        )

    if len(df) > MAX_CSV_ROWS:
        raise ValueError(
            f"CSVファイルの行数が上限を超えています（{len(df):,}行）。"
            f"最大{MAX_CSV_ROWS:,}行までのファイルをアップロードしてください。"
        )

    if len(df) == 0:
        raise ValueError("CSVファイルにデータが含まれていません。")

    return df


def analyze(content: bytes) -> dict:
    df = validate_and_read(content)
    return analyze_dataframe(df)


def analyze_dataframe(df: pd.DataFrame) -> dict:
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    category_cols = df.select_dtypes(include="object").columns.tolist()
    datetime_cols = df.select_dtypes(include="datetime").columns.tolist()

    for col in category_cols[:]:
        try:
            parsed = pd.to_datetime(df[col], infer_datetime_format=True)
            df[col] = parsed
            datetime_cols.append(col)
            category_cols.remove(col)
        except (ValueError, TypeError):
            pass

    result = {
        "overview": {
            "rows": len(df),
            "columns": len(df.columns),
            "column_names": list(df.columns),
            "numeric_columns": numeric_cols,
            "category_columns": category_cols,
            "datetime_columns": datetime_cols,
        },
        "preview": _safe_preview(df, 5),
    }

    if numeric_cols:
        stats = df[numeric_cols].describe().round(2)
        result["statistics"] = {
            col: {k: _safe_val(v) for k, v in stats[col].items()}
            for col in stats.columns
        }

    if len(numeric_cols) >= 2:
        corr = df[numeric_cols].corr().round(3)
        strong = []
        for i, c1 in enumerate(corr.columns):
            for c2 in corr.columns[i + 1:]:
                val = corr.loc[c1, c2]
                if abs(val) >= 0.5 and not np.isnan(val):
                    strong.append({"col1": c1, "col2": c2, "value": float(val)})
        strong.sort(key=lambda x: abs(x["value"]), reverse=True)
        result["correlations"] = strong[:10]

    if category_cols:
        result["category_summary"] = {}
        for col in category_cols[:5]:
            vc = df[col].value_counts().head(10)
            result["category_summary"][col] = {
                str(k): int(v) for k, v in vc.items()
            }

    if numeric_cols:
        trends = {}
        for col in numeric_cols[:5]:
            series = df[col].dropna()
            if len(series) >= 3:
                first_half = series.iloc[: len(series) // 2].mean()
                second_half = series.iloc[len(series) // 2:].mean()
                if first_half != 0:
                    change_pct = round((second_half - first_half) / abs(first_half) * 100, 1)
                    if change_pct > 5:
                        trends[col] = f"上昇傾向 (+{change_pct}%)"
                    elif change_pct < -5:
                        trends[col] = f"下降傾向 ({change_pct}%)"
                    else:
                        trends[col] = "横ばい"
        if trends:
            result["trends"] = trends

    result["chart_ready_data"] = _prepare_chart_data(df, numeric_cols, category_cols)

    return result


def _prepare_chart_data(df: pd.DataFrame, numeric_cols: list, category_cols: list) -> list:
    charts = []

    if category_cols and numeric_cols:
        cat_col = category_cols[0]
        num_col = numeric_cols[0]
        grouped = df.groupby(cat_col)[num_col].sum().head(8)
        charts.append({
            "suggested_type": "bar",
            "title": f"{cat_col}別 {num_col}",
            "categories": [str(c) for c in grouped.index.tolist()],
            "series": [{"name": num_col, "values": [_safe_val(v) for v in grouped.values.tolist()]}],
        })

    if len(numeric_cols) >= 2 and len(df) >= 3:
        col = numeric_cols[0]
        values = df[col].dropna().tolist()
        if len(values) > 12:
            step = max(1, len(values) // 12)
            values = values[::step][:12]
        idx_labels = [str(i + 1) for i in range(len(values))]

        if category_cols:
            try:
                idx_labels = [str(v) for v in df[category_cols[0]].dropna().unique()[:len(values)]]
            except Exception:
                pass

        charts.append({
            "suggested_type": "line",
            "title": f"{col}の推移",
            "categories": idx_labels,
            "series": [{"name": col, "values": [_safe_val(v) for v in values]}],
        })

    if category_cols:
        cat_col = category_cols[0]
        vc = df[cat_col].value_counts().head(6)
        charts.append({
            "suggested_type": "pie",
            "title": f"{cat_col}の構成比",
            "categories": [str(c) for c in vc.index.tolist()],
            "series": [{"name": cat_col, "values": vc.values.tolist()}],
        })

    return charts


def _safe_preview(df: pd.DataFrame, n: int) -> list[dict]:
    preview = df.head(n)
    records = []
    for _, row in preview.iterrows():
        records.append({str(k): _safe_val(v) for k, v in row.items()})
    return records


def _safe_val(v):
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        if np.isnan(v) or np.isinf(v):
            return None
        return round(float(v), 4)
    if isinstance(v, (np.bool_,)):
        return bool(v)
    if pd.isna(v):
        return None
    return v
