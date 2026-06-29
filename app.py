from __future__ import annotations

import json
import math
import os
import re
import uuid
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "outputs"
UPLOAD_DIR = ROOT / "uploads"
MAX_PREVIEW_ROWS = 25
EXCEL_CELL_CHAR_LIMIT = 32767
TRUNCATED_SUFFIX = "\n...[内容过长，已截断]"


def is_object_id_column(column):
    text = str(column).strip().lower().replace("_", "").replace(" ", "")
    return text in {"objectid", "对象id"} or text.endswith("objectid") or text.endswith("对象id")


def stringify_identifier(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        if value.is_integer():
            return f"{value:.0f}"
        return str(value)
    return str(value).strip()


def clean_cell(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def json_safe(value):
    if pd.isna(value):
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def records_for_json(df):
    records = df.to_dict(orient="records")
    return [{key: json_safe(value) for key, value in row.items()} for row in records]


def read_sheet(path, sheet_name=""):
    excel = pd.ExcelFile(path)
    sheet_names = excel.sheet_names
    selected = sheet_name if sheet_name in sheet_names else sheet_names[0]
    return pd.read_excel(path, sheet_name=selected, dtype=str), sheet_names, selected


def parse_jsonish(raw):
    if raw is None or pd.isna(raw):
        return {}, "empty"
    text = str(raw).strip()
    if not text:
        return {}, "empty"
    try:
        parsed = json.loads(text)
        if isinstance(parsed, dict):
            return parsed, ""
        return {"parsed_value": parsed}, ""
    except Exception:
        pass

    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        try:
            parsed = json.loads(text[start : end + 1])
            if isinstance(parsed, dict):
                return parsed, "extracted"
            return {"parsed_value": parsed}, "extracted"
        except Exception:
            return {}, "invalid_json"
    return {}, "invalid_json"


def infer_json_columns(df):
    scores = []
    for column in df.columns:
        sample = df[column].dropna().astype(str).head(60)
        if sample.empty:
            continue
        hits = 0
        for value in sample:
            if "{" not in value or "}" not in value:
                continue
            parsed, status = parse_jsonish(value)
            if parsed and status in ("", "extracted") and "parsed_value" not in parsed:
                hits += 1
        if hits:
            scores.append({"column": str(column), "hits": hits, "sampled": int(len(sample))})
    return sorted(scores, key=lambda item: item["hits"], reverse=True)


def flatten_json_column(df, json_column):
    parsed_rows = []
    all_keys = []
    list_lengths = {}
    statuses = []

    for raw in df[json_column].tolist():
        parsed, status = parse_jsonish(raw)
        parsed_rows.append(parsed)
        statuses.append(status)
        for key, value in parsed.items():
            if key not in all_keys:
                all_keys.append(key)
            if isinstance(value, list):
                list_lengths[key] = max(list_lengths.get(key, 0), len(value))

    result = df.copy()
    output_keys = []
    for key in all_keys:
        if key in list_lengths:
            for index in range(list_lengths[key]):
                column_name = f"{key}_{index + 1}"
                output_keys.append(column_name)
                values = []
                for row in parsed_rows:
                    value = row.get(key, None)
                    if isinstance(value, list):
                        item = value[index] if index < len(value) else None
                    else:
                        item = value if index == 0 else None
                    if isinstance(item, (dict, list)):
                        item = json.dumps(item, ensure_ascii=False)
                    values.append(item)
                result[column_name] = values
        else:
            output_keys.append(key)
            values = [row.get(key, None) for row in parsed_rows]
            result[key] = values

    result["_json_parse_status"] = statuses
    return result, output_keys


def split_positive_labels(positive_label):
    return [
        item.strip()
        for item in re.split(r"[,，\n\r;；]+", positive_label or "")
        if item.strip()
    ]


def normalize_positive(value, positive_labels):
    if value is None or pd.isna(value):
        return False
    labels = positive_labels if isinstance(positive_labels, list) else split_positive_labels(positive_labels)
    if labels:
        return str(value).strip() in labels
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) > 0
    text = str(value).strip().lower()
    return text in {"1", "true", "yes", "y", "是", "正例", "命中", "正确"}


def to_number(series):
    return pd.to_numeric(series, errors="coerce")


def build_stats(df, score_column, truth_column, positive_label="", direction="high"):
    scores = to_number(df[score_column])
    positive_labels = split_positive_labels(positive_label)
    truths = df[truth_column].apply(lambda value: normalize_positive(value, positive_labels))
    valid = scores.notna() & df[truth_column].notna()
    total_valid = int(valid.sum())
    total_positive = int((truths & valid).sum())

    work = pd.DataFrame({"score": scores, "truth": truths, "valid": valid})
    work = work[work["valid"]].copy()
    if work.empty:
        return {
            "summary": {
                "total_valid": 0,
                "total_positive": 0,
                "score_column": score_column,
                "truth_column": truth_column,
            },
            "bucket_stats": [],
            "threshold_stats": [],
        }

    score_edges = [round(i / 100, 2) for i in range(0, 101)]
    bins = [float("-inf"), *score_edges, float("inf")]
    labels = ["<0", *[f"{score_edges[i]:.2f}-{score_edges[i + 1]:.2f}" for i in range(len(score_edges) - 1)], ">1"]
    work["bucket"] = pd.cut(work["score"], bins=bins, labels=labels, right=False)

    bucket_stats = []
    for label in reversed(labels):
        part = work[work["bucket"].astype(str) == label]
        count = int(len(part))
        positives = int(part["truth"].sum()) if count else 0
        bucket_stats.append(
            {
                "score_layer": label,
                "sample_count": count,
                "positive_count": positives,
                "positive_rate": round(positives / count, 4) if count else 0,
                "recall_share": round(positives / total_positive, 4) if total_positive else 0,
            }
        )

    thresholds = [round(i / 100, 2) for i in range(100, -1, -1)]
    threshold_stats = []
    for threshold in thresholds:
        if direction == "low":
            pred = work["score"] <= threshold
            threshold_label = f"<= {threshold:.2f}"
        else:
            pred = work["score"] >= threshold
            threshold_label = f">= {threshold:.2f}"
        predicted = int(pred.sum())
        true_positive = int((pred & work["truth"]).sum())
        false_positive = int((pred & ~work["truth"]).sum())
        false_negative = int((~pred & work["truth"]).sum())
        precision = true_positive / predicted if predicted else 0
        recall = true_positive / total_positive if total_positive else 0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0
        threshold_stats.append(
            {
                "threshold": threshold_label,
                "predicted_positive": predicted,
                "true_positive": true_positive,
                "false_positive": false_positive,
                "false_negative": false_negative,
                "precision": round(precision, 4),
                "recall": round(recall, 4),
                "f1": round(f1, 4),
            }
        )

    return {
        "summary": {
            "total_valid": total_valid,
            "total_positive": total_positive,
            "score_column": score_column,
            "truth_column": truth_column,
            "direction": direction,
            "positive_label": positive_label,
            "positive_labels": "，".join(positive_labels),
        },
        "bucket_stats": bucket_stats,
        "threshold_stats": threshold_stats,
    }


def format_prob_columns(df):
    result = df.copy()
    for column in result.columns:
        if "prob" in str(column).lower():
            converted = pd.to_numeric(result[column], errors="coerce")
            if converted.notna().any():
                result[column] = converted
    return result


def format_identifier_columns(df):
    result = df.copy()
    for column in result.columns:
        if is_object_id_column(column):
            result[column] = result[column].apply(stringify_identifier)
    return result


def protect_excel_text_limits(df):
    result = df.copy()
    truncated_columns_by_row = []
    available = EXCEL_CELL_CHAR_LIMIT - len(TRUNCATED_SUFFIX)

    for index, row in result.iterrows():
        truncated_columns = []
        for column, value in row.items():
            if pd.isna(value) or not isinstance(value, str):
                continue
            if len(value) > EXCEL_CELL_CHAR_LIMIT:
                result.at[index, column] = value[:available] + TRUNCATED_SUFFIX
                truncated_columns.append(str(column))
        truncated_columns_by_row.append("，".join(truncated_columns))

    if any(truncated_columns_by_row):
        result["_excel_truncated_columns"] = truncated_columns_by_row
    return result


def prepare_detail_dataframe(df):
    return protect_excel_text_limits(format_prob_columns(format_identifier_columns(df)))


def apply_workbook_styles(workbook, detail_sheet_names=None):
    detail_sheet_names = set(detail_sheet_names or [])
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.style = "Headline 4"
        for column_cells in sheet.columns:
            header = str(column_cells[0].value or "")
            width = min(max(len(header) + 4, 10), 42)
            sheet.column_dimensions[column_cells[0].column_letter].width = width

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                header = str(sheet.cell(row=1, column=cell.column).value or "")
                normalized_header = header.lower()
                if sheet.title in detail_sheet_names and is_object_id_column(header):
                    if cell.row > 1 and cell.value is not None:
                        cell.value = str(cell.value)
                    cell.number_format = "@"
                elif "prob" in normalized_header or normalized_header in {"precision", "recall", "f1", "positive_rate", "recall_share"}:
                    cell.number_format = "0.00"


def save_flat_workbook(df, file_id):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"model_eval_flattened_{file_id}.xlsx"
    detail = prepare_detail_dataframe(df)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        detail.to_excel(writer, index=False, sheet_name="拆分明细")
        apply_workbook_styles(writer.book, detail_sheet_names={"拆分明细"})
    return output_path


def save_workbook(df, stats, file_id):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"model_eval_result_{file_id}.xlsx"
    detail = prepare_detail_dataframe(df)
    bucket_df = pd.DataFrame(stats["bucket_stats"])
    threshold_df = pd.DataFrame(stats["threshold_stats"])
    summary_df = pd.DataFrame([stats["summary"]])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        detail.to_excel(writer, index=False, sheet_name="拆分明细")
        summary_df.to_excel(writer, index=False, sheet_name="统计概览")
        bucket_df.to_excel(writer, index=False, sheet_name="分层分布")
        threshold_df.to_excel(writer, index=False, sheet_name="阈值准召")

        apply_workbook_styles(writer.book, detail_sheet_names={"拆分明细"})
    return output_path


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def send_json(self, payload, status=HTTPStatus.OK):
        body = json.dumps(payload, ensure_ascii=False, default=json_safe).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            content = (ROOT / "index.html").read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return
        if parsed.path.startswith("/download/"):
            file_name = Path(parsed.path).name
            path = OUTPUT_DIR / file_name
            if not path.exists():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            content = path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{file_name}"')
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        if self.path == "/api/preview":
            self.handle_preview()
            return
        if self.path == "/api/flatten":
            self.handle_flatten()
            return
        if self.path == "/api/process":
            self.handle_process()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def read_multipart(self):
        content_type = self.headers.get("Content-Type", "")
        match = re.search(r"boundary=(?P<boundary>[^;]+)", content_type)
        if not match:
            raise ValueError("请求格式不正确。")
        boundary = match.group("boundary").strip('"').encode("utf-8")
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        fields = {}
        marker = b"--" + boundary

        for part in body.split(marker):
            part = part.strip()
            if not part or part == b"--":
                continue
            if part.endswith(b"--"):
                part = part[:-2].strip()
            header_bytes, _, value = part.partition(b"\r\n\r\n")
            headers = header_bytes.decode("utf-8", errors="replace")
            name_match = re.search(r'name="([^"]+)"', headers)
            if not name_match:
                continue
            name = name_match.group(1)
            if value.endswith(b"\r\n"):
                value = value[:-2]
            if "filename=" in headers:
                fields[name] = value
            else:
                fields[name] = value.decode("utf-8", errors="replace")
        return fields

    def handle_preview(self):
        try:
            form = self.read_multipart()
            file_id = form.get("file_id", "")
            sheet_name = form.get("sheet_name", "")
            if "file" in form:
                file_bytes = form["file"]
                file_id = uuid.uuid4().hex
                UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
                path = UPLOAD_DIR / f"{file_id}.xlsx"
                path.write_bytes(file_bytes)
            else:
                path = UPLOAD_DIR / f"{file_id}.xlsx"
                if not file_id or not path.exists():
                    raise ValueError("请先上传 Excel 文件。")

            df, sheet_names, selected_sheet = read_sheet(path, sheet_name)
            columns = [str(column) for column in df.columns]
            preview = records_for_json(df.head(MAX_PREVIEW_ROWS))
            self.send_json(
                {
                    "file_id": file_id,
                    "sheet_names": sheet_names,
                    "selected_sheet": selected_sheet,
                    "rows": int(len(df)),
                    "columns": columns,
                    "json_candidates": infer_json_columns(df),
                    "preview": preview,
                }
            )
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def handle_flatten(self):
        try:
            form = self.read_multipart()
            file_id = form.get("file_id")
            sheet_name = form.get("sheet_name", "")
            json_column = form.get("json_column")
            if not all([file_id, json_column]):
                raise ValueError("请先选择 JSON 列。")

            path = UPLOAD_DIR / f"{file_id}.xlsx"
            if not path.exists():
                raise ValueError("上传文件已失效，请重新上传。")
            df, sheet_names, selected_sheet = read_sheet(path, sheet_name)
            if json_column not in df.columns:
                raise ValueError("选择的 JSON 列不存在，请重新选择。")

            flat_df, parsed_keys = flatten_json_column(df, json_column)
            output = save_flat_workbook(flat_df, file_id)
            preview = records_for_json(flat_df.head(MAX_PREVIEW_ROWS))
            self.send_json(
                {
                    "parsed_columns": parsed_keys,
                    "columns": [str(column) for column in flat_df.columns],
                    "sheet_names": sheet_names,
                    "selected_sheet": selected_sheet,
                    "preview": preview,
                    "download_url": f"/download/{output.name}",
                }
            )
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def handle_process(self):
        try:
            form = self.read_multipart()
            file_id = form.get("file_id")
            sheet_name = form.get("sheet_name", "")
            json_column = form.get("json_column")
            score_column = form.get("score_column")
            truth_column = form.get("truth_column")
            positive_label = form.get("positive_label", "")
            direction = form.get("direction", "high")
            if not all([file_id, json_column, score_column, truth_column]):
                raise ValueError("请先选择 JSON 列、模型分列和正确结果列。")

            path = UPLOAD_DIR / f"{file_id}.xlsx"
            if not path.exists():
                raise ValueError("上传文件已失效，请重新上传。")
            df, sheet_names, selected_sheet = read_sheet(path, sheet_name)
            flat_df, parsed_keys = flatten_json_column(df, json_column)
            if score_column not in flat_df.columns or truth_column not in flat_df.columns:
                raise ValueError("选择的列不存在，请重新选择。")
            stats = build_stats(flat_df, score_column, truth_column, positive_label, direction)
            output = save_workbook(flat_df, stats, file_id)
            preview = records_for_json(flat_df.head(MAX_PREVIEW_ROWS))
            self.send_json(
                {
                    "parsed_columns": parsed_keys,
                    "columns": [str(column) for column in flat_df.columns],
                    "sheet_names": sheet_names,
                    "selected_sheet": selected_sheet,
                    "preview": preview,
                    "stats": stats,
                    "download_url": f"/download/{output.name}",
                }
            )
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)


def main():
    port = int(os.environ.get("PORT", "8765"))
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"模型评测分析工具已启动：http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
