from __future__ import annotations

import io
import re
import calendar
import base64
import zipfile
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st


RED = "#c43d47"
GREEN = "#187348"
BLUE = "#3568b8"
TARGET = "#718096"
GITHUB_REPOSITORY = "xiangjianli155-byte/kpi-assessment"
GITHUB_BRANCH = "main"
SHARED_UPLOAD_TYPES = {".xlsx", ".xlsm", ".xls", ".csv"}
CANONICAL_PLANNERS = {
    "li chao": "Li Chao",
    "she xiaolei": "She xiaolei",
    "xin mei": "Xin mei",
    "jessie jin": "Jessie Jin",
    "zhang mingshuai": "Zhang Mingshuai",
    "li xiang": "Li Xiang",
    "lixiang": "Li Xiang",
}


st.set_page_config(page_title="KPI 考核", page_icon="📊", layout="wide")
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem}
    .kpi-card {border:1px solid #e4e9ef;border-radius:12px;padding:15px 16px;
      background:#fff;min-height:105px}
    .kpi-label {color:#667386;font-size:13px}
    .kpi-value {font-size:25px;font-weight:650;margin-top:8px;color:#202938}
    .kpi-foot {color:#667386;font-size:12px;margin-top:4px}
    .pill {display:inline-block;border-radius:99px;padding:4px 10px;font-size:12px;font-weight:650}
    .pill-pass {background:#eaf7f0;color:#187348}
    .pill-fail {background:#fff0f0;color:#b92f3a}
    .pill-wait {background:#fff7e6;color:#966512}
    </style>
    """,
    unsafe_allow_html=True,
)


def is_missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def clean_text(value: Any) -> str | None:
    if is_missing(value):
        return None
    text = str(value).strip()
    if not text or text.casefold() in {"nan", "none", "nat", "#n/a", "#value!"}:
        return None
    return text


def normalize_planner(value: Any) -> str | None:
    text = clean_text(value)
    if not text:
        return None
    if "/" in text:
        return "多人负责"
    return CANONICAL_PLANNERS.get(text.casefold(), text)


def normalize_whs(value: Any) -> str | None:
    text = clean_text(value)
    if not text:
        return None
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except ValueError:
        pass
    return text


def number_or_none(value: Any) -> float | None:
    if is_missing(value):
        return None
    try:
        number = pd.to_numeric(value, errors="coerce")
        return None if pd.isna(number) else float(number)
    except (TypeError, ValueError):
        return None


def excel_column_letter(index: int) -> str:
    result = ""
    while index >= 0:
        index, remainder = divmod(index, 26)
        result = chr(65 + remainder) + result
        index -= 1
    return result


def default_year(filename: str) -> int:
    found = re.search(r"\b(20\d{2})\b", filename)
    return int(found.group(1)) if found else date.today().year


def parse_period(value: Any, year: int) -> pd.Timestamp | None:
    if is_missing(value):
        return None
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(value).normalize()

    text = str(value).strip()
    if not text:
        return None

    week_match = re.search(r"(20\d{2})\s*[-_/]?\s*[Ww](\d{1,2})\b", text)
    if week_match:
        try:
            return pd.Timestamp(datetime.fromisocalendar(int(week_match.group(1)), int(week_match.group(2)), 1))
        except ValueError:
            return None

    full_date = re.search(r"(20\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})", text)
    if full_date:
        try:
            return pd.Timestamp(int(full_date.group(1)), int(full_date.group(2)), int(full_date.group(3)))
        except ValueError:
            return None

    month_match = re.search(r"(?<!\d)(\d{1,2})\s*月", text)
    if month_match:
        month = int(month_match.group(1))
        if 1 <= month <= 12:
            return pd.Timestamp(year, month, calendar.monthrange(year, month)[1])

    quarter_match = re.search(r"Q\s*([1-4])", text, flags=re.IGNORECASE)
    if quarter_match:
        month = int(quarter_match.group(1)) * 3
        return pd.Timestamp(year, month, calendar.monthrange(year, month)[1])

    # Date-like strings in file headers, such as 2026-11-01 or 11/01/2026.
    parsed = pd.to_datetime(text, errors="coerce")
    if not pd.isna(parsed):
        return pd.Timestamp(parsed).normalize()
    return None


def header_text(value: Any) -> str:
    return clean_text(value).casefold() if clean_text(value) else ""


def guess_header_row(raw: pd.DataFrame) -> int:
    best_row, best_score = 0, -1
    for row_index in range(min(15, len(raw))):
        cells = [header_text(value) for value in raw.iloc[row_index].tolist()]
        score = 0
        score += 3 if any(value == "whs" or "whs" in value for value in cells) else 0
        score += 3 if any(value == "planner" or "planner" in value for value in cells) else 0
        score += 2 if any("target" in value or "目标" in value for value in cells) else 0
        score += 1 if any("actual" in value or value == "act" or "实际" in value for value in cells) else 0
        if score > best_score:
            best_row, best_score = row_index, score
    return best_row


def find_header_column(headers: list[Any], kind: str) -> int | None:
    labels = [header_text(value) for value in headers]
    if kind == "planner":
        exact = [i for i, label in enumerate(labels) if label == "planner"]
        return exact[-1] if exact else next((i for i, label in enumerate(labels) if "planner" in label), None)
    if kind == "whs":
        return next((i for i, label in enumerate(labels) if label == "whs"), None)
    if kind == "target":
        exact = [i for i, label in enumerate(labels) if label.startswith("target") or "目标" in label]
        return exact[-1] if exact else None
    if kind == "actual":
        return next(
            (i for i, label in enumerate(labels) if label in {"actual", "act", "实际", "实际值", "实绩"}),
            next((i for i, label in enumerate(labels) if "actual" in label or "实际" in label), None),
        )
    if kind == "date":
        return next(
            (
                i
                for i, label in enumerate(labels)
                if any(word in label for word in ("date", "week", "period", "时间", "日期", "周次", "月份"))
            ),
            None,
        )
    return None


def column_has_number(raw: pd.DataFrame, row_start: int, column: int) -> bool:
    values = pd.to_numeric(raw.iloc[row_start:, column], errors="coerce")
    return bool(values.notna().any())


def detect_schema(raw: pd.DataFrame, filename: str) -> dict[str, Any] | None:
    if raw.empty:
        return None
    year = default_year(filename)
    header_row = guess_header_row(raw)
    headers = raw.iloc[header_row].tolist()
    planner_col = find_header_column(headers, "planner")
    whs_col = find_header_column(headers, "whs")
    target_col = find_header_column(headers, "target")
    if planner_col is None or whs_col is None:
        return None

    date_columns: list[tuple[int, pd.Timestamp]] = []
    for col, label in enumerate(headers):
        period = parse_period(label, year)
        if period is not None and column_has_number(raw, header_row + 1, col):
            date_columns.append((col, period))
    if date_columns:
        return {
            "layout": "wide",
            "header_row": header_row,
            "period_row": None,
            "planner_col": planner_col,
            "whs_col": whs_col,
            "target_col": target_col,
            "actual_cols": [col for col, _ in date_columns],
            "periods": dict(date_columns),
            "description": f"宽表：识别到 {len(date_columns)} 个日期列",
        }

    period_row = header_row - 1 if header_row > 0 else None
    act_columns: list[tuple[int, pd.Timestamp]] = []
    if period_row is not None:
        parent_labels = raw.iloc[period_row].tolist()
        for col, label in enumerate(headers):
            if header_text(label) in {"act", "actual", "实际", "实际值", "实绩"}:
                period = parse_period(parent_labels[col], year)
                if period is not None and column_has_number(raw, header_row + 1, col):
                    act_columns.append((col, period))
    if act_columns:
        return {
            "layout": "wide",
            "header_row": header_row,
            "period_row": period_row,
            "planner_col": planner_col,
            "whs_col": whs_col,
            "target_col": target_col,
            "actual_cols": [col for col, _ in act_columns],
            "periods": dict(act_columns),
            "description": f"宽表：识别到 {len(act_columns)} 个期间实绩列",
        }

    date_col = find_header_column(headers, "date")
    actual_col = find_header_column(headers, "actual")
    if date_col is not None and actual_col is not None:
        return {
            "layout": "long",
            "header_row": header_row,
            "planner_col": planner_col,
            "whs_col": whs_col,
            "target_col": target_col,
            "date_col": date_col,
            "actual_col": actual_col,
            "description": "长表：识别到日期、Planner、WHS 和实绩列",
        }
    return None


def column_options(raw: pd.DataFrame, header_row: int) -> list[str]:
    options = []
    headers = raw.iloc[header_row].tolist()
    for index, value in enumerate(headers):
        label = clean_text(value) or "空列"
        options.append(f"{excel_column_letter(index)} · {label}")
    return options


def column_select(
    label: str,
    options: list[str],
    guessed_index: int | None,
    key: str,
    allow_none: bool = False,
) -> int | None:
    display_options = (["（无）"] if allow_none else []) + options
    if allow_none and guessed_index is None:
        default_index = 0
    else:
        offset = 1 if allow_none else 0
        default_index = (guessed_index if guessed_index is not None else 0) + offset
        default_index = max(0, min(default_index, len(display_options) - 1))
    choice = st.selectbox(label, display_options, index=default_index, key=key)
    if allow_none and choice == "（无）":
        return None
    return display_options.index(choice) - (1 if allow_none else 0)


def read_raw_sheet(file_bytes: bytes, filename: str, sheet_name: str) -> pd.DataFrame:
    if filename.casefold().endswith(".csv"):
        for encoding in ("utf-8-sig", "gb18030", "utf-8"):
            try:
                return pd.read_csv(io.BytesIO(file_bytes), header=None, encoding=encoding)
            except UnicodeDecodeError:
                continue
        raise ValueError("无法识别 CSV 文件的文本编码。")
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)


def sheet_names(file_bytes: bytes, filename: str) -> list[str]:
    if filename.casefold().endswith(".csv"):
        return [filename]
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


def default_sheet(options: list[str], topic: str) -> str:
    if topic == "INV":
        return next((name for name in options if "INV" in name.upper()), options[0])
    return next((name for name in options if "EO" in name.upper() or "E&O" in name.upper()), options[0])


def build_wide_data(raw: pd.DataFrame, mapping: dict[str, Any], filename: str, topic: str) -> pd.DataFrame:
    header_row = mapping["header_row"]
    headers = raw.iloc[header_row].tolist()
    year = default_year(filename)
    periods = mapping.get("periods", {})
    period_row = mapping.get("period_row")
    records: list[dict[str, Any]] = []
    for row_index in range(header_row + 1, len(raw)):
        row = raw.iloc[row_index]
        whs = normalize_whs(row.iloc[mapping["whs_col"]])
        if whs is None:
            continue
        planner = normalize_planner(row.iloc[mapping["planner_col"]])
        target_col = mapping.get("target_col")
        target = number_or_none(row.iloc[target_col]) if target_col is not None else None

        for actual_col in mapping["actual_cols"]:
            period = periods.get(actual_col)
            if period is None:
                period_value = headers[actual_col]
                if period_row is not None:
                    period_value = raw.iloc[period_row, actual_col]
                period = parse_period(period_value, year)
            if period is None:
                continue
            actual = number_or_none(row.iloc[actual_col])
            if actual is None and target is None:
                continue
            records.append(
                {
                    "date": period,
                    "Planner": planner or "未分配",
                    "WHS": whs,
                    "actual": actual,
                    "target": target,
                    "topic": topic,
                    "source": filename,
                }
            )
    return pd.DataFrame(records)


def date_from_filename(filename: str, year: int) -> pd.Timestamp | None:
    period = parse_period(filename, year)
    if period is not None:
        return period
    match = re.search(r"(20\d{2})[-_.](\d{1,2})[-_.](\d{1,2})", filename)
    if not match:
        return None
    try:
        return pd.Timestamp(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None


def build_long_data(raw: pd.DataFrame, mapping: dict[str, Any], filename: str, topic: str) -> pd.DataFrame:
    header_row = mapping["header_row"]
    year = default_year(filename)
    fallback_date = date_from_filename(filename, year)
    records: list[dict[str, Any]] = []
    for row_index in range(header_row + 1, len(raw)):
        row = raw.iloc[row_index]
        whs = normalize_whs(row.iloc[mapping["whs_col"]])
        if whs is None:
            continue
        date_col = mapping.get("date_col")
        period = parse_period(row.iloc[date_col], year) if date_col is not None else fallback_date
        planner = normalize_planner(row.iloc[mapping["planner_col"]])
        actual = number_or_none(row.iloc[mapping["actual_col"]])
        target_col = mapping.get("target_col")
        target = number_or_none(row.iloc[target_col]) if target_col is not None else None
        if actual is None and target is None:
            continue
        records.append(
            {
                "date": period,
                "Planner": planner or "未分配",
                "WHS": whs,
                "actual": actual,
                "target": target,
                "topic": topic,
                "source": filename,
            }
        )
    return pd.DataFrame(records)


def auto_parse(raw: pd.DataFrame, filename: str, topic: str) -> tuple[pd.DataFrame, str]:
    mapping = detect_schema(raw, filename)
    if mapping is None:
        raise ValueError("没有自动识别出列结构。请在左侧展开字段映射并指定表头与列。")
    if mapping["layout"] == "wide":
        result = build_wide_data(raw, mapping, filename, topic)
    else:
        result = build_long_data(raw, mapping, filename, topic)
    if result.empty:
        raise ValueError("识别到了表头，但没有找到可用的 WHS、Planner 和实绩数据。")
    return result, mapping["description"]


def manual_parse(raw: pd.DataFrame, mapping: dict[str, Any], filename: str, topic: str) -> pd.DataFrame:
    if mapping["layout"] == "wide":
        periods: dict[int, pd.Timestamp] = {}
        year = default_year(filename)
        headers = raw.iloc[mapping["header_row"]].tolist()
        for col in mapping["actual_cols"]:
            value = headers[col]
            if mapping.get("period_row") is not None:
                value = raw.iloc[mapping["period_row"], col]
            period = parse_period(value, year)
            if period is not None:
                periods[col] = period
        selected = {**mapping, "periods": periods}
        return build_wide_data(raw, selected, filename, topic)
    return build_long_data(raw, mapping, filename, topic)


def configure_topic(
    topic: str,
    active_file: dict[str, Any],
    available_sheets: list[str],
) -> dict[str, Any]:
    prefix = f"{topic}_{active_file['index']}"
    with st.sidebar.expander(f"{topic} 工作表与字段映射", expanded=False):
        initial_sheet = default_sheet(available_sheets, topic)
        sheet = st.selectbox(
            f"{topic} 工作表",
            available_sheets,
            index=available_sheets.index(initial_sheet),
            key=f"{prefix}_sheet",
        )
        try:
            raw = read_raw_sheet(active_file["bytes"], active_file["name"], sheet)
        except Exception as exc:
            st.error(f"读取工作表失败：{exc}")
            return {"sheet": sheet, "layout": "auto", "raw": pd.DataFrame()}

        detected = detect_schema(raw, active_file["name"])
        layout_choice = st.selectbox(
            "数据布局",
            ["自动识别", "长表（每行一条日期记录）", "宽表（每列一个日期/月份）"],
            key=f"{prefix}_layout",
        )
        if layout_choice == "自动识别":
            if detected:
                st.caption(f"已识别：{detected['description']}。如需自定义，请切换为长表或宽表。")
            else:
                st.warning("自动识别未成功。请切换成长表或宽表并指定列。")
            return {"sheet": sheet, "layout": "auto", "raw": raw}

        guessed_row = detected.get("header_row", guess_header_row(raw)) if detected else guess_header_row(raw)
        header_row = int(
            st.number_input(
                "表头所在行（从 1 开始）",
                min_value=1,
                max_value=max(1, min(30, len(raw))),
                value=guessed_row + 1,
                key=f"{prefix}_header_row",
            )
        ) - 1
        options = column_options(raw, header_row)
        header = raw.iloc[header_row].tolist()
        guess = lambda kind: find_header_column(header, kind)

        planner_col = column_select(
            "Planner 列", options, guess("planner"), f"{prefix}_planner_col"
        )
        whs_col = column_select("WHS 列", options, guess("whs"), f"{prefix}_whs_col")
        target_col = column_select(
            "目标值列", options, guess("target"), f"{prefix}_target_col", allow_none=True
        )

        if layout_choice.startswith("长表"):
            date_guess = guess("date")
            date_choice = st.selectbox(
                "日期/周次列",
                ["使用文件名中的日期"] + options,
                index=(date_guess + 1) if date_guess is not None else 0,
                key=f"{prefix}_date_col",
            )
            actual_col = column_select(
                "实际值列", options, guess("actual"), f"{prefix}_actual_col"
            )
            mapping = {
                "layout": "long",
                "header_row": header_row,
                "planner_col": planner_col,
                "whs_col": whs_col,
                "target_col": target_col,
                "date_col": options.index(date_choice) if date_choice in options else None,
                "actual_col": actual_col,
            }
        else:
            auto_actuals = detected.get("actual_cols", []) if detected else []
            actual_labels = st.multiselect(
                "实际值列（可多选）",
                options,
                default=[options[i] for i in auto_actuals if i < len(options)],
                key=f"{prefix}_actual_cols",
            )
            period_options = ["列标题本身含日期"] + [
                f"第 {row + 1} 行的期间标签" for row in range(header_row)
            ]
            suggested_period_row = detected.get("period_row") if detected else None
            period_default = (
                suggested_period_row + 1 if suggested_period_row is not None else 0
            )
            period_choice = st.selectbox(
                "期间标签所在行",
                period_options,
                index=min(period_default, len(period_options) - 1),
                key=f"{prefix}_period_row",
            )
            mapping = {
                "layout": "wide",
                "header_row": header_row,
                "period_row": (
                    period_options.index(period_choice) - 1
                    if period_choice != "列标题本身含日期"
                    else None
                ),
                "planner_col": planner_col,
                "whs_col": whs_col,
                "target_col": target_col,
                "actual_cols": [options.index(item) for item in actual_labels],
            }

        st.caption("列名会显示 Excel 列字母，方便在不同工作簿中选择对应字段。")
        st.dataframe(raw.iloc[header_row : min(header_row + 5, len(raw))], hide_index=True, use_container_width=True)
        return {"sheet": sheet, "layout": "manual", "mapping": mapping, "raw": raw}


def resolve_sheet(names: list[str], requested: str, topic: str) -> str | None:
    if requested in names:
        return requested
    if topic == "INV":
        return next((name for name in names if "INV" in name.upper()), None)
    return next((name for name in names if "EO" in name.upper() or "E&O" in name.upper()), None)


def load_topic_data(
    sources: list[dict[str, Any]],
    selected_source: dict[str, Any],
    merge_uploads: bool,
    config: dict[str, Any],
    topic: str,
) -> tuple[pd.DataFrame, list[str]]:
    candidates = sources if merge_uploads else [selected_source]
    frames: list[pd.DataFrame] = []
    notes: list[str] = []
    for order, source in enumerate(candidates):
        try:
            names = sheet_names(source["bytes"], source["name"])
            sheet = names[0] if source["name"].casefold().endswith(".csv") else resolve_sheet(names, config["sheet"], topic)
            if sheet is None:
                notes.append(f"{source['name']}：没有找到对应的 {topic} 工作表，已跳过。")
                continue
            raw = read_raw_sheet(source["bytes"], source["name"], sheet)
            if config["layout"] == "auto":
                frame, message = auto_parse(raw, source["name"], topic)
            else:
                frame = manual_parse(raw, config["mapping"], source["name"], topic)
                message = "按手动字段映射读取"
            if not frame.empty:
                frame["_source_order"] = order
                frames.append(frame)
                if order == 0 or not merge_uploads:
                    notes.append(f"{source['name']} · {sheet} · {message}")
        except Exception as exc:
            notes.append(f"{source['name']}：{exc}")

    if not frames:
        return pd.DataFrame(), notes
    result = pd.concat(frames, ignore_index=True)
    result["date"] = pd.to_datetime(result["date"], errors="coerce")
    result["actual"] = pd.to_numeric(result["actual"], errors="coerce")
    result["target"] = pd.to_numeric(result["target"], errors="coerce")
    result = result.dropna(subset=["WHS"])
    result = result.sort_values(["date", "_source_order"], na_position="first")
    result = result.drop_duplicates(["date", "Planner", "WHS"], keep="last")
    return result.reset_index(drop=True), notes


def get_app_secret(name: str) -> str:
    try:
        return str(st.secrets.get(name, "")).strip()
    except Exception:
        return ""


def validate_uploads(files: list[Any]) -> list[dict[str, Any]]:
    validated: list[dict[str, Any]] = []
    for uploaded in files:
        filename = Path(uploaded.name).name
        suffix = Path(filename).suffix.casefold()
        if suffix not in SHARED_UPLOAD_TYPES:
            raise ValueError(f"不支持的文件格式：{filename}")

        file_bytes = uploaded.getvalue()
        if not file_bytes:
            raise ValueError(f"文件为空：{filename}")

        names = sheet_names(file_bytes, filename)
        if not names:
            raise ValueError(f"工作簿没有可读取的工作表：{filename}")

        readable_sheet = False
        for sheet in names:
            try:
                raw = read_raw_sheet(file_bytes, filename, sheet)
            except Exception:
                continue
            if not raw.dropna(how="all").empty:
                readable_sheet = True
                break
        if not readable_sheet:
            raise ValueError(f"文件中没有可读取的数据：{filename}")

        validated.append({"name": filename, "bytes": file_bytes})
    if not validated:
        raise ValueError("请先选择至少一个 Excel 或 CSV 文件。")
    return validated


def build_batch_archive(files: list[dict[str, Any]]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for index, file in enumerate(files, start=1):
            safe_name = Path(file["name"]).name.replace("/", "_").replace("\\", "_")
            archive.writestr(f"{index:02d}_{safe_name}", file["bytes"])
    return output.getvalue()


def make_batch_name() -> str:
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return f"{timestamp}_{uuid4().hex[:8]}.zip"


def commit_batch_to_github(archive_bytes: bytes, batch_name: str) -> None:
    token = get_app_secret("GITHUB_TOKEN")
    if not token:
        raise RuntimeError("尚未配置 GitHub 上传凭据，请检查 Streamlit Cloud 的 Secrets。")

    url = (
        "https://api.github.com/repos/"
        f"{GITHUB_REPOSITORY}/contents/data/weekly/{batch_name}"
    )
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    payload = {
        "message": f"Add weekly KPI data batch {batch_name}",
        "content": base64.b64encode(archive_bytes).decode("ascii"),
        "branch": GITHUB_BRANCH,
    }
    try:
        response = requests.put(url, headers=headers, json=payload, timeout=60)
    except requests.RequestException as exc:
        raise RuntimeError("连接 GitHub 失败，请稍后重试。") from exc

    if response.status_code not in (200, 201):
        try:
            detail = str(response.json().get("message", ""))
        except (ValueError, AttributeError):
            detail = ""
        suffix = f"：{detail}" if detail else "。"
        raise RuntimeError(f"GitHub 保存失败（HTTP {response.status_code}）{suffix}")


def read_shared_batches(
    data_dir: Path,
) -> tuple[list[dict[str, Any]], list[str], set[str]]:
    sources: list[dict[str, Any]] = []
    warnings: list[str] = []
    batch_names: set[str] = set()
    weekly_dir = data_dir / "weekly"

    if not weekly_dir.is_dir():
        return sources, warnings, batch_names

    for archive_path in sorted(weekly_dir.glob("*.zip")):
        batch_names.add(archive_path.name)
        try:
            with zipfile.ZipFile(archive_path) as archive:
                members = [
                    item for item in archive.infolist()
                    if not item.is_dir()
                    and Path(item.filename).suffix.casefold() in SHARED_UPLOAD_TYPES
                ]
                if not members:
                    warnings.append(f"{archive_path.name}：压缩包中没有可读取的数据文件。")
                    continue
                for item in members:
                    display_name = f"{archive_path.stem} · {Path(item.filename).name}"
                    sources.append({
                        "name": display_name,
                        "bytes": archive.read(item),
                        "index": len(sources),
                    })
        except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
            warnings.append(f"{archive_path.name}：无法读取共享批次（{exc}）。")

    return sources, warnings, batch_names


def status_for(actual: Any, target: Any, lower_better: bool) -> str:
    if pd.isna(actual) or pd.isna(target):
        return "待更新"
    if lower_better:
        return "已达标" if actual <= target else "未达标"
    return "已达标" if actual >= target else "未达标"


def sum_or_nan(values: pd.Series) -> float:
    return float(values.sum(min_count=1)) if values.notna().any() else float("nan")


def latest_per_whs(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return data.copy()
    return (
        data.sort_values(["date", "_source_order"], na_position="first")
        .drop_duplicates(["Planner", "WHS"], keep="last")
        .reset_index(drop=True)
    )


def bucket_data(data: pd.DataFrame, grain: str) -> pd.DataFrame:
    empty_result = pd.DataFrame(columns=["Planner", "期间", "actual", "target", "比例"])
    if data.empty:
        return empty_result
    working = data.dropna(subset=["date"]).copy()
    working = working.dropna(subset=["actual", "target"])
    working = working[working["target"] != 0]
    if working.empty:
        return empty_result
    if grain == "按周":
        working["期间"] = working["date"].dt.to_period("W-SUN").dt.start_time
    else:
        working["期间"] = working["date"].dt.to_period("M").dt.to_timestamp()
    working = working.sort_values(["date", "_source_order"], na_position="first")
    working = working.drop_duplicates(["Planner", "WHS", "期间"], keep="last")
    grouped = (
        working.groupby(["Planner", "期间"], as_index=False)
        .agg(
            actual=("actual", lambda values: values.sum(min_count=1)),
            target=("target", lambda values: values.sum(min_count=1)),
        )
        .sort_values("期间")
    )
    grouped["比例"] = grouped["actual"] / grouped["target"] * 100
    return grouped


def fmt(value: Any, decimals: int = 0) -> str:
    if pd.isna(value):
        return "—"
    return f"{float(value):,.{decimals}f}"


def fmt_money(value: Any, decimals: int = 0, signed: bool = False) -> str:
    if pd.isna(value):
        return "—"
    amount = float(value)
    sign = "-" if amount < 0 else "+" if signed and amount > 0 else ""
    return f"{sign}${abs(amount):,.{decimals}f}"


def render_kpi_cards(actual: float, target: float, lower_better: bool) -> dict[str, Any]:
    gap = actual - target if pd.notna(actual) and pd.notna(target) else float("nan")
    status = status_for(actual, target, lower_better)
    color_class = "pass" if status == "已达标" else "fail" if status == "未达标" else "wait"
    gap_text = fmt_money(gap, signed=True)
    ratio = actual / target * 100 if pd.notna(actual) and pd.notna(target) and target != 0 else float("nan")
    columns = st.columns(4)
    values = [
        ("实际值", fmt_money(actual), "所选范围内每个 WHS 的最新记录"),
        ("目标值", fmt_money(target), "按 WHS 目标汇总"),
        ("目标差距", gap_text, "实际值 − 目标值"),
        ("达标状态", status, f"实际/目标 {fmt(ratio, 1)}%"),
    ]
    for column, (label, value, foot) in zip(columns, values):
        with column:
            if label == "达标状态":
                st.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
                    f'<div style="margin-top:12px"><span class="pill pill-{color_class}">{value}</span></div>'
                    f'<div class="kpi-foot">{foot}</div></div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
                    f'<div class="kpi-value">{value}</div><div class="kpi-foot">{foot}</div></div>',
                    unsafe_allow_html=True,
                )
    return {"actual": actual, "target": target, "gap": gap, "status": status, "ratio": ratio}


def make_detail_table(rows: pd.DataFrame, lower_better: bool) -> pd.DataFrame:
    if rows.empty:
        return pd.DataFrame(columns=["Planner", "WHS", "期间", "实际值", "目标值", "差距", "状态"])
    result = rows[["Planner", "WHS", "date", "actual", "target"]].copy()
    result = result.rename(columns={"date": "期间", "actual": "实际值", "target": "目标值"})
    result["期间"] = pd.to_datetime(result["期间"], errors="coerce").dt.strftime("%Y-%m-%d")
    result["差距"] = result["实际值"] - result["目标值"]
    result["状态"] = [
        status_for(actual, target, lower_better)
        for actual, target in zip(result["实际值"], result["目标值"])
    ]
    result = result.sort_values(["Planner", "WHS"], kind="stable")
    return result.reset_index(drop=True)


def style_detail_table(table: pd.DataFrame):
    def row_style(row: pd.Series) -> list[str]:
        status = row.get("状态")
        if status == "已达标":
            style = "background-color: #eaf7f0; color: #187348;"
        elif status == "未达标":
            style = "background-color: #fff0f0; color: #b92f3a;"
        else:
            style = "background-color: #fff7e6; color: #966512;"
        return [style if column in {"差距", "状态"} else "" for column in row.index]

    return table.style.apply(row_style, axis=1).format(
        {
            "实际值": lambda value: fmt_money(value, 2),
            "目标值": lambda value: fmt_money(value, 2),
            "差距": lambda value: fmt_money(value, 2, signed=True),
        },
        na_rep="—",
    )


def planner_comparison(latest: pd.DataFrame, lower_better: bool) -> pd.DataFrame:
    if latest.empty:
        return pd.DataFrame(columns=["Planner", "actual", "target", "比例", "状态"])
    latest = latest.dropna(subset=["actual", "target"])
    latest = latest[latest["target"] != 0]
    if latest.empty:
        return pd.DataFrame(columns=["Planner", "actual", "target", "比例", "状态"])
    grouped = (
        latest.groupby("Planner", as_index=False)
        .agg(
            actual=("actual", lambda values: values.sum(min_count=1)),
            target=("target", lambda values: values.sum(min_count=1)),
        )
    )
    grouped["比例"] = grouped["actual"] / grouped["target"] * 100
    grouped["状态"] = [
        status_for(actual, target, lower_better)
        for actual, target in zip(grouped["actual"], grouped["target"])
    ]
    return grouped.dropna(subset=["比例"]).sort_values("比例", ascending=False)


def make_chart(
    planner: str,
    current_rows: pd.DataFrame,
    trend: pd.DataFrame,
    lower_better: bool,
    grain: str,
) -> go.Figure | None:
    if current_rows.empty:
        return None
    if planner == "全部 Planner":
        comparison = planner_comparison(current_rows, lower_better)
        if comparison.empty:
            return None
        comparison = comparison.sort_values("actual", ascending=lower_better)
        actual_colors = [RED if status == "未达标" else GREEN for status in comparison["状态"]]
        figure = go.Figure()
        figure.add_trace(
            go.Bar(
                x=comparison["target"],
                y=comparison["Planner"],
                orientation="h",
                name="标准值",
                marker_color=TARGET,
                text=[fmt_money(value, 0) for value in comparison["target"]],
                textposition="outside",
                cliponaxis=False,
                customdata=comparison[["actual", "状态"]],
                hovertemplate=(
                    "%{y}<br>标准值：$%{x:,.2f}"
                    "<br>实际值：$%{customdata[0]:,.2f}"
                    "<br>%{customdata[1]}<extra></extra>"
                ),
            )
        )
        figure.add_trace(
            go.Bar(
                x=comparison["actual"],
                y=comparison["Planner"],
                orientation="h",
                name="实际值",
                marker_color=actual_colors,
                text=[fmt_money(value, 0) for value in comparison["actual"]],
                textposition="outside",
                cliponaxis=False,
                customdata=comparison[["target", "状态"]],
                hovertemplate=(
                    "%{y}<br>实际值：$%{x:,.2f}"
                    "<br>标准值：$%{customdata[0]:,.2f}"
                    "<br>%{customdata[1]}<extra></extra>"
                ),
            )
        )
        min_value = min(0.0, float(comparison["actual"].min()), float(comparison["target"].min()))
        max_value = max(0.0, float(comparison["actual"].max()), float(comparison["target"].max()))
        value_span = max(max_value - min_value, abs(max_value), abs(min_value), 1.0)
        min_axis = min_value - value_span * 0.15 if min_value < 0 else 0.0
        max_axis = max_value + value_span * 0.3 if max_value > 0 else value_span * 0.1
        figure.add_vline(x=0, line_color="#cbd5e1", line_width=1)
        figure.update_layout(
            title="各 Planner 实际值与标准值对比",
            xaxis_title="金额（USD）",
            yaxis_title="",
            xaxis=dict(range=[min_axis, max_axis], tickprefix="$", tickformat=","),
            yaxis=dict(autorange="reversed"),
            barmode="group",
            showlegend=True,
        )
    else:
        if trend.empty or not {"Planner", "期间", "actual", "target", "比例"}.issubset(trend.columns):
            return None
        filtered = trend[trend["Planner"] == planner].sort_values("期间")
        if filtered.empty:
            return None
        marker_colors = [RED if status_for(a, t, lower_better) == "未达标" else GREEN
                         for a, t in zip(filtered["actual"], filtered["target"])]
        period_format = "%Y-%m" if grain == "按月" else "%Y-%m-%d"
        period_labels = filtered["期间"].dt.strftime(period_format)
        figure = go.Figure()
        # A single observation cannot form a trend line. Monthly comparisons
        # are also clearer as grouped bars, even when several months are shown.
        use_bars = grain == "按月" or len(filtered) == 1
        if use_bars:
            statuses = [status_for(a, t, lower_better)
                        for a, t in zip(filtered["actual"], filtered["target"])]
            figure.add_trace(
                go.Bar(
                    x=period_labels,
                    y=filtered["target"],
                    name="目标值",
                    marker_color=TARGET,
                    text=[fmt_money(value, 0) for value in filtered["target"]],
                    textposition="outside",
                    cliponaxis=False,
                    hovertemplate="%{x}<br>目标值：$%{y:,.2f}<extra></extra>",
                )
            )
            figure.add_trace(
                go.Bar(
                    x=period_labels,
                    y=filtered["actual"],
                    name="实际值",
                    marker_color=marker_colors,
                    text=[fmt_money(value, 0) for value in filtered["actual"]],
                    textposition="outside",
                    cliponaxis=False,
                    customdata=filtered[["target", "比例"]].assign(状态=statuses),
                    hovertemplate=(
                        "%{x}<br>实际值：$%{y:,.2f}"
                        "<br>目标值：$%{customdata[0]:,.2f}"
                        "<br>实际/目标：%{customdata[1]:.1f}%"
                        "<br>%{customdata[2]}<extra></extra>"
                    ),
                )
            )
            chart_title = f"{planner} 实际值与目标对比（{grain}）"
            xaxis_title = "月份" if grain == "按月" else "周开始日期"
            figure.update_layout(barmode="group")
        else:
            figure.add_trace(
                go.Scatter(
                    x=period_labels,
                    y=filtered["target"],
                    mode="lines+markers",
                    name="目标值",
                    line=dict(color=TARGET, width=2, dash="dash"),
                    hovertemplate="%{x}<br>目标值：$%{y:,.2f}<extra></extra>",
                )
            )
            figure.add_trace(
                go.Scatter(
                    x=period_labels,
                    y=filtered["actual"],
                    mode="lines+markers",
                    name="实际值",
                    line=dict(color=BLUE, width=3),
                    marker=dict(color=marker_colors, size=10, line=dict(color="white", width=1.5)),
                    customdata=filtered[["target", "比例"]],
                    hovertemplate=(
                        "%{x}<br>实际值：$%{y:,.2f}"
                        "<br>目标值：$%{customdata[0]:,.2f}"
                        "<br>实际/目标：%{customdata[1]:.1f}%<extra></extra>"
                    ),
                )
            )
            chart_title = f"{planner} 实际值趋势与目标（{grain}）"
            xaxis_title = "周开始日期"
        figure.update_layout(
            title=chart_title,
            xaxis_title=xaxis_title,
            yaxis_title="金额（USD）",
            xaxis=dict(type="category"),
            yaxis=dict(rangemode="tozero", tickprefix="$", tickformat=","),
            showlegend=True,
        )
    figure.update_layout(
        height=380,
        margin=dict(l=20, r=35, t=60, b=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial, Microsoft YaHei, sans-serif", color="#263244"),
        hoverlabel=dict(namelength=-1),
    )
    return figure


def ppt_bytes(
    topic: str,
    planner: str,
    grain: str,
    start_date: date,
    end_date: date,
    summary: dict[str, Any],
    current_rows: pd.DataFrame,
    trend: pd.DataFrame,
    lower_better: bool,
) -> bytes:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt

    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial"]
    plt.rcParams["axes.unicode_minus"] = False

    def currency_tick(value, _position):
        return f"-${abs(value):,.0f}" if value < 0 else f"${value:,.0f}"

    currency_formatter = FuncFormatter(currency_tick)

    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank = presentation.slide_layouts[6]
    slide = presentation.slides.add_slide(blank)

    title = slide.shapes.add_textbox(Inches(0.55), Inches(0.3), Inches(12.2), Inches(0.55)).text_frame
    title.text = f"KPI 考核 · {topic} 专题"
    title.paragraphs[0].font.size = Pt(24)
    title.paragraphs[0].font.bold = True

    subtitle = slide.shapes.add_textbox(Inches(0.58), Inches(0.92), Inches(12.0), Inches(0.35)).text_frame
    subtitle.text = f"Planner：{planner}    时间：{start_date:%Y-%m-%d} 至 {end_date:%Y-%m-%d}    粒度：{grain}"
    subtitle.paragraphs[0].font.size = Pt(11)
    subtitle.paragraphs[0].font.color.rgb = RGBColor(95, 108, 126)

    card_values = [
        ("实际值", fmt_money(summary["actual"], 2)),
        ("目标值", fmt_money(summary["target"], 2)),
        ("差距", fmt_money(summary["gap"], 2, signed=True)),
        ("达标状态", summary["status"]),
    ]
    card_width = 2.95
    for index, (label, value) in enumerate(card_values):
        x = Inches(0.58 + index * 3.1)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.4), Inches(card_width), Inches(0.9))
        shape.fill.solid()
        if label == "达标状态":
            color = GREEN if summary["status"] == "已达标" else RED if summary["status"] == "未达标" else TARGET
            shape.fill.fore_color.rgb = RGBColor.from_string(color.replace("#", "").upper())
        else:
            shape.fill.fore_color.rgb = RGBColor(244, 247, 250)
        shape.line.color.rgb = RGBColor(221, 227, 234)
        tf = shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = f"{label}\n{value}"
        p.font.size = Pt(15)
        p.font.bold = True

    chart_buffer = io.BytesIO()
    fig, ax = plt.subplots(figsize=(11.5, 4.0))
    if planner == "全部 Planner":
        chart_data = planner_comparison(current_rows, lower_better).sort_values("actual", ascending=lower_better)
        if not chart_data.empty:
            colors = [RED if status == "未达标" else GREEN for status in chart_data["状态"]]
            positions = list(range(len(chart_data)))
            bar_height = 0.34
            target_positions = [position - bar_height / 2 for position in positions]
            actual_positions = [position + bar_height / 2 for position in positions]
            ax.barh(target_positions, chart_data["target"], height=bar_height,
                    color=TARGET, label="标准值")
            ax.barh(actual_positions, chart_data["actual"], height=bar_height,
                    color=colors, label="实际值")
            ax.set_yticks(positions, chart_data["Planner"])
            ax.invert_yaxis()
            min_value = min(0.0, float(chart_data["actual"].min()), float(chart_data["target"].min()))
            max_value = max(0.0, float(chart_data["actual"].max()), float(chart_data["target"].max()))
            value_span = max(max_value - min_value, abs(max_value), abs(min_value), 1.0)
            for y, target_value, actual_value in zip(positions, chart_data["target"], chart_data["actual"]):
                for label_y, value in ((y - bar_height / 2, target_value), (y + bar_height / 2, actual_value)):
                    offset = value_span * 0.015
                    ax.text(value + offset if value >= 0 else value - offset, label_y, fmt_money(value, 0),
                            va="center", ha="left" if value >= 0 else "right", fontsize=8)
            ax.set_xlim(
                min_value - value_span * 0.15 if min_value < 0 else 0,
                max_value + value_span * 0.3 if max_value > 0 else value_span * 0.1,
            )
            ax.axvline(0, color="#cbd5e1", linewidth=1)
            ax.set_xlabel("金额（USD）")
            ax.set_title("各 Planner 实际值与标准值对比")
            ax.legend()
            ax.xaxis.set_major_formatter(currency_formatter)
    else:
        chart_data = trend[trend["Planner"] == planner].sort_values("期间")
        if not chart_data.empty:
            use_bars = grain == "按月" or len(chart_data) == 1
            if use_bars:
                labels = chart_data["期间"].dt.strftime("%Y-%m" if grain == "按月" else "%Y-%m-%d")
                positions = list(range(len(chart_data)))
                width = 0.36
                colors = [RED if status_for(row["actual"], row["target"], lower_better) == "未达标" else GREEN
                          for _, row in chart_data.iterrows()]
                target_bars = ax.bar([pos - width / 2 for pos in positions], chart_data["target"], width,
                                     color=TARGET, label="目标值")
                actual_bars = ax.bar([pos + width / 2 for pos in positions], chart_data["actual"], width,
                                     color=colors, label="实际值")
                ax.bar_label(target_bars, labels=[fmt_money(value, 0) for value in chart_data["target"]],
                             padding=2, fontsize=7)
                ax.bar_label(actual_bars, labels=[fmt_money(value, 0) for value in chart_data["actual"]],
                             padding=2, fontsize=7)
                ax.set_xticks(positions, labels)
                ax.set_xlabel("月份" if grain == "按月" else "周开始日期")
                ax.set_title(f"{planner} 实际值与目标对比")
            else:
                ax.plot(chart_data["期间"], chart_data["target"], color=TARGET, linestyle="--", linewidth=1.8, label="目标值")
                ax.plot(chart_data["期间"], chart_data["actual"], color=BLUE, linewidth=2.5, label="实际值")
                for _, row in chart_data.iterrows():
                    color = RED if status_for(row["actual"], row["target"], lower_better) == "未达标" else GREEN
                    ax.scatter(row["期间"], row["actual"], color=color, s=42, zorder=3)
                ax.set_title(f"{planner} 实际值趋势与目标")
                fig.autofmt_xdate()
            ax.set_ylabel("金额（USD）")
            ax.yaxis.set_major_formatter(currency_formatter)
            ax.legend()
    ax.grid(axis="x", alpha=0.2)
    fig.tight_layout()
    fig.savefig(chart_buffer, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    chart_buffer.seek(0)
    slide.shapes.add_picture(chart_buffer, Inches(0.7), Inches(2.55), width=Inches(11.9), height=Inches(4.45))

    detail = make_detail_table(current_rows, lower_better)
    page_size = 18
    if detail.empty:
        detail = pd.DataFrame([["—", "—", "—", None, None, None, "待更新"]],
                              columns=["Planner", "WHS", "期间", "实际值", "目标值", "差距", "状态"])
    columns = ["Planner", "WHS", "期间", "实际值", "目标值", "差距", "状态"]
    for page_index in range(0, len(detail), page_size):
        slide = presentation.slides.add_slide(blank)
        heading = slide.shapes.add_textbox(Inches(0.55), Inches(0.3), Inches(12.2), Inches(0.45)).text_frame
        heading.text = f"{topic} · Planner × WHS 明细"
        heading.paragraphs[0].font.size = Pt(20)
        heading.paragraphs[0].font.bold = True
        chunk = detail.iloc[page_index : page_index + page_size]
        table_shape = slide.shapes.add_table(
            len(chunk) + 1, len(columns), Inches(0.45), Inches(1.0), Inches(12.4), Inches(5.9)
        )
        table = table_shape.table
        widths = [1.55, 0.9, 1.45, 1.95, 1.95, 1.95, 1.3]
        for col, width in enumerate(widths):
            table.columns[col].width = Inches(width)
        for col, name in enumerate(columns):
            cell = table.cell(0, col)
            cell.text = name
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(235, 240, 246)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.bold = True
                paragraph.font.size = Pt(10)
        for row_index, (_, row) in enumerate(chunk.iterrows(), start=1):
            for col_index, name in enumerate(columns):
                value = row[name]
                if name in {"实际值", "目标值", "差距"}:
                    text = fmt_money(value, 2, signed=name == "差距")
                else:
                    text = "—" if pd.isna(value) else str(value)
                cell = table.cell(row_index, col_index)
                cell.text = text
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.size = Pt(9)
                if name in {"状态", "差距"}:
                    cell.fill.solid()
                    if row["状态"] == "已达标":
                        cell.fill.fore_color.rgb = RGBColor(234, 247, 240)
                    elif row["状态"] == "未达标":
                        cell.fill.fore_color.rgb = RGBColor(255, 240, 240)
                    else:
                        cell.fill.fore_color.rgb = RGBColor(255, 247, 230)

    output = io.BytesIO()
    presentation.save(output)
    return output.getvalue()


def pdf_bytes(
    topic: str,
    planner: str,
    grain: str,
    start_date: date,
    end_date: date,
    summary: dict[str, Any],
    current_rows: pd.DataFrame,
    trend: pd.DataFrame,
    lower_better: bool,
) -> bytes:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    font_name = "STSong-Light"
    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont(font_name))

    page_size = landscape(A4)
    page_width, page_height = page_size
    left_margin = right_margin = 14 * mm
    top_margin = 13 * mm
    bottom_margin = 15 * mm
    output = io.BytesIO()
    document = SimpleDocTemplate(
        output,
        pagesize=page_size,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
        title=f"KPI 考核 · {topic} 专题",
        author="KPI 考核",
    )

    title_style = ParagraphStyle(
        "KpiPdfTitle", fontName=font_name, fontSize=20, leading=26,
        textColor=colors.HexColor("#202938"), spaceAfter=5,
    )
    body_style = ParagraphStyle(
        "KpiPdfBody", fontName=font_name, fontSize=9, leading=13,
        textColor=colors.HexColor("#263244"), wordWrap="CJK",
    )
    small_style = ParagraphStyle(
        "KpiPdfSmall", fontName=font_name, fontSize=8, leading=11,
        textColor=colors.HexColor("#263244"), wordWrap="CJK",
    )
    header_style = ParagraphStyle(
        "KpiPdfHeader", fontName=font_name, fontSize=9, leading=12,
        textColor=colors.white, alignment=1, wordWrap="CJK",
    )
    section_style = ParagraphStyle(
        "KpiPdfSection", fontName=font_name, fontSize=14, leading=18,
        textColor=colors.HexColor("#202938"), spaceAfter=7,
    )

    story = [
        Paragraph(f"KPI 考核 · {escape(topic)} 专题", title_style),
        Paragraph(
            f"Planner：{escape(planner)}　　时间：{start_date:%Y-%m-%d} 至 {end_date:%Y-%m-%d}　　粒度：{escape(grain)}",
            body_style,
        ),
        Spacer(1, 7 * mm),
    ]

    card_values = [
        ("实际值", fmt_money(summary["actual"], 2)),
        ("目标值", fmt_money(summary["target"], 2)),
        ("目标差距", fmt_money(summary["gap"], 2, signed=True)),
        ("达标状态", str(summary["status"])),
    ]
    card_width = (page_width - left_margin - right_margin) / 4
    card_cells = [[
        Paragraph(
            f"<font color='#667386'>{escape(label)}</font><br/><font size='14'><b>{escape(value)}</b></font>",
            body_style,
        )
        for label, value in card_values
    ]]
    cards = Table(card_cells, colWidths=[card_width] * 4, rowHeights=[19 * mm])
    card_commands = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dce3eb")),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f4f7fa")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]
    status_color = GREEN if summary["status"] == "已达标" else RED if summary["status"] == "未达标" else TARGET
    card_commands.append(("BACKGROUND", (3, 0), (3, 0), colors.HexColor(status_color)))
    card_commands.append(("TEXTCOLOR", (3, 0), (3, 0), colors.white))
    cards.setStyle(TableStyle(card_commands))
    story.extend([cards, Spacer(1, 7 * mm)])

    matplotlib.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial"]
    matplotlib.rcParams["axes.unicode_minus"] = False

    def currency_tick(value, _position):
        return f"-${abs(value):,.0f}" if value < 0 else f"${value:,.0f}"

    currency_formatter = FuncFormatter(currency_tick)
    fig, ax = plt.subplots(figsize=(11.2, 4.0))
    chart_available = False
    if planner == "全部 Planner":
        chart_data = planner_comparison(current_rows, lower_better).sort_values(
            "actual", ascending=lower_better
        )
        if not chart_data.empty:
            chart_available = True
            bar_colors = [RED if status == "未达标" else GREEN for status in chart_data["状态"]]
            positions = list(range(len(chart_data)))
            bar_height = 0.34
            target_positions = [position - bar_height / 2 for position in positions]
            actual_positions = [position + bar_height / 2 for position in positions]
            ax.barh(target_positions, chart_data["target"], height=bar_height,
                    color=TARGET, label="标准值")
            ax.barh(actual_positions, chart_data["actual"], height=bar_height,
                    color=bar_colors, label="实际值")
            ax.set_yticks(positions, chart_data["Planner"])
            ax.invert_yaxis()
            min_value = min(0.0, float(chart_data["actual"].min()), float(chart_data["target"].min()))
            max_value = max(0.0, float(chart_data["actual"].max()), float(chart_data["target"].max()))
            value_span = max(max_value - min_value, abs(max_value), abs(min_value), 1.0)
            for y, target_value, actual_value in zip(positions, chart_data["target"], chart_data["actual"]):
                for label_y, value in ((y - bar_height / 2, target_value), (y + bar_height / 2, actual_value)):
                    offset = value_span * 0.015
                    ax.text(
                        value + offset if value >= 0 else value - offset,
                        label_y,
                        fmt_money(value, 0),
                        va="center",
                        ha="left" if value >= 0 else "right",
                        fontsize=8,
                    )
            ax.set_xlim(
                min_value - value_span * 0.15 if min_value < 0 else 0,
                max_value + value_span * 0.3 if max_value > 0 else value_span * 0.1,
            )
            ax.axvline(0, color="#cbd5e1", linewidth=1)
            ax.set_xlabel("金额（USD）")
            ax.set_title("各 Planner 实际值与标准值对比")
            ax.legend()
            ax.xaxis.set_major_formatter(currency_formatter)
    else:
        chart_data = trend[trend["Planner"] == planner].sort_values("期间") if not trend.empty else trend
        if not chart_data.empty:
            chart_available = True
            use_bars = grain == "按月" or len(chart_data) == 1
            if use_bars:
                labels = chart_data["期间"].dt.strftime("%Y-%m" if grain == "按月" else "%Y-%m-%d")
                positions = list(range(len(chart_data)))
                width = 0.36
                bar_colors = [
                    RED if status_for(row["actual"], row["target"], lower_better) == "未达标" else GREEN
                    for _, row in chart_data.iterrows()
                ]
                target_bars = ax.bar([pos - width / 2 for pos in positions], chart_data["target"], width,
                                     color=TARGET, label="目标值")
                actual_bars = ax.bar([pos + width / 2 for pos in positions], chart_data["actual"], width,
                                     color=bar_colors, label="实际值")
                ax.bar_label(target_bars, labels=[fmt_money(value, 0) for value in chart_data["target"]],
                             padding=2, fontsize=7)
                ax.bar_label(actual_bars, labels=[fmt_money(value, 0) for value in chart_data["actual"]],
                             padding=2, fontsize=7)
                ax.set_xticks(positions, labels)
                ax.set_xlabel("月份" if grain == "按月" else "周开始日期")
                ax.set_title(f"{planner} 实际值与目标对比")
            else:
                ax.plot(chart_data["期间"], chart_data["target"], color=TARGET, linestyle="--",
                        linewidth=1.8, label="目标值")
                ax.plot(chart_data["期间"], chart_data["actual"], color=BLUE, linewidth=2.5,
                        label="实际值")
                for _, row in chart_data.iterrows():
                    color = RED if status_for(row["actual"], row["target"], lower_better) == "未达标" else GREEN
                    ax.scatter(row["期间"], row["actual"], color=color, s=42, zorder=3)
                ax.set_title(f"{planner} 实际值趋势与目标")
                fig.autofmt_xdate()
            ax.set_ylabel("金额（USD）")
            ax.yaxis.set_major_formatter(currency_formatter)
            ax.legend()
    chart_buffer = io.BytesIO()
    if chart_available:
        ax.grid(axis="x", alpha=0.2)
        fig.tight_layout()
        fig.savefig(chart_buffer, format="png", dpi=160, bbox_inches="tight")
        chart_buffer.seek(0)
        chart_image = Image(chart_buffer, width=page_width - left_margin - right_margin, height=93 * mm)
        chart_image.hAlign = "CENTER"
        story.append(chart_image)
    else:
        story.append(Paragraph("所选范围没有可绘制的实际值和目标值记录。", body_style))
    plt.close(fig)

    story.extend([PageBreak(), Paragraph("Planner × WHS 达标明细", section_style)])
    detail = make_detail_table(current_rows, lower_better)
    columns = ["Planner", "WHS", "期间", "实际值", "目标值", "差距", "状态"]
    table_data = [[Paragraph(escape(name), header_style) for name in columns]]
    if detail.empty:
        table_data.append([Paragraph("—", small_style)] * len(columns))
    else:
        for _, row in detail.iterrows():
            table_data.append([
                Paragraph(escape(str(row["Planner"])), small_style),
                Paragraph(escape(str(row["WHS"])), small_style),
                Paragraph(escape(str(row["期间"])), small_style),
                Paragraph(fmt_money(row["实际值"], 2), small_style),
                Paragraph(fmt_money(row["目标值"], 2), small_style),
                Paragraph(fmt_money(row["差距"], 2, signed=True), small_style),
                Paragraph(escape(str(row["状态"])), small_style),
            ])

    detail_widths = [35 * mm, 25 * mm, 31 * mm, 39 * mm, 39 * mm, 36 * mm, 28 * mm]
    detail_table = Table(table_data, colWidths=detail_widths, repeatRows=1, hAlign="LEFT")
    table_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#52647a")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#d8dee7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for index, status in enumerate(detail["状态"], start=1):
        fill = "#eaf7f0" if status == "已达标" else "#fff0f0" if status == "未达标" else "#fff7e6"
        table_commands.append(("BACKGROUND", (5, index), (6, index), colors.HexColor(fill)))
    detail_table.setStyle(TableStyle(table_commands))
    story.append(detail_table)

    def draw_footer(canvas, doc) -> None:
        canvas.saveState()
        canvas.setFont(font_name, 8)
        canvas.setFillColor(colors.HexColor("#667386"))
        canvas.drawString(left_margin, 7 * mm, f"KPI 考核 · {topic}")
        canvas.drawRightString(page_width - right_margin, 7 * mm, f"第 {doc.page} 页")
        canvas.restoreState()

    document.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
    return output.getvalue()


def render_topic(topic: str, data: pd.DataFrame, source_note: str) -> None:
    if data.empty or data["date"].dropna().empty:
        st.warning(f"{topic} 没有可用于按时间分析的记录。请检查表头与字段映射。")
        return

    valid_dates = data["date"].dropna().sort_values().unique()
    latest_date = pd.Timestamp(valid_dates[-1]).date()
    gaps = pd.Series(valid_dates).diff().dropna().dt.days
    monthly_default = gaps.empty or float(gaps.median()) > 14
    default_grain = "按月" if monthly_default else "按周"
    planner_names = sorted(data["Planner"].dropna().astype(str).unique().tolist())

    control_1, control_2, control_3 = st.columns([1.2, 1.1, 2.5])
    with control_1:
        planner = st.selectbox("Planner", ["全部 Planner"] + planner_names, key=f"{topic}_planner_filter")
    with control_2:
        grain = st.radio(
            "时间粒度",
            ["按周", "按月"],
            index=0 if default_grain == "按周" else 1,
            horizontal=True,
            key=f"{topic}_grain_filter",
        )
    if grain == "按月":
        start_default = latest_date.replace(day=1)
        end_default = latest_date.replace(day=calendar.monthrange(latest_date.year, latest_date.month)[1])
    else:
        start_default = latest_date - timedelta(days=latest_date.weekday())
        end_default = start_default + timedelta(days=6)
    with control_3:
        chosen_range = st.date_input(
            "日期范围",
            value=(start_default, end_default),
            format="YYYY-MM-DD",
            key=f"{topic}_date_range",
        )

    if isinstance(chosen_range, tuple) and len(chosen_range) == 2:
        start_date, end_date = chosen_range
    elif isinstance(chosen_range, date):
        start_date = end_date = chosen_range
    else:
        start_date, end_date = start_default, end_default
    if start_date > end_date:
        start_date, end_date = end_date, start_date

    lower_better = st.selectbox(
        "达标规则",
        ["实际值 ≤ 目标值", "实际值 ≥ 目标值"],
        index=0,
        key=f"{topic}_direction",
        help="INV 通常是库存越低越好。如 EO 指标定义相反，可在这里切换。",
    ) == "实际值 ≤ 目标值"

    filtered = data[
        data["date"].between(pd.Timestamp(start_date), pd.Timestamp(end_date) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1))
    ].copy()
    if planner != "全部 Planner":
        filtered = filtered[filtered["Planner"] == planner]
    weekly_unavailable = grain == "按周" and not gaps.empty and float(gaps.median()) > 14
    if weekly_unavailable:
        filtered = filtered.iloc[0:0].copy()

    current_rows = latest_per_whs(filtered)
    actual_total = sum_or_nan(current_rows["actual"]) if not current_rows.empty else float("nan")
    target_total = sum_or_nan(current_rows["target"]) if not current_rows.empty else float("nan")
    summary = render_kpi_cards(actual_total, target_total, lower_better)

    if weekly_unavailable:
        st.info("当前数据是按月记录，没有真实的周度记录。请切换到“按月”查看，或上传包含每周日期的数据。")

    trend = bucket_data(filtered, grain)
    figure = make_chart(planner, current_rows, trend, lower_better, grain)
    if figure is None:
        st.info("所选时间范围没有可比较的实际值与目标值。")
    else:
        st.plotly_chart(figure, use_container_width=True)

    st.subheader("Planner × WHS 达标明细")
    detail = make_detail_table(current_rows, lower_better)
    if detail.empty:
        st.info("所选范围没有 WHS 明细。")
    else:
        st.dataframe(
            style_detail_table(detail),
            use_container_width=True,
            hide_index=True,
            height=min(600, 38 + 35 * len(detail)),
        )

    try:
        export_data = ppt_bytes(
            topic,
            planner,
            grain,
            start_date,
            end_date,
            summary,
            current_rows,
            trend,
            lower_better,
        )
        st.download_button(
            "导出当前视图为 PPT",
            data=export_data,
            file_name=f"Movex_{topic}_KPI_{start_date:%Y%m%d}_{end_date:%Y%m%d}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key=f"{topic}_ppt_download",
        )
    except Exception as exc:
        st.error(f"PPT 文件生成失败：{exc}")

    try:
        export_pdf = pdf_bytes(
            topic,
            planner,
            grain,
            start_date,
            end_date,
            summary,
            current_rows,
            trend,
            lower_better,
        )
        st.download_button(
            "导出当前视图为 PDF",
            data=export_pdf,
            file_name=f"KPI_{topic}_{start_date:%Y%m%d}_{end_date:%Y%m%d}.pdf",
            mime="application/pdf",
            key=f"{topic}_pdf_download",
        )
    except Exception as exc:
        st.error(f"PDF 文件生成失败：{exc}")


def main() -> None:
    st.title("KPI 考核")
    st.caption("每周上传的数据会累积到共享历史中；所有访客查看同一份 KPI 数据。")

    data_dir = Path(__file__).resolve().parent / "data"
    github_token = get_app_secret("GITHUB_TOKEN")
    with st.sidebar:
        st.header("共享数据")
        st.caption("所有访客都可以上传；每次上传会追加为新的共享批次。")
        if not github_token:
            st.info("共享上传尚未配置。请在 Streamlit Cloud 的 Secrets 中设置 GitHub 写入凭据。")
        else:
            with st.form("kpi_shared_upload_form", clear_on_submit=True):
                uploaded_files = st.file_uploader(
                    "选择 Excel/CSV 文件（可多选）",
                    type=["xlsx", "xlsm", "xls", "csv"],
                    accept_multiple_files=True,
                    key="kpi_shared_batch_files",
                )
                if uploaded_files:
                    st.caption("本批次文件：" + "、".join(file.name for file in uploaded_files))
                submit_upload = st.form_submit_button("追加到共享历史")

            if submit_upload:
                try:
                    validated_files = validate_uploads(uploaded_files or [])
                    archive_bytes = build_batch_archive(validated_files)
                    batch_name = make_batch_name()
                    commit_batch_to_github(archive_bytes, batch_name)

                    pending_batches = st.session_state.setdefault("kpi_pending_batches", {})
                    pending_batches[batch_name] = validated_files
                    st.session_state["kpi_pending_batches"] = pending_batches
                    st.success("上传成功，已添加到共享历史。")
                    st.info("网站正在同步新数据；其他访客刷新后即可看到更新。")
                except Exception as exc:
                    st.error(f"上传失败：{exc}")

    sources, archive_warnings, stored_batch_names = read_shared_batches(data_dir)
    pending_batches = st.session_state.get("kpi_pending_batches", {})
    for batch_name, files in sorted(pending_batches.items()):
        if batch_name in stored_batch_names:
            continue
        batch_label = Path(batch_name).stem
        for file in files:
            sources.append({
                "name": f"{batch_label} · {file['name']}",
                "bytes": file["bytes"],
                "index": len(sources),
            })

    if archive_warnings:
        with st.sidebar:
            for warning in archive_warnings:
                st.warning(warning)

    if not sources:
        st.info("目前还没有共享 KPI 数据。可以在左侧上传第一批 Excel 文件。")
        st.stop()

    with st.sidebar:
        st.caption(f"当前累计 {len(sources)} 个数据文件；每次上传会追加新批次。")

    source_labels = [f"{index + 1}. {source['name']}" for index, source in enumerate(sources)]
    with st.sidebar:
        active_label = st.selectbox("当前工作簿", source_labels, index=len(source_labels) - 1)
        active_index = source_labels.index(active_label)
        active_file = sources[active_index]
        merge_uploads = st.checkbox(
            "合并累计历史文件",
            value=len(sources) > 1,
            help="按上传先后合并；同一日期、Planner 和 WHS 重复时采用较新的记录。",
        )
        try:
            available_sheets = sheet_names(active_file["bytes"], active_file["name"])
        except Exception as exc:
            st.error(f"无法打开工作簿：{exc}")
            st.stop()
        if active_file["name"].casefold().endswith(".csv"):
            st.caption("CSV 文件作为一个数据表读取。")

        inv_config = configure_topic("INV", active_file, available_sheets)
        eo_config = configure_topic("EO", active_file, available_sheets)

    inv_data, inv_notes = load_topic_data(sources, active_file, merge_uploads, inv_config, "INV")
    eo_data, eo_notes = load_topic_data(sources, active_file, merge_uploads, eo_config, "EO")
    inv_note = "；".join(inv_notes)
    eo_note = "；".join(eo_notes)

    inv_tab, eo_tab = st.tabs(["On Hand INV 专题", "EO 专题"])
    with inv_tab:
        render_topic("INV", inv_data, inv_note)
    with eo_tab:
        render_topic("EO", eo_data, eo_note)


if __name__ == "__main__":
    main()

