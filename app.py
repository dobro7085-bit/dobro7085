from __future__ import annotations

import io
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots


APP_TITLE = "AGM 감액량 분석 대시보드"
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_STANDARD_PATH = BASE_DIR / "app" / "data" / "standard.xlsx"
STANDARD_RELATIVE_PATH = Path("app") / "data" / "standard.xlsx"

FIXED_START_ROW = 5
DATE_INDEX = 70  # Excel BS column
SERIAL_INDEX = 1  # Excel B column
MODEL_INDEX = 3  # Excel D column
REDUCTION_INDEX = 68  # Excel BQ column
STANDARD_MODEL_INDEX = 0  # Excel A column
STANDARD_UCL_INDEX = 4  # Excel E column
STANDARD_CENTER_INDEX = 5  # Excel F column
STANDARD_LCL_INDEX = 6  # Excel G column

STANDARD_SHEET_NAMES = {"기준정보", "기준", "spec", "master"}

BACKDATA_CANDIDATES = {
    "date": ["검사일자", "생산일자", "날짜", "일자", "date", "inspectiondate", "inspection date"],
    "serial": ["시리얼넘버", "시리얼", "serial", "serialno", "serialnumber", "s/n", "sn"],
    "model": ["형명", "제품형명", "모델", "제품명", "model", "product"],
    "reduction": ["감액량", "agm감액량", "loss", "reduction", "value"],
}

STANDARD_CANDIDATES = {
    "model": ["형명", "제품형명", "모델", "제품명", "model", "product"],
    "center": ["중심치", "cl", "center", "target", "기준값"],
    "ucl": ["상한선", "ucl", "upper", "usl"],
    "lcl": ["하한선", "lcl", "lower", "lsl"],
}

RESULT_COLORS = {
    "합격": "#2563eb",
    "부적합": "#dc2626",
    "판정 제외": "#737373",
}


def runtime_base_dir() -> Path:
    """EXE 배포 시에는 실행 파일 폴더, 개발 실행 시에는 소스 폴더를 기준으로 사용한다."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return BASE_DIR


def default_standard_paths() -> List[Path]:
    """팀원이 교체 가능한 외부 standard.xlsx를 우선 찾고, 개발용 경로를 fallback으로 둔다."""
    candidates = [
        Path.cwd() / STANDARD_RELATIVE_PATH,
        runtime_base_dir() / STANDARD_RELATIVE_PATH,
        DEFAULT_STANDARD_PATH,
    ]

    unique_paths: List[Path] = []
    seen: set[str] = set()
    for path in candidates:
        resolved_key = str(path.resolve()) if path.exists() else str(path)
        if resolved_key not in seen:
            unique_paths.append(path)
            seen.add(resolved_key)
    return unique_paths


def excel_column_name(zero_based_index: int) -> str:
    """0부터 시작하는 컬럼 인덱스를 Excel 컬럼 문자로 바꾼다."""
    number = zero_based_index + 1
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def normalize_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"[\s_\-./()]+", "", str(value).strip().lower())


def is_blank(value: Any) -> bool:
    if pd.isna(value):
        return True
    text = str(value).strip()
    return text == "" or text.lower() in {"nan", "none", "nat", "null"}


def clean_text(value: Any) -> str:
    if is_blank(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def make_unique_headers(headers: Iterable[Any]) -> List[str]:
    """Excel 헤더명이 중복될 때 pandas 컬럼명으로 안전하게 만든다."""
    seen: Dict[str, int] = {}
    unique_headers: List[str] = []
    for idx, header in enumerate(headers):
        name = clean_text(header) or f"Column {idx + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        unique_headers.append(name)
    return unique_headers


def get_excel_bytes(uploaded_file: Any) -> Optional[bytes]:
    if uploaded_file is None:
        return None
    return uploaded_file.getvalue()


def get_excel_file_name(uploaded_file: Any) -> str:
    if uploaded_file is None:
        return ""
    return str(getattr(uploaded_file, "name", "") or "")


def get_excel_engine(file_name: Optional[str]) -> str:
    """Excel 확장자에 맞는 pandas 엔진을 선택한다. .xls는 xlrd가 필요하다."""
    suffix = Path(str(file_name or "")).suffix.lower()
    if suffix == ".xls":
        return "xlrd"
    return "openpyxl"


def get_sheet_names(file_bytes: bytes, file_name: Optional[str] = None) -> List[str]:
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes), engine=get_excel_engine(file_name))
    return excel_file.sheet_names


def is_standard_sheet_name(sheet_name: str) -> bool:
    normalized = normalize_text(sheet_name)
    return normalized in STANDARD_SHEET_NAMES


def find_standard_sheet(sheet_names: List[str]) -> Optional[str]:
    for sheet_name in sheet_names:
        if is_standard_sheet_name(sheet_name):
            return sheet_name
    return None


def load_excel_file(
    file_bytes: bytes,
    sheet_name: Optional[str] = None,
    header: Optional[int] = 0,
    file_name: Optional[str] = None,
) -> pd.DataFrame:
    """일반 Excel 시트를 DataFrame으로 읽는다."""
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header, engine=get_excel_engine(file_name))


def row_looks_like_header(row: pd.Series) -> bool:
    keywords = []
    for candidates in BACKDATA_CANDIDATES.values():
        keywords.extend(candidates)
    normalized_keywords = {normalize_text(keyword) for keyword in keywords}

    match_count = 0
    for value in row.dropna().tolist():
        normalized_value = normalize_text(value)
        if not normalized_value:
            continue
        if normalized_value in normalized_keywords:
            match_count += 1
            continue
        if any(len(keyword) >= 3 and keyword in normalized_value for keyword in normalized_keywords):
            match_count += 1
    return match_count >= 2


def load_backdata_excel(
    uploaded_file: Any,
    sheet_name: Optional[str] = None,
    start_row: int = FIXED_START_ROW,
    header_mode: str = "자동 판단",
) -> Tuple[pd.DataFrame, bool]:
    """
    백데이터 Excel 파일을 읽는다.
    데이터는 기본적으로 5행부터 시작하며,
    BS열은 날짜, B열은 시리얼넘버, D열은 원본 형명, BQ열은 감액량으로 사용한다.
    """
    file_bytes = get_excel_bytes(uploaded_file)
    if file_bytes is None:
        return pd.DataFrame(), False

    file_name = get_excel_file_name(uploaded_file)
    raw_df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        skiprows=max(start_row - 1, 0),
        engine=get_excel_engine(file_name),
    )
    raw_df = raw_df.dropna(how="all").reset_index(drop=True)
    if raw_df.empty:
        return raw_df, False

    if header_mode == "5행을 컬럼명으로 사용":
        use_first_row_as_header = True
    elif header_mode == "5행부터 실제 데이터":
        use_first_row_as_header = False
    else:
        use_first_row_as_header = row_looks_like_header(raw_df.iloc[0])

    if use_first_row_as_header:
        raw_df.columns = make_unique_headers(raw_df.iloc[0].tolist())
        raw_df = raw_df.iloc[1:].reset_index(drop=True)
    else:
        raw_df.columns = list(range(raw_df.shape[1]))

    return raw_df, use_first_row_as_header


def detect_columns(df: pd.DataFrame, candidates: Dict[str, List[str]]) -> Dict[str, Optional[Any]]:
    """컬럼명 후보 목록을 기준으로 날짜, 형명, 감액량 같은 필드를 자동 추정한다."""
    detected: Dict[str, Optional[Any]] = {}
    normalized_columns = {column: normalize_text(column) for column in df.columns}

    for field, field_candidates in candidates.items():
        normalized_candidates = [normalize_text(candidate) for candidate in field_candidates]
        selected = None

        for column, normalized_column in normalized_columns.items():
            if normalized_column in normalized_candidates:
                selected = column
                break

        if selected is None:
            for column, normalized_column in normalized_columns.items():
                for candidate in normalized_candidates:
                    if len(candidate) <= 2:
                        continue
                    if candidate in normalized_column or normalized_column in candidate:
                        selected = column
                        break
                if selected is not None:
                    break

        detected[field] = selected

    return detected


def parse_date_value(value: Any) -> pd.Timestamp:
    if is_blank(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value
    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return pd.to_datetime(value, errors="coerce")

    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        if np.isfinite(value) and 1 <= float(value) <= 80000:
            return pd.to_datetime(float(value), unit="D", origin="1899-12-30", errors="coerce")
        return pd.NaT

    text = str(value).strip()
    numeric_text = text.replace(",", "")
    if re.fullmatch(r"\d+(\.0+)?", numeric_text):
        numeric_value = float(numeric_text)
        if 1 <= numeric_value <= 80000:
            return pd.to_datetime(numeric_value, unit="D", origin="1899-12-30", errors="coerce")

    return pd.to_datetime(text, errors="coerce")


def parse_date_series(series: pd.Series) -> pd.Series:
    return series.apply(parse_date_value)


def fill_invalid_dates_with_previous_plus_seconds(parsed_dates: pd.Series) -> Tuple[pd.Series, pd.Series]:
    """날짜 변환 실패 행은 바로 위 최종 날짜에 3초를 더해 보정한다."""
    filled_dates: List[pd.Timestamp] = []
    methods: List[str] = []
    previous_date = pd.NaT

    for value in parsed_dates:
        if pd.notna(value):
            current_date = pd.Timestamp(value)
            filled_dates.append(current_date)
            methods.append("원본 날짜 사용")
            previous_date = current_date
        elif pd.notna(previous_date):
            current_date = pd.Timestamp(previous_date) + pd.Timedelta(seconds=3)
            filled_dates.append(current_date)
            methods.append("이전 행 + 3초 보정")
            previous_date = current_date
        else:
            filled_dates.append(pd.NaT)
            methods.append("보정 실패")

    return pd.Series(filled_dates, index=parsed_dates.index), pd.Series(methods, index=parsed_dates.index)


def detect_best_date_column(df: pd.DataFrame) -> Optional[Any]:
    """컬럼명이 없을 때 날짜처럼 변환되는 비율이 높은 컬럼을 찾는다."""
    best_column = None
    best_score = 0.0
    for column in df.columns[: min(len(df.columns), 20)]:
        sample = df[column].dropna().head(80)
        if sample.empty:
            continue
        parsed = parse_date_series(sample)
        score = parsed.notna().mean()
        if score > best_score:
            best_column = column
            best_score = score
    if best_score >= 0.45:
        return best_column
    return None


def to_numeric_series(series: pd.Series) -> pd.Series:
    text_series = series.astype("string").str.strip().str.replace(",", "", regex=False)
    text_series = text_series.replace({"": pd.NA, "-": pd.NA, "nan": pd.NA, "None": pd.NA})
    return pd.to_numeric(text_series, errors="coerce")


def format_column_option(df: pd.DataFrame, column: Any) -> str:
    position = list(df.columns).index(column)
    label = clean_text(column) if not isinstance(column, int) else f"컬럼 {position + 1}"
    return f"{excel_column_name(position)}열 - {label}"


def default_column_by_index(df: pd.DataFrame, index: int) -> Optional[Any]:
    if df.shape[1] > index:
        return df.columns[index]
    return None


def extract_fixed_columns(df: pd.DataFrame, date_col: Optional[Any] = None) -> pd.DataFrame:
    """
    백데이터에서 고정 위치 기준으로 필요한 컬럼을 추출한다.
    BS열: 날짜, B열: 시리얼넘버, D열: 원본 형명, BQ열: 감액량
    """
    if df.shape[1] <= REDUCTION_INDEX:
        raise ValueError("백데이터에 BQ열 감액량 데이터가 없습니다.")

    if date_col is None:
        date_col = default_column_by_index(df, DATE_INDEX)

    extracted = pd.DataFrame(
        {
            "날짜 원본": df[date_col] if date_col is not None else pd.Series([pd.NaT] * len(df)),
            "시리얼넘버": df.iloc[:, SERIAL_INDEX] if df.shape[1] > SERIAL_INDEX else pd.NA,
            "원본 형명": df.iloc[:, MODEL_INDEX] if df.shape[1] > MODEL_INDEX else pd.NA,
            "원본 감액량": df.iloc[:, REDUCTION_INDEX],
        }
    )
    return extracted


def extract_manual_columns(df: pd.DataFrame, column_map: Dict[str, Any]) -> pd.DataFrame:
    """사용자가 선택한 컬럼 기준으로 분석에 필요한 값을 추출한다."""
    return pd.DataFrame(
        {
            "날짜 원본": df[column_map["date"]],
            "시리얼넘버": df[column_map["serial"]],
            "원본 형명": df[column_map["model"]],
            "원본 감액량": df[column_map["reduction"]],
        }
    )


def generate_model_from_serial(serial: Any) -> Optional[str]:
    """
    시리얼넘버를 이용해 AGM 형명을 자동 생성한다.
    예:
    080C24K010206011H2 -> AGM80_H3
    105C24D170022502H3 -> AGM105_H3
    105A24D170022502H3 -> AGM105_H1
    """
    model, _ = generate_model_from_serial_with_reason(serial)
    return model


def generate_model_from_serial_with_reason(serial: Any) -> Tuple[Optional[str], str]:
    serial_text = clean_text(serial)
    if not serial_text:
        return None, "시리얼넘버 없음"
    if len(serial_text) < 4:
        return None, "시리얼넘버 길이 부족"

    prefix = serial_text[:3]
    if not prefix.isdigit():
        return None, "앞 3자리 숫자 아님"

    fourth_char = serial_text[3].upper()
    suffix_map = {"A": "1", "C": "3"}
    if fourth_char not in suffix_map:
        return None, "네 번째 문자가 A/C 아님"

    second_from_end = serial_text[-2].upper()
    return f"AGM{int(prefix)}_{second_from_end}{suffix_map[fourth_char]}", ""


def preprocess_backdata(df: pd.DataFrame) -> pd.DataFrame:
    """
    시리얼넘버, 원본 형명, 보정 형명, 감액량을 전처리한다.
    D열 형명이 비어 있으면 시리얼넘버 기반으로 보정 형명을 생성한다.
    """
    processed = df.copy()
    processed["시리얼넘버"] = processed["시리얼넘버"].apply(clean_text)
    processed["원본 형명"] = processed["원본 형명"].apply(clean_text)

    corrected_models: List[str] = []
    model_methods: List[str] = []
    model_errors: List[str] = []

    for original_model, serial in zip(processed["원본 형명"], processed["시리얼넘버"]):
        if not is_blank(original_model):
            corrected_models.append(clean_text(original_model))
            model_methods.append("D열 사용")
            model_errors.append("")
            continue

        generated_model, reason = generate_model_from_serial_with_reason(serial)
        if generated_model:
            corrected_models.append(generated_model)
            model_methods.append("시리얼넘버 자동 생성")
            model_errors.append("")
        else:
            corrected_models.append("")
            model_methods.append("생성 실패")
            model_errors.append(reason)

    processed["보정 형명"] = corrected_models
    processed["제품 형명"] = processed["보정 형명"]
    processed["형명 생성 방식"] = model_methods
    processed["형명 생성 오류"] = model_errors
    processed["형명 생성 실패"] = processed["형명 생성 방식"].eq("생성 실패")

    parsed_dates = parse_date_series(processed["날짜 원본"])
    processed["날짜 원본 변환 실패"] = parsed_dates.isna()
    processed["날짜"], processed["날짜 보정 방식"] = fill_invalid_dates_with_previous_plus_seconds(parsed_dates)
    processed["날짜 변환 실패"] = processed["날짜"].isna()

    processed["감액량"] = to_numeric_series(processed["원본 감액량"])
    processed["감액량 변환 실패"] = processed["감액량"].isna()
    return processed


def preprocess_data(df: pd.DataFrame) -> pd.DataFrame:
    """기존 함수명 호환을 위한 백데이터 전처리 래퍼."""
    return preprocess_backdata(df)


def load_standard_data(
    standard_uploaded_file: Optional[Any] = None,
    backdata_file: Optional[Any] = None,
    standard_sheet_name: Optional[str] = None,
) -> Tuple[Optional[pd.DataFrame], str, Optional[str]]:
    """
    기준값 데이터를 읽는다.
    우선순위: 별도 업로드 파일 -> 백데이터 내 기준 시트 -> app/data/standard.xlsx
    """
    standard_bytes = get_excel_bytes(standard_uploaded_file)
    if standard_bytes is not None:
        file_name = get_excel_file_name(standard_uploaded_file)
        sheet_name = standard_sheet_name or get_sheet_names(standard_bytes, file_name)[0]
        return load_excel_file(standard_bytes, sheet_name=sheet_name, header=0, file_name=file_name), "별도 기준값 Excel 파일", sheet_name

    backdata_bytes = get_excel_bytes(backdata_file)
    if backdata_bytes is not None:
        file_name = get_excel_file_name(backdata_file)
        sheet_names = get_sheet_names(backdata_bytes, file_name)
        sheet_name = find_standard_sheet(sheet_names)
        if sheet_name:
            return load_excel_file(backdata_bytes, sheet_name=sheet_name, header=0, file_name=file_name), f"백데이터 내 '{sheet_name}' 시트", sheet_name

    for standard_path in default_standard_paths():
        if not standard_path.exists():
            continue
        file_bytes = standard_path.read_bytes()
        file_name = str(standard_path)
        sheet_name = get_sheet_names(file_bytes, file_name)[0]
        return load_excel_file(file_bytes, sheet_name=sheet_name, header=0, file_name=file_name), f"기본 기준값 파일: {standard_path}", sheet_name

    return None, "기준값 없음", None


def normalize_standard_data(df: pd.DataFrame, column_map: Dict[str, Any]) -> pd.DataFrame:
    """기준값 컬럼명을 제품 형명, 중심치, UCL, LCL로 표준화한다."""
    standard = pd.DataFrame(
        {
            "제품 형명": df[column_map["model"]].apply(clean_text),
            "중심치": to_numeric_series(df[column_map["center"]]),
            "UCL": to_numeric_series(df[column_map["ucl"]]),
            "LCL": to_numeric_series(df[column_map["lcl"]]),
        }
    )
    standard = standard[~standard["제품 형명"].apply(is_blank)].copy()
    standard = standard[~(standard["중심치"].isna() & standard["UCL"].isna() & standard["LCL"].isna())].copy()
    standard["기준값 있음"] = True
    standard["UCL/LCL 누락"] = standard["UCL"].isna() | standard["LCL"].isna()
    standard["LCL>UCL 오류"] = standard["LCL"].notna() & standard["UCL"].notna() & (standard["LCL"] > standard["UCL"])
    return standard


def fixed_standard_column_map(df: pd.DataFrame) -> Dict[str, Any]:
    """기준값 Excel의 고정 위치(A/E/F/G)를 표준 컬럼 매핑으로 변환한다."""
    if df.shape[1] <= STANDARD_LCL_INDEX:
        raise ValueError("기준값 Excel 파일에 G열 하한 규격 데이터가 없습니다.")

    return {
        "model": df.columns[STANDARD_MODEL_INDEX],
        "ucl": df.columns[STANDARD_UCL_INDEX],
        "center": df.columns[STANDARD_CENTER_INDEX],
        "lcl": df.columns[STANDARD_LCL_INDEX],
    }


def merge_standard(data: pd.DataFrame, standard: Optional[pd.DataFrame]) -> pd.DataFrame:
    """보정 형명을 기준으로 백데이터와 기준값을 병합한다."""
    data = data.copy()
    data["제품 형명"] = data["보정 형명"]

    if standard is None or standard.empty:
        data["중심치"] = np.nan
        data["UCL"] = np.nan
        data["LCL"] = np.nan
        data["기준값 있음"] = False
        data["UCL/LCL 누락"] = False
        data["LCL>UCL 오류"] = False
        return data

    standard_columns = ["제품 형명", "중심치", "UCL", "LCL", "기준값 있음", "UCL/LCL 누락", "LCL>UCL 오류"]
    merged = data.merge(standard[standard_columns], on="제품 형명", how="left")
    merged["기준값 있음"] = merged["기준값 있음"].fillna(False).astype(bool)
    merged["UCL/LCL 누락"] = merged["UCL/LCL 누락"].fillna(False).astype(bool)
    merged["LCL>UCL 오류"] = merged["LCL>UCL 오류"].fillna(False).astype(bool)
    return merged


def judge_result(df: pd.DataFrame) -> pd.DataFrame:
    """감액량이 LCL~UCL 범위 안에 있는지 판정한다."""
    judged = df.copy()
    judged["판정 가능"] = (
        ~judged["형명 생성 실패"]
        & ~judged["감액량 변환 실패"]
        & ~judged["날짜 변환 실패"]
        & judged["기준값 있음"]
        & judged["UCL"].notna()
        & judged["LCL"].notna()
        & (judged["LCL"] <= judged["UCL"])
    )

    judged["판정 결과"] = "판정 제외"
    judged.loc[judged["판정 가능"] & (judged["감액량"] < judged["LCL"]), "판정 결과"] = "부적합"
    judged.loc[judged["판정 가능"] & (judged["감액량"] > judged["UCL"]), "판정 결과"] = "부적합"
    judged.loc[
        judged["판정 가능"] & judged["감액량"].between(judged["LCL"], judged["UCL"], inclusive="both"),
        "판정 결과",
    ] = "합격"

    judged["이탈 방향"] = ""
    judged.loc[judged["판정 가능"] & (judged["감액량"] > judged["UCL"]), "이탈 방향"] = "상한 초과"
    judged.loc[judged["판정 가능"] & (judged["감액량"] < judged["LCL"]), "이탈 방향"] = "하한 미달"

    reasons = []
    for _, row in judged.iterrows():
        if row["판정 가능"]:
            reasons.append("")
        elif row["형명 생성 실패"]:
            reasons.append("형명 생성 실패")
        elif row["날짜 변환 실패"]:
            reasons.append("날짜 변환 실패")
        elif row["감액량 변환 실패"]:
            reasons.append("감액량 변환 실패")
        elif not row["기준값 있음"]:
            reasons.append("기준값 없음")
        elif row["UCL/LCL 누락"]:
            reasons.append("UCL/LCL 누락")
        elif row["LCL>UCL 오류"]:
            reasons.append("LCL>UCL 오류")
        else:
            reasons.append("판정 제외")
    judged["판정 제외 사유"] = reasons
    return judged


def denominator_mask(df: pd.DataFrame, denominator_option: str) -> pd.Series:
    if denominator_option.startswith("기준값이 있는"):
        return df["판정 가능"]
    return ~df["형명 생성 실패"] & ~df["감액량 변환 실패"] & ~df["날짜 변환 실패"]


def make_kpi(df: pd.DataFrame, denominator_option: str) -> Dict[str, Any]:
    """필터 적용 후 전체 KPI를 계산한다."""
    denominator = denominator_mask(df, denominator_option)
    total = int(denominator.sum())
    fail_count = int((df["판정 결과"] == "부적합").sum())
    pass_count = int((df["판정 결과"] == "합격").sum())
    excluded_count = int((df["판정 결과"] == "판정 제외").sum())
    fail_rate = fail_count / total * 100 if total else np.nan
    average_loss = df.loc[denominator, "감액량"].mean() if total else np.nan

    return {
        "total": total,
        "fail_count": fail_count,
        "pass_count": pass_count,
        "excluded_count": excluded_count,
        "fail_rate": fail_rate,
        "average_loss": average_loss,
    }


def make_summary(df: pd.DataFrame, denominator_option: str) -> pd.DataFrame:
    """제품 형명별 품질 현황을 집계한다."""
    if df.empty:
        return pd.DataFrame()

    rows: List[Dict[str, Any]] = []
    denominator = denominator_mask(df, denominator_option)

    for model, group in df.groupby("제품 형명", dropna=False):
        group_denominator = denominator.loc[group.index]
        total = int(group_denominator.sum())
        fail_count = int((group["판정 결과"] == "부적합").sum())
        pass_count = int((group["판정 결과"] == "합격").sum())
        excluded_count = int((group["판정 결과"] == "판정 제외").sum())
        defect_rate = fail_count / total * 100 if total else np.nan

        valid_loss = group["감액량"].dropna()
        if not group["기준값 있음"].any():
            standard_status = "기준값 없음"
        elif group["LCL>UCL 오류"].any():
            standard_status = "LCL>UCL 오류"
        elif group["UCL/LCL 누락"].any():
            standard_status = "UCL/LCL 누락"
        else:
            standard_status = "정상"

        rows.append(
            {
                "제품 형명": model,
                "총 댓수": total,
                "합격 수량": pass_count,
                "부적합 수량": fail_count,
                "부적합률": defect_rate,
                "평균 감액량": valid_loss.mean() if not valid_loss.empty else np.nan,
                "최소 감액량": valid_loss.min() if not valid_loss.empty else np.nan,
                "최대 감액량": valid_loss.max() if not valid_loss.empty else np.nan,
                "표준편차": valid_loss.std(ddof=1) if len(valid_loss) > 1 else 0.0 if len(valid_loss) == 1 else np.nan,
                "중심치": group["중심치"].dropna().iloc[0] if group["중심치"].notna().any() else np.nan,
                "UCL": group["UCL"].dropna().iloc[0] if group["UCL"].notna().any() else np.nan,
                "LCL": group["LCL"].dropna().iloc[0] if group["LCL"].notna().any() else np.nan,
                "판정 제외 수": excluded_count,
                "전체 데이터 수": int(len(group)),
                "기준 상태": standard_status,
            }
        )

    summary = pd.DataFrame(rows)
    summary = summary.sort_values(["부적합률", "부적합 수량"], ascending=[False, False], na_position="last")
    return summary.reset_index(drop=True)


def make_daily_model_quality_summary(df: pd.DataFrame) -> pd.DataFrame:
    """제품 형명과 날짜별로 정상품, 상한/하한 이탈 수량과 비율을 집계한다."""
    if df.empty:
        return pd.DataFrame()

    working_df = df.copy()
    working_df["날짜"] = working_df["날짜"].dt.date

    def make_quality_row(group: pd.DataFrame, model: Any, date_value: Any) -> Dict[str, Any]:
        total_count = int(len(group))
        normal_count = int((group["판정 결과"] == "합격").sum())
        upper_count = int((group["이탈 방향"] == "상한 초과").sum())
        lower_count = int((group["이탈 방향"] == "하한 미달").sum())

        return {
            "제품 형명": model,
            "날짜": date_value,
            "전체 수량": total_count,
            "정상품 수량": normal_count,
            "정상품 비율": normal_count / total_count * 100 if total_count else np.nan,
            "상한 초과 수량": upper_count,
            "상한 초과 비율": upper_count / total_count * 100 if total_count else np.nan,
            "하한 초과 수량": lower_count,
            "하한 초과 비율": lower_count / total_count * 100 if total_count else np.nan,
        }

    rows: List[Dict[str, Any]] = []
    rows.append(make_quality_row(working_df, "전체 형명", "전체 날짜"))

    detail_rows: List[Dict[str, Any]] = []
    for (model, date_value), group in working_df.groupby(["제품 형명", "날짜"], dropna=False):
        detail_rows.append(make_quality_row(group, model, date_value))

    if detail_rows:
        rows.extend(pd.DataFrame(detail_rows).sort_values(["제품 형명", "날짜"]).to_dict("records"))

    return pd.DataFrame(rows).reset_index(drop=True)


def build_hover_data(df: pd.DataFrame) -> np.ndarray:
    return np.stack(
        [
            df["날짜"].dt.strftime("%Y-%m-%d").fillna("-"),
            df["제품 형명"].fillna("-"),
            df["감액량"].round(3).astype("string").fillna("-"),
            df["판정 결과"].fillna("-"),
            df["이탈 방향"].replace("", "-").fillna("-"),
            df["시리얼넘버"].replace("", "-").fillna("-"),
        ],
        axis=-1,
    )


def add_spec_segment(
    fig: go.Figure,
    x_values: List[Any],
    y_value: Any,
    name: str,
    color: str,
    dash: str,
    showlegend: bool,
    row: Optional[int] = None,
    col: Optional[int] = None,
) -> None:
    if pd.isna(y_value):
        return
    trace = go.Scatter(
        x=x_values,
        y=[y_value, y_value],
        mode="lines",
        name=name,
        legendgroup=name,
        showlegend=showlegend,
        line=dict(color=color, width=2, dash=dash),
        hoverinfo="skip",
    )
    if row is None or col is None:
        fig.add_trace(trace)
    else:
        fig.add_trace(trace, row=row, col=col)


def make_distribution_chart(df: pd.DataFrame) -> go.Figure:
    """형명별 감액량 분포 그래프를 만든다."""
    plot_df = df[df["감액량"].notna() & ~df["제품 형명"].apply(is_blank)].copy()
    fig = go.Figure()
    if plot_df.empty:
        fig.add_annotation(text="그래프에 표시할 데이터가 없습니다.", x=0.5, y=0.5, showarrow=False)
        return fig

    models = sorted(plot_df["제품 형명"].unique())
    model_to_x = {model: idx for idx, model in enumerate(models)}
    rng = np.random.default_rng(20240521)
    plot_df["_x"] = plot_df["제품 형명"].map(model_to_x).astype(float) + rng.uniform(-0.22, 0.22, size=len(plot_df))

    for result, result_df in plot_df.groupby("판정 결과"):
        fig.add_trace(
            go.Scatter(
                x=result_df["_x"],
                y=result_df["감액량"],
                mode="markers",
                name=result,
                marker=dict(
                    color=RESULT_COLORS.get(result, "#737373"),
                    size=8 if result != "부적합" else 10,
                    opacity=0.78,
                    line=dict(width=0.8, color="white"),
                    symbol="circle" if result != "부적합" else "x",
                ),
                customdata=build_hover_data(result_df),
                hovertemplate=(
                    "날짜: %{customdata[0]}<br>"
                    "제품 형명: %{customdata[1]}<br>"
                    "감액량: %{customdata[2]}<br>"
                    "판정 결과: %{customdata[3]}<br>"
                    "이탈 방향: %{customdata[4]}<br>"
                    "시리얼넘버: %{customdata[5]}<extra></extra>"
                ),
            )
        )

    legend_flags = {"UCL": True, "LCL": True, "중심치": True}
    for model in models:
        model_rows = plot_df[plot_df["제품 형명"] == model]
        standard_row = model_rows[model_rows["기준값 있음"]].head(1)
        if standard_row.empty:
            continue
        row = standard_row.iloc[0]
        x_center = model_to_x[model]
        x_values = [x_center - 0.36, x_center + 0.36]
        add_spec_segment(fig, x_values, row["UCL"], "UCL", "#ef4444", "dash", legend_flags["UCL"])
        add_spec_segment(fig, x_values, row["LCL"], "LCL", "#ef4444", "dash", legend_flags["LCL"])
        add_spec_segment(fig, x_values, row["중심치"], "중심치", "#16a34a", "dot", legend_flags["중심치"])
        legend_flags = {key: False for key in legend_flags}

    fig.update_layout(
        title="형명별 감액량 분포",
        xaxis=dict(title="제품 형명", tickmode="array", tickvals=list(model_to_x.values()), ticktext=models),
        yaxis=dict(title="감액량"),
        template="plotly_white",
        height=560,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(family="Malgun Gothic, Apple SD Gothic Neo, Arial"),
        margin=dict(l=40, r=30, t=90, b=80),
    )
    return fig


def make_trend_chart(df: pd.DataFrame) -> go.Figure:
    """날짜 기준 감액량 추이 그래프를 만든다."""
    plot_df = df[df["감액량"].notna() & df["날짜"].notna() & ~df["제품 형명"].apply(is_blank)].copy()
    if plot_df.empty:
        fig = go.Figure()
        fig.add_annotation(text="그래프에 표시할 데이터가 없습니다.", x=0.5, y=0.5, showarrow=False)
        return fig

    plot_df = plot_df.sort_values(["제품 형명", "날짜"])
    models = sorted(plot_df["제품 형명"].unique())
    x_min = plot_df["날짜"].min()
    x_max = plot_df["날짜"].max()
    if x_min == x_max:
        x_min = x_min - pd.Timedelta(days=1)
        x_max = x_max + pd.Timedelta(days=1)

    if len(models) == 1:
        fig = go.Figure()
        model_df = plot_df[plot_df["제품 형명"] == models[0]]
        for result, result_df in model_df.groupby("판정 결과"):
            fig.add_trace(
                go.Scatter(
                    x=result_df["날짜"],
                    y=result_df["감액량"],
                    mode="lines+markers",
                    name=result,
                    line=dict(color=RESULT_COLORS.get(result, "#737373"), width=1.5),
                    marker=dict(
                        color=RESULT_COLORS.get(result, "#737373"),
                        size=9 if result != "부적합" else 11,
                        symbol="circle" if result != "부적합" else "x",
                    ),
                    customdata=build_hover_data(result_df),
                    hovertemplate=(
                        "날짜: %{customdata[0]}<br>"
                        "제품 형명: %{customdata[1]}<br>"
                        "감액량: %{customdata[2]}<br>"
                        "판정 결과: %{customdata[3]}<br>"
                        "이탈 방향: %{customdata[4]}<br>"
                        "시리얼넘버: %{customdata[5]}<extra></extra>"
                    ),
                )
            )

        standard_row = model_df[model_df["기준값 있음"]].head(1)
        if not standard_row.empty:
            row = standard_row.iloc[0]
            add_spec_segment(fig, [x_min, x_max], row["UCL"], "UCL", "#ef4444", "dash", True)
            add_spec_segment(fig, [x_min, x_max], row["LCL"], "LCL", "#ef4444", "dash", True)
            add_spec_segment(fig, [x_min, x_max], row["중심치"], "중심치", "#16a34a", "dot", True)

        fig.update_layout(title=f"{models[0]} 날짜별 감액량 추이", height=540)
    else:
        fig = make_subplots(
            rows=len(models),
            cols=1,
            shared_xaxes=True,
            vertical_spacing=min(0.04, 0.18 / max(len(models), 1)),
            subplot_titles=models,
        )
        shown_legends: set[str] = set()
        for row_idx, model in enumerate(models, start=1):
            model_df = plot_df[plot_df["제품 형명"] == model]
            for result, result_df in model_df.groupby("판정 결과"):
                showlegend = result not in shown_legends
                shown_legends.add(result)
                fig.add_trace(
                    go.Scatter(
                        x=result_df["날짜"],
                        y=result_df["감액량"],
                        mode="markers+lines",
                        name=result,
                        legendgroup=result,
                        showlegend=showlegend,
                        line=dict(color=RESULT_COLORS.get(result, "#737373"), width=1),
                        marker=dict(
                            color=RESULT_COLORS.get(result, "#737373"),
                            size=7 if result != "부적합" else 9,
                            symbol="circle" if result != "부적합" else "x",
                        ),
                        customdata=build_hover_data(result_df),
                        hovertemplate=(
                            "날짜: %{customdata[0]}<br>"
                            "제품 형명: %{customdata[1]}<br>"
                            "감액량: %{customdata[2]}<br>"
                            "판정 결과: %{customdata[3]}<br>"
                            "이탈 방향: %{customdata[4]}<br>"
                            "시리얼넘버: %{customdata[5]}<extra></extra>"
                        ),
                    ),
                    row=row_idx,
                    col=1,
                )

            standard_row = model_df[model_df["기준값 있음"]].head(1)
            if not standard_row.empty:
                standard = standard_row.iloc[0]
                for name, color, dash, value in [
                    ("UCL", "#ef4444", "dash", standard["UCL"]),
                    ("LCL", "#ef4444", "dash", standard["LCL"]),
                    ("중심치", "#16a34a", "dot", standard["중심치"]),
                ]:
                    showlegend = name not in shown_legends
                    shown_legends.add(name)
                    add_spec_segment(fig, [x_min, x_max], value, name, color, dash, showlegend, row=row_idx, col=1)

            fig.update_yaxes(title_text="감액량", row=row_idx, col=1)

        fig.update_layout(title="날짜별 감액량 추이", height=max(520, 260 * len(models)))

    fig.update_layout(
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(family="Malgun Gothic, Apple SD Gothic Neo, Arial"),
        margin=dict(l=50, r=30, t=90, b=60),
    )
    fig.update_xaxes(title_text="날짜")
    fig.update_yaxes(title_text="감액량")
    return fig


def make_pass_rate_trend_chart(df: pd.DataFrame) -> go.Figure:
    """날짜별 합격률(합격 수량 / (합격 + 부적합)) 추이 그래프를 만든다.

    판정 제외 데이터는 합격률 분모에서 제외한다.
    형명이 하나면 단일 추이선, 여러 개면 형명별 추이선 + 전체 합격률 선을 함께 표시한다.
    """
    plot_df = df[
        df["날짜"].notna()
        & ~df["제품 형명"].apply(is_blank)
        & df["판정 결과"].isin(["합격", "부적합"])
    ].copy()

    fig = go.Figure()
    if plot_df.empty:
        fig.add_annotation(text="합격률을 계산할 판정 데이터가 없습니다.", x=0.5, y=0.5, showarrow=False)
        fig.update_layout(
            title="날짜별 합격률 추이",
            template="plotly_white",
            height=540,
            font=dict(family="Malgun Gothic, Apple SD Gothic Neo, Arial"),
        )
        return fig

    plot_df["판정일"] = plot_df["날짜"].dt.date

    def rate_frame(source: pd.DataFrame) -> pd.DataFrame:
        grouped = source.groupby("판정일")
        summary = pd.DataFrame(
            {
                "합격 수량": grouped.apply(lambda g: int((g["판정 결과"] == "합격").sum())),
                "판정 대상 수량": grouped.size().astype(int),
            }
        ).reset_index()
        summary["합격률"] = summary["합격 수량"] / summary["판정 대상 수량"] * 100
        return summary.sort_values("판정일")

    models = sorted(plot_df["제품 형명"].unique())

    if len(models) > 1:
        overall = rate_frame(plot_df)
        fig.add_trace(
            go.Scatter(
                x=overall["판정일"],
                y=overall["합격률"],
                mode="lines+markers",
                name="전체 합격률",
                line=dict(color="#111827", width=2.4),
                marker=dict(color="#111827", size=8),
                customdata=np.stack([overall["합격 수량"], overall["판정 대상 수량"]], axis=-1),
                hovertemplate=(
                    "날짜: %{x}<br>"
                    "합격률: %{y:.2f}%<br>"
                    "합격 수량: %{customdata[0]}<br>"
                    "판정 대상 수량: %{customdata[1]}<extra>전체</extra>"
                ),
            )
        )

    palette = ["#2563eb", "#16a34a", "#d97706", "#7c3aed", "#dc2626", "#0891b2", "#db2777", "#65a30d"]
    for index, model in enumerate(models):
        model_summary = rate_frame(plot_df[plot_df["제품 형명"] == model])
        color = palette[index % len(palette)]
        fig.add_trace(
            go.Scatter(
                x=model_summary["판정일"],
                y=model_summary["합격률"],
                mode="lines+markers",
                name=model,
                line=dict(color=color, width=1.6),
                marker=dict(color=color, size=7),
                customdata=np.stack(
                    [model_summary["합격 수량"], model_summary["판정 대상 수량"]], axis=-1
                ),
                hovertemplate=(
                    "날짜: %{x}<br>"
                    "제품 형명: " + str(model) + "<br>"
                    "합격률: %{y:.2f}%<br>"
                    "합격 수량: %{customdata[0]}<br>"
                    "판정 대상 수량: %{customdata[1]}<extra></extra>"
                ),
            )
        )

    fig.update_layout(
        title="날짜별 합격률 추이",
        xaxis=dict(title="날짜"),
        yaxis=dict(title="합격률 (%)", range=[0, 105], ticksuffix="%"),
        template="plotly_white",
        height=540,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(family="Malgun Gothic, Apple SD Gothic Neo, Arial"),
        margin=dict(l=50, r=30, t=90, b=60),
    )
    return fig


def safe_sheet_name(sheet_name: str) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", sheet_name)
    return cleaned[:31] or "Sheet1"


def convert_df_to_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """DataFrame을 한글 컬럼명이 유지되는 Excel 파일 bytes로 변환한다."""
    output = io.BytesIO()
    export_df = df.copy()
    for column in export_df.columns:
        if pd.api.types.is_datetime64_any_dtype(export_df[column]):
            export_df[column] = export_df[column].dt.strftime("%Y-%m-%d")

    sheet_name = safe_sheet_name(sheet_name)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.book[sheet_name]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))
            worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 28)
    return output.getvalue()


def format_number(value: Any, digits: int = 3) -> str:
    if pd.isna(value):
        return "-"
    return f"{float(value):.{digits}f}"


def style_summary(summary: pd.DataFrame) -> pd.io.formats.style.Styler:
    def highlight(row: pd.Series) -> List[str]:
        base_style = "color: #111827;"
        rate = row.get("부적합률")
        status = row.get("기준 상태")
        if status != "정상":
            return [f"{base_style} background-color: #fff7ed;"] * len(row)
        if pd.notna(rate) and rate >= 5:
            return [f"{base_style} background-color: #fef2f2;"] * len(row)
        if pd.notna(rate) and rate >= 1:
            return [f"{base_style} background-color: #fff7ed;"] * len(row)
        return [f"{base_style} background-color: #ffffff;"] * len(row)

    formatters = {
        "부적합률": lambda value: "-" if pd.isna(value) else f"{value:.2f}%",
        "평균 감액량": lambda value: format_number(value, 3),
        "최소 감액량": lambda value: format_number(value, 3),
        "최대 감액량": lambda value: format_number(value, 3),
        "표준편차": lambda value: format_number(value, 3),
        "중심치": lambda value: format_number(value, 3),
        "UCL": lambda value: format_number(value, 3),
        "LCL": lambda value: format_number(value, 3),
    }
    return summary.style.apply(highlight, axis=1).format(formatters)


def standard_duplicate_rows(standard: pd.DataFrame) -> pd.DataFrame:
    if standard.empty:
        return pd.DataFrame()
    duplicated = standard["제품 형명"].duplicated(keep=False)
    return standard[duplicated].sort_values("제품 형명")


def filter_valid_analysis_rows(judged: pd.DataFrame) -> pd.DataFrame:
    """형명, 날짜, 감액량 전처리에 성공한 데이터만 조회/분석 대상으로 둔다."""
    return judged[
        ~judged["형명 생성 실패"] & ~judged["날짜 변환 실패"] & ~judged["감액량 변환 실패"]
    ].copy()


def make_graph_download_data(df: pd.DataFrame) -> pd.DataFrame:
    """그래프에 표시되는 데이터를 시리얼넘버 순서와 지정 컬럼 순서로 정리한다."""
    columns = ["날짜", "시리얼넘버", "제품 형명", "감액량", "LCL", "중심치", "UCL", "판정 결과", "이탈 방향"]
    graph_df = df[df["감액량"].notna() & df["날짜"].notna() & ~df["제품 형명"].apply(is_blank)][columns].copy()
    for column in ["감액량", "LCL", "중심치", "UCL"]:
        graph_df[column] = graph_df[column].round(3)
    graph_df["_시리얼정렬키"] = graph_df["시리얼넘버"].apply(lambda value: clean_text(value).upper())
    graph_df = graph_df.sort_values(["_시리얼정렬키", "날짜", "제품 형명"], kind="stable")
    return graph_df.drop(columns=["_시리얼정렬키"]).reset_index(drop=True)


def render_kpi_cards(kpi: Dict[str, Any]) -> None:
    columns = st.columns(5)
    columns[0].metric("전체 총 댓수", f"{kpi['total']:,}")
    columns[1].metric("전체 부적합 수량", f"{kpi['fail_count']:,}")
    columns[2].metric("전체 합격 수량", f"{kpi['pass_count']:,}")
    columns[3].metric("전체 부적합률", "-" if pd.isna(kpi["fail_rate"]) else f"{kpi['fail_rate']:.2f}%")
    columns[4].metric("전체 평균 감액량", format_number(kpi["average_loss"], 3))

    if kpi["excluded_count"]:
        st.caption(f"판정 제외 데이터 {kpi['excluded_count']:,}건은 기준값 없음, UCL/LCL 누락, 형명/날짜/감액량 오류 등의 사유로 합격/부적합 판정에서 제외되었습니다.")


def render_excluded_data(processed: pd.DataFrame, filtered: pd.DataFrame) -> None:
    st.subheader("제외 데이터 확인")
    with st.expander("형명 생성 실패 / 감액량 변환 실패 / 날짜 변환 실패 / 기준값 없음 데이터", expanded=False):
        model_fail = processed[processed["형명 생성 실패"]][
            ["날짜 원본", "시리얼넘버", "원본 형명", "원본 감액량", "형명 생성 오류", "형명 생성 방식"]
        ]
        reduction_fail = processed[processed["감액량 변환 실패"]][
            ["날짜 원본", "시리얼넘버", "원본 형명", "보정 형명", "원본 감액량"]
        ]
        date_fail = processed[processed["날짜 변환 실패"]][
            ["날짜 원본", "날짜 보정 방식", "시리얼넘버", "원본 형명", "보정 형명", "원본 감액량"]
        ]
        no_standard = filtered[~filtered["기준값 있음"]][
            ["날짜", "시리얼넘버", "원본 형명", "보정 형명", "감액량", "판정 제외 사유"]
        ]

        st.write(f"형명 생성 실패: {len(model_fail):,}건")
        if not model_fail.empty:
            st.dataframe(model_fail, use_container_width=True, hide_index=True)

        st.write(f"감액량 변환 실패: {len(reduction_fail):,}건")
        if not reduction_fail.empty:
            st.dataframe(reduction_fail, use_container_width=True, hide_index=True)

        adjusted_dates = processed[processed["날짜 보정 방식"].eq("이전 행 + 3초 보정")][
            ["날짜 원본", "날짜", "날짜 보정 방식", "시리얼넘버", "원본 형명", "보정 형명"]
        ]
        st.write(f"날짜 자동 보정: {len(adjusted_dates):,}건")
        if not adjusted_dates.empty:
            st.dataframe(adjusted_dates, use_container_width=True, hide_index=True)

        st.write(f"날짜 변환 실패: {len(date_fail):,}건")
        if not date_fail.empty:
            st.dataframe(date_fail, use_container_width=True, hide_index=True)

        st.write(f"기준값 없는 데이터: {len(no_standard):,}건")
        if not no_standard.empty:
            st.dataframe(no_standard, use_container_width=True, hide_index=True)


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    st.markdown(
        """
        **분석 기준 안내**  
        감액량이 LCL 미만 또는 UCL 초과이면 부적합, LCL 이상 UCL 이하이면 합격으로 판정합니다.
        제품 형명은 기본적으로 D열을 사용하고, D열이 비어 있으면 B열 시리얼넘버로 보정 형명을 자동 생성합니다.
        """
    )

    st.sidebar.header("파일 업로드")
    backdata_file = st.sidebar.file_uploader("백데이터 Excel 업로드", type=["xlsx", "xlsm", "xls"])
    standard_file = st.sidebar.file_uploader("기준값 Excel 업로드(선택)", type=["xlsx", "xlsm", "xls"])

    st.sidebar.header("백데이터 읽기 설정")
    start_row = st.sidebar.number_input("백데이터 시작 행", min_value=1, max_value=50, value=FIXED_START_ROW, step=1)
    header_mode = st.sidebar.radio(
        "5행 처리 방식",
        ["자동 판단", "5행부터 실제 데이터", "5행을 컬럼명으로 사용"],
        index=0,
        help="기본 양식은 5행부터 실제 데이터입니다. 5행이 헤더인 파일도 자동으로 판단합니다.",
    )

    mapping_mode = st.sidebar.radio(
        "백데이터 컬럼 모드",
        ["고정 컬럼 사용", "수동 컬럼 선택"],
        index=0,
        help="기본값은 BS열 날짜, B열 시리얼넘버, D열 원본 형명, BQ열 감액량입니다.",
    )

    denominator_option = st.sidebar.radio(
        "부적합률 계산 기준",
        ["기준값이 있는 데이터만 판정 대상에 포함", "기준값 없는 데이터도 총 댓수에 포함"],
        index=0,
    )

    graph_mode = st.sidebar.selectbox(
        "그래프 모드",
        ["형명별 분포 그래프", "날짜 추이 그래프", "날짜별 합격률 추이"],
    )

    if backdata_file is None:
        st.info("왼쪽 사이드바에서 백데이터 Excel 파일을 업로드하면 분석을 시작합니다.")
        st.markdown(
            """
            기본 백데이터 위치 기준은 아래와 같습니다.
            - 데이터 시작 행: 5행
            - BS열: 날짜
            - B열: 시리얼넘버
            - D열: 원본 형명
            - BQ열: 감액량
            """
        )
        return

    try:
        backdata_bytes = get_excel_bytes(backdata_file)
        sheet_names = get_sheet_names(backdata_bytes, get_excel_file_name(backdata_file))
        data_sheet_candidates = [name for name in sheet_names if not is_standard_sheet_name(name)] or sheet_names
        data_sheet = st.sidebar.selectbox("백데이터 시트", data_sheet_candidates)

        raw_df, used_header_row = load_backdata_excel(backdata_file, data_sheet, int(start_row), header_mode)
        if raw_df.empty:
            st.warning("백데이터 시트에 분석할 데이터가 없습니다.")
            return

        detected_columns = detect_columns(raw_df, BACKDATA_CANDIDATES)
        fixed_date_default = default_column_by_index(raw_df, DATE_INDEX)
        if mapping_mode == "고정 컬럼 사용":
            date_default = fixed_date_default or detected_columns.get("date") or detect_best_date_column(raw_df)
        else:
            date_default = detected_columns.get("date") or detect_best_date_column(raw_df) or fixed_date_default

        with st.sidebar.expander("컬럼 매핑 설정", expanded=True):
            st.caption("고정 컬럼 모드에서도 날짜 컬럼은 파일 양식에 따라 선택할 수 있습니다.")
            column_options = list(raw_df.columns)
            date_col = st.selectbox(
                "날짜 컬럼",
                column_options,
                index=column_options.index(date_default) if date_default in column_options else 0,
                format_func=lambda column: format_column_option(raw_df, column),
            )

            if mapping_mode == "수동 컬럼 선택":
                serial_default = detected_columns.get("serial") or default_column_by_index(raw_df, SERIAL_INDEX)
                model_default = detected_columns.get("model") or default_column_by_index(raw_df, MODEL_INDEX)
                reduction_default = detected_columns.get("reduction") or default_column_by_index(raw_df, REDUCTION_INDEX)

                serial_col = st.selectbox(
                    "시리얼넘버 컬럼",
                    column_options,
                    index=column_options.index(serial_default) if serial_default in column_options else 0,
                    format_func=lambda column: format_column_option(raw_df, column),
                )
                model_col = st.selectbox(
                    "제품 형명 컬럼",
                    column_options,
                    index=column_options.index(model_default) if model_default in column_options else 0,
                    format_func=lambda column: format_column_option(raw_df, column),
                )
                reduction_col = st.selectbox(
                    "감액량 컬럼",
                    column_options,
                    index=column_options.index(reduction_default) if reduction_default in column_options else 0,
                    format_func=lambda column: format_column_option(raw_df, column),
                )
                extracted_df = extract_manual_columns(
                    raw_df,
                    {"date": date_col, "serial": serial_col, "model": model_col, "reduction": reduction_col},
                )
            else:
                st.caption("현재 적용: BS열 → 날짜, B열 → 시리얼넘버, D열 → 원본 형명, BQ열 → 감액량")
                extracted_df = extract_fixed_columns(raw_df, date_col=date_col)

        processed = preprocess_data(extracted_df)

        standard_sheet_name = None
        if standard_file is not None:
            standard_sheet_names = get_sheet_names(get_excel_bytes(standard_file), get_excel_file_name(standard_file))
            standard_sheet_name = st.sidebar.selectbox("기준값 시트", standard_sheet_names)

        standard_raw, standard_source, _ = load_standard_data(standard_file, backdata_file, standard_sheet_name)
        standard = None
        if standard_raw is not None and not standard_raw.empty:
            standard_detected = detect_columns(standard_raw, STANDARD_CANDIDATES)
            missing_standard_mapping = any(standard_detected.get(key) is None for key in ["model", "center", "ucl", "lcl"])

            standard_mapping_mode = st.sidebar.radio(
                "기준값 컬럼 모드",
                ["고정 컬럼 사용", "수동 컬럼 선택"],
                index=0,
                help="기본값은 A열=형명, E열=UCL(상한규격), F열=중심 규격, G열=LCL(하한규격)입니다.",
            )

            if standard_mapping_mode == "고정 컬럼 사용":
                st.sidebar.caption("현재 적용: A열 → 형명, E열 → UCL, F열 → 중심치, G열 → LCL")
                standard_map = fixed_standard_column_map(standard_raw)
            else:
                standard_map = {}

            with st.sidebar.expander("기준값 컬럼 매핑", expanded=missing_standard_mapping or standard_mapping_mode == "수동 컬럼 선택"):
                standard_columns = list(standard_raw.columns)
                if standard_mapping_mode == "고정 컬럼 사용":
                    st.write("고정 위치 기준으로 기준값을 읽습니다.")
                    st.write("- A열: 형명")
                    st.write("- E열: UCL(상한규격)")
                    st.write("- F열: 중심 규격")
                    st.write("- G열: LCL(하한규격)")
                else:
                    labels = {"model": "제품 형명", "center": "중심치", "ucl": "상한선/UCL", "lcl": "하한선/LCL"}
                    fixed_defaults = {
                        "model": default_column_by_index(standard_raw, STANDARD_MODEL_INDEX),
                        "ucl": default_column_by_index(standard_raw, STANDARD_UCL_INDEX),
                        "center": default_column_by_index(standard_raw, STANDARD_CENTER_INDEX),
                        "lcl": default_column_by_index(standard_raw, STANDARD_LCL_INDEX),
                    }
                    for key in ["model", "ucl", "center", "lcl"]:
                        default_col = standard_detected.get(key) or fixed_defaults.get(key) or (standard_columns[0] if standard_columns else None)
                        standard_map[key] = st.selectbox(
                            f"기준값 {labels[key]} 컬럼",
                            standard_columns,
                            index=standard_columns.index(default_col) if default_col in standard_columns else 0,
                            format_func=lambda column: format_column_option(standard_raw, column),
                            key=f"standard_{key}",
                        )
            standard = normalize_standard_data(standard_raw, standard_map)
            duplicate_standard = standard_duplicate_rows(standard)
            if not duplicate_standard.empty:
                st.error("동일한 제품 형명에 여러 기준값이 중복 등록되어 있습니다. 기준값 파일을 수정한 뒤 다시 업로드해 주세요.")
                st.dataframe(duplicate_standard, use_container_width=True, hide_index=True)
                return
        else:
            st.warning("기준값 파일 또는 기준정보 시트를 찾지 못했습니다. 기준값이 없는 데이터는 판정 제외로 표시됩니다.")

        st.caption(f"백데이터 시트: {data_sheet} / 5행 헤더 사용 여부: {'예' if used_header_row else '아니오'} / 기준값 출처: {standard_source}")

        merged = merge_standard(processed, standard)
        judged = judge_result(merged)

        standard_issue_rows = judged[
            judged["기준값 있음"] & (judged["UCL/LCL 누락"] | judged["LCL>UCL 오류"])
        ][["제품 형명", "중심치", "UCL", "LCL", "UCL/LCL 누락", "LCL>UCL 오류"]].drop_duplicates()

        if not standard_issue_rows.empty:
            st.error("UCL/LCL이 없거나 LCL이 UCL보다 큰 기준값이 있습니다. 해당 형명은 판정에서 제외됩니다.")
            st.dataframe(standard_issue_rows, use_container_width=True, hide_index=True)

        generation_fail_count = int(processed["형명 생성 실패"].sum())
        reduction_fail_count = int(processed["감액량 변환 실패"].sum())
        date_adjusted_count = int(processed["날짜 보정 방식"].eq("이전 행 + 3초 보정").sum())
        date_fail_count = int(processed["날짜 변환 실패"].sum())

        if generation_fail_count:
            st.warning(f"형명 생성 실패 데이터 {generation_fail_count:,}건은 기준값 병합 및 판정에서 제외됩니다.")
        if reduction_fail_count:
            st.warning(f"감액량이 비어 있거나 숫자로 변환되지 않은 데이터 {reduction_fail_count:,}건은 분석에서 제외됩니다.")
        if date_adjusted_count:
            st.info(f"날짜 변환 실패 데이터 {date_adjusted_count:,}건은 바로 위 행 날짜에 3초를 더해 자동 보정했습니다.")
        if date_fail_count:
            st.warning(f"날짜로 변환할 수 없는 데이터 {date_fail_count:,}건은 분석에서 제외됩니다.")

        analysis_base = filter_valid_analysis_rows(judged)
        if analysis_base.empty:
            st.warning("전처리 후 분석 가능한 데이터가 없습니다.")
            render_excluded_data(processed, judged)
            return

        min_date = analysis_base["날짜"].min().date()
        max_date = analysis_base["날짜"].max().date()

        st.sidebar.header("조회 필터")
        date_range = st.sidebar.date_input("날짜 범위", value=(min_date, max_date), min_value=min_date, max_value=max_date)
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
        else:
            start_date = end_date = date_range

        model_options = sorted(analysis_base["제품 형명"].dropna().unique().tolist())
        selected_models = st.sidebar.multiselect(
            "제품 형명 선택",
            model_options,
            default=model_options,
            help="기본값은 전체 형명입니다. 한 형명만 선택하면 요약표, 그래프, 부적합 상세 목록이 모두 해당 형명만 표시됩니다.",
        )

        if not selected_models:
            st.warning("조회 결과 없음")
            st.info("사이드바에서 제품 형명을 하나 이상 선택해 주세요.")
            render_excluded_data(processed, judged)
            return

        filtered = analysis_base[
            (analysis_base["날짜"].dt.date >= start_date)
            & (analysis_base["날짜"].dt.date <= end_date)
            & (analysis_base["제품 형명"].isin(selected_models))
        ].copy()

        if filtered.empty:
            st.warning("조회 결과 없음")
            render_excluded_data(processed, judged)
            return

        no_standard_models = filtered.loc[~filtered["기준값 있음"], "제품 형명"].dropna().unique().tolist()
        if no_standard_models:
            st.warning("기준값이 없는 제품 형명: " + ", ".join(sorted(no_standard_models)))

        st.subheader("전체 KPI")
        kpi = make_kpi(filtered, denominator_option)
        render_kpi_cards(kpi)

        st.subheader("제품 형명별 요약표")
        summary = make_summary(filtered, denominator_option)
        st.dataframe(style_summary(summary), use_container_width=True, hide_index=True)

        summary_excel = convert_df_to_excel(summary, "형명별 요약")
        filtered_export_columns = [
            "날짜",
            "날짜 원본",
            "날짜 보정 방식",
            "시리얼넘버",
            "원본 형명",
            "보정 형명",
            "형명 생성 방식",
            "원본 감액량",
            "감액량",
            "중심치",
            "UCL",
            "LCL",
            "판정 결과",
            "이탈 방향",
            "판정 제외 사유",
        ]
        filtered_excel = convert_df_to_excel(filtered[filtered_export_columns], "필터데이터")

        download_cols = st.columns(2)
        download_cols[0].download_button(
            "형명별 요약 Excel 다운로드",
            data=summary_excel,
            file_name="AGM_감액량_형명별_요약.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        download_cols[1].download_button(
            "필터 데이터 Excel 다운로드",
            data=filtered_excel,
            file_name="AGM_감액량_필터데이터.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("감액량 분석 그래프")
        if graph_mode == "형명별 분포 그래프":
            figure = make_distribution_chart(filtered)
        elif graph_mode == "날짜별 합격률 추이":
            figure = make_pass_rate_trend_chart(filtered)
        else:
            figure = make_trend_chart(filtered)
        st.plotly_chart(figure, use_container_width=True)

        graph_download_data = make_graph_download_data(filtered)
        graph_excel = convert_df_to_excel(graph_download_data, "그래프 데이터")
        st.download_button(
            "그래프 데이터 Excel 다운로드",
            data=graph_excel,
            file_name="AGM_감액량_그래프데이터.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("부적합 데이터 상세 목록")
        defect_columns = ["날짜", "제품 형명", "감액량", "LCL", "UCL", "중심치", "판정 결과", "이탈 방향"]
        defects = filtered[filtered["판정 결과"] == "부적합"][defect_columns].copy()
        st.caption(f"부적합 상세 데이터 {len(defects):,}건은 화면에 표시하지 않고 Excel 다운로드로 제공합니다.")
        defect_excel = convert_df_to_excel(defects, "부적합 상세")
        st.download_button(
            "부적합 상세 Excel 다운로드",
            data=defect_excel,
            file_name="AGM_감액량_부적합_상세.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("형명별 날짜별 품질 요약")
        daily_model_summary = make_daily_model_quality_summary(filtered)
        st.caption("하한 초과 수량은 감액량이 LCL보다 낮은 하한 이탈 데이터를 의미합니다.")
        st.dataframe(
            daily_model_summary,
            use_container_width=True,
            hide_index=True,
            column_config={
                "정상품 비율": st.column_config.NumberColumn(format="%.2f%%"),
                "상한 초과 비율": st.column_config.NumberColumn(format="%.2f%%"),
                "하한 초과 비율": st.column_config.NumberColumn(format="%.2f%%"),
            },
        )

        render_excluded_data(processed, judged)

    except ValueError as error:
        st.error(str(error))
    except Exception as error:  # noqa: BLE001
        st.error("처리 중 오류가 발생했습니다. Excel 파일 형식과 컬럼 매핑을 확인해 주세요.")
        st.exception(error)


if __name__ == "__main__":
    main()
