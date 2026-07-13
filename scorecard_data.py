"""Extract System Scorecard table blocks from 2026 - GSE Scorecards.xlsx."""

from __future__ import annotations

import logging
import re

import pandas as pd

logger = logging.getLogger(__name__)

SECTION_START_PATTERNS = (
    (re.compile(r"safety.*security", re.I), "Safety & Security"),
    (re.compile(r"customer\s*experience", re.I), "Customer Experience"),
    (re.compile(r"^operations\b|\boperations\s*\(", re.I), "Operations"),
    (re.compile(r"\bfinance\b", re.I), "Finance"),
    (re.compile(r"\bpeople\b", re.I), "People"),
)

SUMMARY_1_STOP_PATTERNS = (
    re.compile(r"\bfinance\b", re.I),
    re.compile(r"\bpeople\b", re.I),
    re.compile(r"overall\s+system", re.I),
    re.compile(r"total\s+system", re.I),
)

FINANCE_START_PATTERN = re.compile(r"\bfinance\b", re.I)
OPPORTUNITIES_PATTERN = re.compile(r"opportunities", re.I)
OVERALL_TOTAL_SCORE_PATTERN = re.compile(r"overall\s+total\s+score", re.I)

SCORECARD_BLOCKS = {
    "summary_1": "Safety & Security, Customer Experience, Operations (slide 3)",
    "summary_2": "Finance, People, overall total score, opportunities (slide 4)",
}

MONTH_HEADERS = tuple(
    name.upper()
    for name in (
        "JAN",
        "FEB",
        "MAR",
        "APR",
        "MAY",
        "JUN",
        "JUL",
        "AUG",
        "SEP",
        "SEPT",
        "OCT",
        "NOV",
        "DEC",
        "YTD",
        "YE",
    )
)

SCORECARD_ROW_TYPES = {"plan", "actual", "percent", "score", "total score"}


def _row_text(row: pd.Series) -> str:
    parts = [str(v).strip() for v in row.tolist() if v is not None and str(v).strip() and str(v).lower() != "nan"]
    return " ".join(parts)


def _detect_section(row_text: str) -> str | None:
    for pattern, name in SECTION_START_PATTERNS:
        if pattern.search(row_text):
            return name
    return None


def _trim_scorecard_block(block: pd.DataFrame) -> pd.DataFrame:
    while len(block.columns) > 0 and block.iloc[:, -1].isna().all():
        block = block.iloc[:, :-1]
    while len(block.columns) > 0 and block.iloc[:, 0].isna().all():
        block = block.iloc[:, 1:]
    return block.reset_index(drop=True)


def _should_stop_summary_1(row_text: str, seen_operations: bool) -> bool:
    if not seen_operations:
        return False
    for pattern in SUMMARY_1_STOP_PATTERNS:
        if pattern.search(row_text):
            return True
    return False


def find_system_scorecard_sheet(sheet_names: list[str]) -> str | None:
    """Pick the sheet most likely to hold the system scorecard."""
    priorities = (
        "system scorecard",
        "system",
        "scorecard",
        "summary",
    )
    lower_map = {name.lower(): name for name in sheet_names}
    for key in priorities:
        for lower, original in lower_map.items():
            if key in lower:
                return original
    return sheet_names[0] if sheet_names else None


def extract_system_scorecard_block(df: pd.DataFrame) -> pd.DataFrame:
    """Return Excel rows for red + blue + yellow sections (Safety, CX, Operations)."""
    if df.empty:
        return df

    start_row: int | None = None
    end_row: int | None = None
    seen_operations = False

    for idx in range(len(df)):
        text = _row_text(df.iloc[idx])
        if not text:
            continue

        section = _detect_section(text)
        if section == "Safety & Security" and start_row is None:
            start_row = idx
        if section == "Operations":
            seen_operations = True

        if start_row is not None and _should_stop_summary_1(text, seen_operations):
            end_row = idx - 1
            break

    if start_row is None:
        logger.warning("System scorecard summary_1: Safety & Security section not found in sheet")
        return pd.DataFrame()

    if end_row is None:
        end_row = len(df) - 1
        for idx in range(start_row + 1, len(df)):
            text = _row_text(df.iloc[idx])
            if seen_operations and _should_stop_summary_1(text, True):
                end_row = idx - 1
                break

    block = _trim_scorecard_block(df.iloc[start_row : end_row + 1].copy())
    logger.info(
        "System scorecard summary_1 block: rows %s-%s (%s rows x %s cols)",
        start_row,
        end_row,
        len(block),
        len(block.columns),
    )
    return block


def extract_finance_people_scorecard_block(df: pd.DataFrame) -> pd.DataFrame:
    """Return Excel rows for green + light blue sections plus footer rows (slide 4)."""
    if df.empty:
        return df

    start_row: int | None = None
    end_row: int | None = None

    for idx in range(len(df)):
        text = _row_text(df.iloc[idx])
        if not text:
            continue
        if start_row is None and FINANCE_START_PATTERN.search(text):
            start_row = idx
        if start_row is not None and OPPORTUNITIES_PATTERN.search(text):
            end_row = idx
            break

    if start_row is None:
        logger.warning("System scorecard summary_2: Finance section not found in sheet")
        return pd.DataFrame()

    if end_row is None:
        seen_people = False
        for idx in range(start_row + 1, len(df)):
            text = _row_text(df.iloc[idx])
            if not text:
                continue
            if re.search(r"\bpeople\b", text, re.I):
                seen_people = True
            if seen_people and OPPORTUNITIES_PATTERN.search(text):
                end_row = idx
                break
            if seen_people and OVERALL_TOTAL_SCORE_PATTERN.search(text):
                for follow in range(idx, min(idx + 5, len(df))):
                    follow_text = _row_text(df.iloc[follow])
                    if OPPORTUNITIES_PATTERN.search(follow_text):
                        end_row = follow
                        break
                if end_row is not None:
                    break

    if end_row is None:
        end_row = len(df) - 1
        for idx in range(start_row + 1, len(df)):
            text = _row_text(df.iloc[idx])
            if OVERALL_TOTAL_SCORE_PATTERN.search(text):
                for follow in range(idx, min(idx + 6, len(df))):
                    if OPPORTUNITIES_PATTERN.search(_row_text(df.iloc[follow])):
                        end_row = follow
                        break
                if end_row != len(df) - 1:
                    break

    block = _trim_scorecard_block(df.iloc[start_row : end_row + 1].copy())
    text = block.to_string()
    if "Finance" not in text:
        logger.warning("System scorecard summary_2: Finance rows missing from extracted block")
    if "People" not in text:
        logger.warning("System scorecard summary_2: People rows missing from extracted block")
    if not OPPORTUNITIES_PATTERN.search(text):
        logger.warning("System scorecard summary_2: OPPORTUNITIES footer row not found in block")

    logger.info(
        "System scorecard summary_2 block: rows %s-%s (%s rows x %s cols)",
        start_row,
        end_row,
        len(block),
        len(block.columns),
    )
    return block


def extract_scorecard_block(df: pd.DataFrame, *, block: str = "summary_1") -> pd.DataFrame:
    if block == "summary_2":
        return extract_finance_people_scorecard_block(df)
    return extract_system_scorecard_block(df)


def load_system_scorecard(
    data,
    *,
    workbook: str = "scorecards",
    sheet_name: str | None = None,
    block: str = "summary_1",
) -> pd.DataFrame:
    """Load a system scorecard grid from the scorecards workbook."""
    names = data.sheet_names(workbook)
    if not names:
        logger.warning("Scorecards workbook has no sheets")
        return pd.DataFrame()

    sheet = sheet_name or find_system_scorecard_sheet(names)
    if not sheet:
        return pd.DataFrame()

    raw = data.store.read_sheet(workbook, sheet, raw=True)
    if raw.empty:
        return raw

    return extract_scorecard_block(raw, block=block)


def format_scorecard_cell(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return f"{value:.2f}"
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if text.lower() in ("nan", "none"):
        return ""
    return text


def _normalize_header(text: str) -> str:
    cleaned = text.strip().upper()
    if cleaned == "SEPT":
        return "SEP"
    return cleaned


def find_scorecard_header_row(df: pd.DataFrame, *, search_rows: int = 12) -> int | None:
    """Locate the row containing KPI / month column headers."""
    for idx in range(min(search_rows, len(df))):
        row_vals = [format_scorecard_cell(df.iat[idx, c]) for c in range(len(df.columns))]
        upper = {_normalize_header(v) for v in row_vals if v}
        if "KPI" in upper and ("WEIGHT" in upper or "JAN" in upper):
            return idx
    return None


def scorecard_column_map(df: pd.DataFrame, header_row: int, table_headers: list[str]) -> dict[int, int]:
    """Map Excel column index -> PPT column index using header labels."""
    table_norm = {_normalize_header(h): idx for idx, h in enumerate(table_headers) if h.strip()}
    mapping: dict[int, int] = {}
    for excel_col in range(len(df.columns)):
        header = _normalize_header(format_scorecard_cell(df.iat[header_row, excel_col]))
        if not header:
            continue
        if header in table_norm:
            mapping[excel_col] = table_norm[header]
            continue
        for label, table_col in table_norm.items():
            if header in label or label in header:
                mapping[excel_col] = table_col
                break
    return mapping


def scorecard_row_key(row_values: list[str], *, label_cols: int = 4) -> str:
    """Build a normalized match key from the label portion of a scorecard row."""
    parts: list[str] = []
    for value in row_values[:label_cols]:
        text = value.strip()
        if not text:
            continue
        lowered = text.lower()
        if lowered in SCORECARD_ROW_TYPES:
            parts.append(lowered)
        elif "%" in text or re.search(r"\(\d", text):
            parts.append(re.sub(r"\s+", " ", lowered))
        else:
            parts.append(re.sub(r"\s+", " ", lowered))
    return "|".join(parts)


def contextual_scorecard_key(
    row_values: list[str],
    *,
    current_section: str,
    current_kpi: str,
) -> tuple[str, str, str]:
    """Return updated section/kpi plus a stable lookup key for one scorecard row."""
    joined = " ".join(v.strip() for v in row_values if v and v.strip())
    section = current_section
    kpi = current_kpi

    detected = _detect_section(joined)
    if detected:
        section = detected
        return section, kpi, f"{section}|section".lower()

    weight = row_values[1].strip() if len(row_values) > 1 else ""
    label = row_values[2].strip() if len(row_values) > 2 else ""
    lowered = label.lower()

    if weight.upper() == "WEIGHT" and label.upper() == "KPI":
        return section, kpi, ""

    if lowered in {"plan", "actual", "percent", "score"}:
        return section, kpi, f"{kpi}|{lowered}".lower()

    if lowered == "total score":
        return section, kpi, f"{section}|total score".lower()

    if OVERALL_TOTAL_SCORE_PATTERN.search(label):
        return section, kpi, "overall total score"

    if OPPORTUNITIES_PATTERN.search(label):
        return section, kpi, "opportunities"

    if label and lowered != "kpi":
        kpi = label
    elif weight and "%" in weight and label:
        kpi = label

    return section, kpi, ""


def build_scorecard_row_index(df: pd.DataFrame, *, header_row: int | None = None) -> dict[str, int]:
    """Index extracted scorecard rows by section/KPI context for table matching."""
    index: dict[str, int] = {}
    section = ""
    kpi = ""

    for row_idx in range(len(df)):
        values = [format_scorecard_cell(df.iat[row_idx, c]) for c in range(len(df.columns))]
        if not any(values):
            continue
        section, kpi, key = contextual_scorecard_key(values, current_section=section, current_kpi=kpi)
        if key and not key.endswith("|section"):
            index[key] = row_idx

    return index


def find_scorecard_table_header_row(table) -> int:
    """Locate the header row inside a scorecard PPT table."""
    for row_idx in range(min(8, len(table.rows))):
        headers = [table.cell(row_idx, col).text.strip() for col in range(len(table.columns))]
        upper = {_normalize_header(h) for h in headers if h}
        if "KPI" in upper and ("WEIGHT" in upper or "JAN" in upper):
            return row_idx
    return 0
