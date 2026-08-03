"""
prepare_data.py

Daily preparation step for the bank reconciliation project.

Run this script BEFORE main.py.

Inputs
------
1. company.xlsx
   - Uses the first worksheet.
   - Reads the column whose header is "Company Keyword".

2. data.xlsx
   - Processes Sheet1 only.
   - Reads the column whose header is "文本".
   - Adds or updates a new rightmost column named "Key word".

Matching rules
--------------
1. Normal Company Keyword
   - Case-insensitive.
   - Must appear as a complete word or complete keyword combination.
   - Example: ARM matches "PAYMENT TO ARM", but not FARM or ARMY.

2. KZ
   - KZ or KASZON must be separately distinguishable.
   - They may appear as standalone words or after separators such as -, /, _.
   - Examples matched: KZ, ABC-KZ, KASZON, ABC/KASZON.
   - Examples not matched: XYZKZABC, KASZONIC.
   - Output is always KZ.

3. GR
   - The following complete parts are recognized:
       GR
       GREAT RESOURCES
       ACMV
       E
       EL
       FP
       PS
   - They may be standalone or separated by spaces, -, /, _.
   - Examples matched: GR, ABC-GR, GREAT RESOURCES, PROJECT - ACMV, PROJECT FP.
   - Examples not matched: GRAPH, ELITE, FPSYSTEM.
   - Output is always GR.

4. Multiple matches
   - Unique keywords are joined with "; ".

This script does NOT process Sheet2 and does NOT run main.py.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Pattern

from openpyxl import load_workbook


COMPANY_FILE = Path("company.xlsx")
DATA_FILE = Path("data.xlsx")

COMPANY_KEYWORD_HEADER = "Company Keyword"
TEXT_HEADER = "文本"
OUTPUT_HEADER = "Key word"

HEADER_SEARCH_ROWS = 5


def normalize_spaces(value: str) -> str:
    """Collapse repeated whitespace and trim both ends."""
    return re.sub(r"\s+", " ", value).strip()


def require_file(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(
            f"File not found: {path}\n"
            "Please place it in the project folder and check the filename casing."
        )


def find_header(
    worksheet,
    header_name: str,
    search_rows: int = HEADER_SEARCH_ROWS,
) -> tuple[int, int]:
    """Return (header_row, column_number)."""
    last_row = min(search_rows, worksheet.max_row)

    for row in range(1, last_row + 1):
        for column in range(1, worksheet.max_column + 1):
            value = worksheet.cell(row, column).value
            if value is None:
                continue

            if str(value).strip().casefold() == header_name.casefold():
                return row, column

    raise ValueError(
        f'Cannot find header "{header_name}" in the first {last_row} rows '
        f'of worksheet "{worksheet.title}".'
    )


def find_or_create_output_column(
    worksheet,
    header_row: int,
    header_name: str,
) -> int:
    """
    Reuse an existing output column if present.
    Otherwise add it to the current right side of Sheet1.
    """
    for column in range(1, worksheet.max_column + 1):
        value = worksheet.cell(header_row, column).value
        if value is None:
            continue

        if str(value).strip().casefold() == header_name.casefold():
            return column

    output_column = worksheet.max_column + 1
    worksheet.cell(header_row, output_column).value = header_name
    return output_column


def contains_chinese(text: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def build_normal_keyword_pattern(keyword: str) -> Pattern[str]:
    """
    Build a case-insensitive complete-word / complete-combination pattern.

    Spaces inside a multi-word keyword may be one or more whitespace characters.
    Alphanumeric characters are not allowed immediately before or after it.
    """
    parts = keyword.split()

    if len(parts) > 1:
        body = r"\s+".join(re.escape(part) for part in parts)
    else:
        body = re.escape(keyword)

    return re.compile(
        rf"(?<![A-Za-z0-9]){body}(?![A-Za-z0-9])",
        re.IGNORECASE,
    )


# KZ must be independently distinguishable.
KZ_PATTERN = re.compile(
    r"(?<![A-Za-z0-9])(?:KZ|KASZON)(?![A-Za-z0-9])",
    re.IGNORECASE,
)

# GR special phrases/codes must also be complete parts.
GR_PATTERN = re.compile(
    r"(?<![A-Za-z0-9])"
    r"(?:GREAT\s+RESOURCES|GR|ACMV|E|EL|FP|PS)"
    r"(?![A-Za-z0-9])",
    re.IGNORECASE,
)


def load_company_keywords() -> tuple[list[tuple[str, Pattern[str] | None]], int]:
    """
    Load unique Company Keyword values from company.xlsx.

    KZ and GR are excluded from the normal keyword list because they use
    dedicated business rules.
    """
    require_file(COMPANY_FILE)

    workbook = load_workbook(
        COMPANY_FILE,
        data_only=True,
        read_only=True,
    )
    worksheet = workbook[workbook.sheetnames[0]]

    header_row, keyword_column = find_header(
        worksheet,
        COMPANY_KEYWORD_HEADER,
    )

    unique_keywords: list[str] = []
    seen: set[str] = set()

    for row in range(header_row + 1, worksheet.max_row + 1):
        value = worksheet.cell(row, keyword_column).value

        if value is None:
            continue

        keyword = normalize_spaces(str(value))
        if not keyword:
            continue

        normalized = keyword.casefold()

        if normalized in {"kz", "gr"}:
            continue

        if normalized in seen:
            continue

        seen.add(normalized)
        unique_keywords.append(keyword)

    # Longer phrases first. This does not suppress shorter matches, but it keeps
    # the output order more useful when a text contains both.
    unique_keywords.sort(key=lambda item: (-len(item), item.casefold()))

    compiled: list[tuple[str, Pattern[str] | None]] = []

    for keyword in unique_keywords:
        if contains_chinese(keyword):
            compiled.append((keyword, None))
        else:
            compiled.append((keyword, build_normal_keyword_pattern(keyword)))

    return compiled, header_row


def unique_in_order(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()

    for value in values:
        normalized = value.casefold()
        if normalized in seen:
            continue

        seen.add(normalized)
        result.append(value)

    return result


def match_text(
    text: str,
    normal_keywords: list[tuple[str, Pattern[str] | None]],
) -> list[str]:
    """Return all unique matched keywords for one Sheet1 text cell."""
    matches: list[str] = []

    # Special business rules first.
    if KZ_PATTERN.search(text):
        matches.append("KZ")

    if GR_PATTERN.search(text):
        matches.append("GR")

    # Normal Company Keyword rules.
    for keyword, pattern in normal_keywords:
        if pattern is None:
            # Chinese keywords such as 零星 / 中国 are matched as literal text.
            if keyword in text:
                matches.append(keyword)
        elif pattern.search(text):
            matches.append(keyword)

    return unique_in_order(matches)


def main() -> None:
    require_file(DATA_FILE)

    print("Loading Company Keyword master data...")
    normal_keywords, _ = load_company_keywords()
    print(f"Loaded {len(normal_keywords)} unique normal keywords.")
    print("Special rules enabled: KZ, GR")

    print("Opening data.xlsx...")
    workbook = load_workbook(DATA_FILE)

    if "Sheet1" not in workbook.sheetnames:
        raise ValueError('data.xlsx does not contain a worksheet named "Sheet1".')

    worksheet = workbook["Sheet1"]

    header_row, text_column = find_header(worksheet, TEXT_HEADER)
    output_column = find_or_create_output_column(
        worksheet,
        header_row,
        OUTPUT_HEADER,
    )

    # Clear old Key word results so rerunning the script is safe.
    for row in range(header_row + 1, worksheet.max_row + 1):
        worksheet.cell(row, output_column).value = None

    total_rows = worksheet.max_row - header_row
    matched_rows = 0
    unmatched_rows = 0
    empty_rows = 0
    multiple_match_rows = 0

    print(f"Processing {total_rows} Sheet1 rows...")

    for row in range(header_row + 1, worksheet.max_row + 1):
        processed = row - header_row

        if processed % 200 == 0:
            print(f"  Processed {processed}/{total_rows} rows...")

        raw_text = worksheet.cell(row, text_column).value

        if raw_text is None or not str(raw_text).strip():
            empty_rows += 1
            continue

        text = normalize_spaces(str(raw_text))
        matches = match_text(text, normal_keywords)

        if not matches:
            unmatched_rows += 1
            continue

        worksheet.cell(row, output_column).value = "; ".join(matches)
        matched_rows += 1

        if len(matches) > 1:
            multiple_match_rows += 1

    print("Saving data.xlsx...")
    workbook.save(DATA_FILE)

    print("=" * 60)
    print("Sheet1 Key word generation finished")
    print("=" * 60)
    print(f"Normal keywords loaded : {len(normal_keywords)}")
    print(f"Rows processed         : {total_rows}")
    print(f"Matched rows           : {matched_rows}")
    print(f"Unmatched rows         : {unmatched_rows}")
    print(f"Empty text rows        : {empty_rows}")
    print(f"Multiple-match rows    : {multiple_match_rows}")
    print(f"Output column          : {output_column}")
    print(f"Saved file             : {DATA_FILE.resolve()}")
    print("=" * 60)


if __name__ == "__main__":
    main()