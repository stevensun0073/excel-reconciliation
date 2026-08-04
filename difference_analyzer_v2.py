from __future__ import annotations

"""
difference_analyzer_v2.py

独立测试版 Difference Analyzer。

功能：
- 不覆盖原 difference_analyzer.py
- 不修改 Excel
- 不修改正式匹配状态
- 最多输出 5 组精确候选
- 优先笔数最少
- 最多 10 条记录
- 最长搜索 120 秒
"""

from dataclasses import dataclass
from decimal import Decimal
from time import perf_counter

from excel_io import load_excel
from final_yellow_matcher import match_final_yellow_records
from keyword_difference_matcher import match_keyword_differences
from matcher import Matcher


INPUT_FILE = "data.xlsx"
MAX_ITEMS = 10
MAX_CANDIDATES = 5
TIME_LIMIT_SECONDS = 120


@dataclass(frozen=True)
class AnalysisItem:
    source_sheet: str
    source_row: int
    original_amount: Decimal
    signed_amount: Decimal
    key_word: str


@dataclass(frozen=True)
class CandidateGroup:
    items: tuple[AnalysisItem, ...]
    signed_total: Decimal
    item_count: int
    sheet_count: int
    blank_keyword_count: int
    distinct_keyword_count: int
    row_span: int


def get_record_keyword(record) -> str:
    for key, value in record.extra.items():
        if str(key).strip().casefold() in {"key word", "keyword"}:
            return "" if value is None else str(value).strip()

    return ""


def build_analysis_items(
    sheet1_records,
    sheet2_records,
) -> list[AnalysisItem]:
    items: list[AnalysisItem] = []

    for record in sheet1_records:
        if record.matched:
            continue

        items.append(
            AnalysisItem(
                source_sheet="Sheet1",
                source_row=record.row,
                original_amount=record.amount,
                signed_amount=record.amount,
                key_word=get_record_keyword(record),
            )
        )

    for record in sheet2_records:
        if record.matched:
            continue

        items.append(
            AnalysisItem(
                source_sheet="Sheet2",
                source_row=record.row,
                original_amount=record.amount,
                signed_amount=-record.amount,
                key_word=get_record_keyword(record),
            )
        )

    return items


def calculate_target_difference(
    sheet1_records,
    sheet2_records,
) -> Decimal:
    sheet1_total = sum(
        (
            record.amount
            for record in sheet1_records
            if not record.matched
        ),
        Decimal("0"),
    )

    sheet2_total = sum(
        (
            record.amount
            for record in sheet2_records
            if not record.matched
        ),
        Decimal("0"),
    )

    return sheet1_total - sheet2_total


def candidate_identity(
    items: tuple[AnalysisItem, ...],
) -> tuple[tuple[str, int], ...]:
    return tuple(
        sorted(
            (
                item.source_sheet,
                item.source_row,
            )
            for item in items
        )
    )


def build_candidate_group(
    items: tuple[AnalysisItem, ...],
) -> CandidateGroup:
    sheets = {
        item.source_sheet
        for item in items
    }

    blank_keyword_count = sum(
        1
        for item in items
        if not item.key_word
    )

    nonblank_keywords = {
        item.key_word.casefold()
        for item in items
        if item.key_word
    }

    rows = [
        item.source_row
        for item in items
    ]

    return CandidateGroup(
        items=items,
        signed_total=sum(
            (
                item.signed_amount
                for item in items
            ),
            Decimal("0"),
        ),
        item_count=len(items),
        sheet_count=len(sheets),
        blank_keyword_count=blank_keyword_count,
        distinct_keyword_count=len(nonblank_keywords),
        row_span=(
            max(rows) - min(rows)
            if rows
            else 0
        ),
    )


def candidate_sort_key(
    candidate: CandidateGroup,
):
    stable_rows = tuple(
        sorted(
            (
                item.source_sheet,
                item.source_row,
            )
            for item in candidate.items
        )
    )

    return (
        candidate.item_count,
        candidate.sheet_count,
        -candidate.blank_keyword_count,
        candidate.distinct_keyword_count,
        candidate.row_span,
        stable_rows,
    )


def find_candidates(
    items: list[AnalysisItem],
    target: Decimal,
    max_items: int = MAX_ITEMS,
    max_candidates: int = MAX_CANDIDATES,
    time_limit_seconds: int = TIME_LIMIT_SECONDS,
) -> tuple[list[CandidateGroup], bool, float]:
    start_time = perf_counter()
    deadline = start_time + time_limit_seconds

    sorted_items = sorted(
        items,
        key=lambda item: (
            abs(item.signed_amount),
            item.source_sheet,
            item.source_row,
        ),
    )

    item_count = len(sorted_items)
    found: list[CandidateGroup] = []
    seen: set[tuple[tuple[str, int], ...]] = set()
    timed_out = False

    for group_size in range(
        1,
        min(max_items, item_count) + 1,
    ):
        selected_indexes: list[int] = []

        def search(
            start_index: int,
            remaining_count: int,
            current_sum: Decimal,
        ) -> bool:
            nonlocal timed_out

            if perf_counter() >= deadline:
                timed_out = True
                return True

            if len(found) >= max_candidates:
                return True

            if remaining_count == 0:
                if current_sum != target:
                    return False

                selected_items = tuple(
                    sorted_items[index]
                    for index in selected_indexes
                )

                identity = candidate_identity(
                    selected_items
                )

                if identity in seen:
                    return False

                seen.add(identity)
                found.append(
                    build_candidate_group(
                        selected_items
                    )
                )

                return len(found) >= max_candidates

            if item_count - start_index < remaining_count:
                return False

            final_start = item_count - remaining_count

            for index in range(
                start_index,
                final_start + 1,
            ):
                selected_indexes.append(index)

                should_stop = search(
                    start_index=index + 1,
                    remaining_count=remaining_count - 1,
                    current_sum=(
                        current_sum
                        + sorted_items[index].signed_amount
                    ),
                )

                selected_indexes.pop()

                if should_stop:
                    return True

            return False

        search(
            start_index=0,
            remaining_count=group_size,
            current_sum=Decimal("0"),
        )

        if timed_out or len(found) >= max_candidates:
            break

    found.sort(
        key=candidate_sort_key
    )

    return (
        found[:max_candidates],
        timed_out,
        perf_counter() - start_time,
    )


def format_keyword_summary(
    candidate: CandidateGroup,
) -> str:
    keywords: list[str] = []
    seen: set[str] = set()

    for item in candidate.items:
        if not item.key_word:
            continue

        normalized = item.key_word.casefold()

        if normalized in seen:
            continue

        seen.add(normalized)
        keywords.append(item.key_word)

    return "; ".join(keywords) if keywords else "All blank"


def describe_candidate(
    candidate: CandidateGroup,
) -> str:
    reasons = [
        "single sheet"
        if candidate.sheet_count == 1
        else "mixed sheets"
    ]

    if candidate.blank_keyword_count:
        reasons.append(
            f"{candidate.blank_keyword_count} blank keyword"
        )

    if candidate.distinct_keyword_count == 0:
        reasons.append("no nonblank keyword")
    elif candidate.distinct_keyword_count == 1:
        reasons.append("one keyword group")
    else:
        reasons.append(
            f"{candidate.distinct_keyword_count} keyword groups"
        )

    return "; ".join(reasons)


def print_candidate(
    number: int,
    candidate: CandidateGroup,
) -> None:
    print()
    print("-" * 76)
    print(f"Candidate {number}")
    print("-" * 76)
    print(f"Items           : {candidate.item_count}")
    print(f"Sheets involved : {candidate.sheet_count}")
    print(f"Signed total    : {candidate.signed_total}")
    print(f"Keywords        : {format_keyword_summary(candidate)}")
    print(f"Ranking reason  : {describe_candidate(candidate)}")
    print()

    for item in sorted(
        candidate.items,
        key=lambda value: (
            value.source_sheet,
            value.source_row,
        ),
    ):
        keyword_display = (
            item.key_word
            if item.key_word
            else "<blank>"
        )

        print(
            f"{item.source_sheet:<8} "
            f"Row {item.source_row:<6} "
            f"Original {item.original_amount:>14} "
            f"Search {item.signed_amount:>14} "
            f"Key word: {keyword_display}"
        )


def run_existing_matching(
    sheet1_records,
    sheet2_records,
) -> None:
    matcher = Matcher(
        sheet1_records,
        sheet2_records,
    )
    matcher.run()

    keyword_groups = match_keyword_differences(
        sheet1_records=sheet1_records,
        sheet2_records=sheet2_records,
    )

    final_results = match_final_yellow_records(
        sheet1_records=sheet1_records,
        sheet2_records=sheet2_records,
    )

    print()
    print(
        "Keyword-Difference groups: "
        f"{keyword_groups}"
    )
    print(
        "Final Yellow groups      : "
        f"{sum(final_results.values())}"
    )


def main() -> None:
    print("=" * 76)
    print("Difference Analyzer V2 — Multiple Candidate Test")
    print("=" * 76)

    _, sheet1_records, sheet2_records = load_excel(
        INPUT_FILE
    )

    print(f"Input file     : {INPUT_FILE}")
    print(f"Sheet1 records : {len(sheet1_records)}")
    print(f"Sheet2 records : {len(sheet2_records)}")

    run_existing_matching(
        sheet1_records=sheet1_records,
        sheet2_records=sheet2_records,
    )

    target = calculate_target_difference(
        sheet1_records=sheet1_records,
        sheet2_records=sheet2_records,
    )

    items = build_analysis_items(
        sheet1_records=sheet1_records,
        sheet2_records=sheet2_records,
    )

    print()
    print("=" * 76)
    print("Difference Analysis V2")
    print("=" * 76)
    print(
        f"Unmatched Sheet1 : "
        f"{sum(1 for record in sheet1_records if not record.matched)}"
    )
    print(
        f"Unmatched Sheet2 : "
        f"{sum(1 for record in sheet2_records if not record.matched)}"
    )
    print(f"Combined rows    : {len(items)}")
    print(f"Target difference: {target}")
    print(f"Maximum items    : {MAX_ITEMS}")
    print(f"Maximum groups   : {MAX_CANDIDATES}")
    print(f"Time limit       : {TIME_LIMIT_SECONDS} seconds")

    candidates, timed_out, elapsed = find_candidates(
        items=items,
        target=target,
    )

    print()
    print("=" * 76)
    print("Candidate Summary")
    print("=" * 76)
    print(f"Candidate groups found: {len(candidates)}")
    print(
        f"Search status          : "
        f"{'TIME LIMIT' if timed_out else 'COMPLETED'}"
    )
    print(f"Analysis time          : {elapsed:.3f} seconds")

    if not candidates:
        print()
        print(
            "No exact candidate combination was found "
            "within the current limits."
        )
        return

    for number, candidate in enumerate(
        candidates,
        start=1,
    ):
        print_candidate(
            number=number,
            candidate=candidate,
        )

    print()
    print("=" * 76)
    print(
        "Analysis-only test: no Excel file and no official "
        "match result was changed."
    )
    print("=" * 76)


if __name__ == "__main__":
    main()