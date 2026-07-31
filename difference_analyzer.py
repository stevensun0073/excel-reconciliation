from __future__ import annotations

import time
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Sequence

from models import Record


MAX_ITEMS = 10
TIME_LIMIT_SECONDS = 120.0


class SearchTimeoutError(Exception):
    """差额组合搜索超过时间限制。"""


@dataclass(frozen=True)
class DifferenceItem:
    """差额分析中的一条未匹配记录。"""

    source_sheet: str
    source_row: int
    original_amount: Decimal
    search_amount: Decimal


@dataclass
class DifferenceAnalysisResult:
    """差额分析结果。"""

    sheet1_unmatched_count: int
    sheet2_unmatched_count: int
    items: list[DifferenceItem]

    signed_difference: Decimal
    target_difference: Decimal

    max_items: int
    time_limit_seconds: float

    found: bool = False
    timed_out: bool = False
    minimum_item_count: int | None = None

    selected_items: list[DifferenceItem] = field(
        default_factory=list
    )

    elapsed_seconds: float = 0.0

    @property
    def selected_total(self) -> Decimal:
        """候选组合的搜索金额合计。"""

        return sum(
            (
                item.search_amount
                for item in self.selected_items
            ),
            Decimal("0"),
        )


def amount_to_cents(amount: Decimal) -> int:
    """把 Decimal 金额转换成整数分。"""

    return int(
        (amount * Decimal("100")).quantize(
            Decimal("1")
        )
    )


def build_difference_items(
    sheet1_records: Sequence[Record],
    sheet2_records: Sequence[Record],
) -> list[DifferenceItem]:
    """
    合并两张表中的未匹配记录。

    Sheet1：
    搜索金额保持原金额。

    Sheet2：
    搜索金额乘以 -1。
    """

    items: list[DifferenceItem] = []

    for record in sheet1_records:
        if record.matched:
            continue

        items.append(
            DifferenceItem(
                source_sheet="Sheet1",
                source_row=record.row,
                original_amount=record.amount,
                search_amount=record.amount,
            )
        )

    for record in sheet2_records:
        if record.matched:
            continue

        items.append(
            DifferenceItem(
                source_sheet="Sheet2",
                source_row=record.row,
                original_amount=record.amount,
                search_amount=-record.amount,
            )
        )

    return items


def build_exact_sum_bounds(
    values: list[int],
    max_items: int,
) -> tuple[list[list[int | None]], list[list[int | None]]]:
    """
    计算从每个位置开始，恰好选择指定笔数时：

    - 能取得的最小合计；
    - 能取得的最大合计。

    用于安全剪枝，不会漏掉正确答案。
    """

    item_count = len(values)

    minimum = [
        [None] * (max_items + 1)
        for _ in range(item_count + 1)
    ]

    maximum = [
        [None] * (max_items + 1)
        for _ in range(item_count + 1)
    ]

    minimum[item_count][0] = 0
    maximum[item_count][0] = 0

    for index in range(item_count - 1, -1, -1):
        minimum[index][0] = 0
        maximum[index][0] = 0

        available = item_count - index
        largest_count = min(max_items, available)

        for count in range(1, largest_count + 1):
            min_candidates: list[int] = []
            max_candidates: list[int] = []

            skip_min = minimum[index + 1][count]
            skip_max = maximum[index + 1][count]

            if skip_min is not None:
                min_candidates.append(skip_min)

            if skip_max is not None:
                max_candidates.append(skip_max)

            take_previous_min = minimum[index + 1][count - 1]
            take_previous_max = maximum[index + 1][count - 1]

            if take_previous_min is not None:
                min_candidates.append(
                    values[index] + take_previous_min
                )

            if take_previous_max is not None:
                max_candidates.append(
                    values[index] + take_previous_max
                )

            if min_candidates:
                minimum[index][count] = min(
                    min_candidates
                )

            if max_candidates:
                maximum[index][count] = max(
                    max_candidates
                )

    return minimum, maximum


def find_exact_subset_for_count(
    values: list[int],
    target: int,
    required_count: int,
    minimum_bounds: list[list[int | None]],
    maximum_bounds: list[list[int | None]],
    deadline: float,
) -> list[int] | None:
    """
    寻找恰好 required_count 笔、合计等于 target 的组合。

    返回原 values 中的下标；找不到则返回 None。
    """

    item_count = len(values)
    failed_states: set[tuple[int, int, int]] = set()

    def search(
        index: int,
        remaining_count: int,
        current_sum: int,
    ) -> tuple[int, ...] | None:
        if time.monotonic() >= deadline:
            raise SearchTimeoutError

        if remaining_count == 0:
            if current_sum == target:
                return ()

            return None

        if item_count - index < remaining_count:
            return None

        if index >= item_count:
            return None

        minimum_remaining = minimum_bounds[index][
            remaining_count
        ]
        maximum_remaining = maximum_bounds[index][
            remaining_count
        ]

        if (
            minimum_remaining is None
            or maximum_remaining is None
        ):
            return None

        needed = target - current_sum

        if (
            needed < minimum_remaining
            or needed > maximum_remaining
        ):
            return None

        state = (
            index,
            remaining_count,
            current_sum,
        )

        if state in failed_states:
            return None

        value = values[index]

        take_result = search(
            index=index + 1,
            remaining_count=remaining_count - 1,
            current_sum=current_sum + value,
        )

        if take_result is not None:
            return (index,) + take_result

        skip_result = search(
            index=index + 1,
            remaining_count=remaining_count,
            current_sum=current_sum,
        )

        if skip_result is not None:
            return skip_result

        failed_states.add(state)
        return None

    result = search(
        index=0,
        remaining_count=required_count,
        current_sum=0,
    )

    if result is None:
        return None

    return list(result)


def find_minimum_subset(
    items: list[DifferenceItem],
    signed_target: Decimal,
    max_items: int,
    time_limit_seconds: float,
) -> tuple[list[DifferenceItem], bool, float]:
    """
    按1笔、2笔……依次搜索。

    第一次找到的组合就是最少笔数组合。
    """

    start_time = time.monotonic()
    deadline = start_time + time_limit_seconds

    indexed_values = [
        (
            original_index,
            amount_to_cents(item.search_amount),
        )
        for original_index, item in enumerate(items)
    ]

    # 绝对金额较大的记录优先，有助于更快找到答案。
    indexed_values.sort(
        key=lambda pair: abs(pair[1]),
        reverse=True,
    )

    sorted_original_indexes = [
        pair[0]
        for pair in indexed_values
    ]

    values = [
        pair[1]
        for pair in indexed_values
    ]

    target_cents = amount_to_cents(signed_target)

    minimum_bounds, maximum_bounds = (
        build_exact_sum_bounds(
            values=values,
            max_items=max_items,
        )
    )

    try:
        for required_count in range(
            1,
            max_items + 1,
        ):
            selected_sorted_indexes = (
                find_exact_subset_for_count(
                    values=values,
                    target=target_cents,
                    required_count=required_count,
                    minimum_bounds=minimum_bounds,
                    maximum_bounds=maximum_bounds,
                    deadline=deadline,
                )
            )

            if selected_sorted_indexes is None:
                continue

            selected_original_indexes = [
                sorted_original_indexes[index]
                for index in selected_sorted_indexes
            ]

            selected_items = [
                items[index]
                for index in selected_original_indexes
            ]

            elapsed = time.monotonic() - start_time

            return selected_items, False, elapsed

    except SearchTimeoutError:
        elapsed = time.monotonic() - start_time
        return [], True, elapsed

    elapsed = time.monotonic() - start_time
    return [], False, elapsed


def analyze_difference(
    sheet1_records: Sequence[Record],
    sheet2_records: Sequence[Record],
    max_items: int = MAX_ITEMS,
    time_limit_seconds: float = TIME_LIMIT_SECONDS,
) -> DifferenceAnalysisResult:
    """分析剩余差额并寻找最少笔数组合。"""

    if max_items < 1:
        raise ValueError("max_items必须至少为1。")

    if time_limit_seconds <= 0:
        raise ValueError(
            "time_limit_seconds必须大于0。"
        )

    items = build_difference_items(
        sheet1_records,
        sheet2_records,
    )

    sheet1_unmatched_count = sum(
        1
        for item in items
        if item.source_sheet == "Sheet1"
    )

    sheet2_unmatched_count = sum(
        1
        for item in items
        if item.source_sheet == "Sheet2"
    )

    signed_difference = sum(
        (
            item.search_amount
            for item in items
        ),
        Decimal("0"),
    )

    target_difference = abs(signed_difference)

    result = DifferenceAnalysisResult(
        sheet1_unmatched_count=sheet1_unmatched_count,
        sheet2_unmatched_count=sheet2_unmatched_count,
        items=items,
        signed_difference=signed_difference,
        target_difference=target_difference,
        max_items=max_items,
        time_limit_seconds=time_limit_seconds,
    )

    if signed_difference == Decimal("0"):
        result.found = True
        result.minimum_item_count = 0
        return result

    (
        selected_items,
        timed_out,
        elapsed_seconds,
    ) = find_minimum_subset(
        items=items,
        signed_target=signed_difference,
        max_items=max_items,
        time_limit_seconds=time_limit_seconds,
    )

    result.selected_items = selected_items
    result.timed_out = timed_out
    result.elapsed_seconds = elapsed_seconds

    if selected_items:
        result.found = True
        result.minimum_item_count = len(
            selected_items
        )

    return result


def print_difference_summary(
    result: DifferenceAnalysisResult,
) -> None:
    """在终端打印差额分析结果。"""

    print()
    print("=" * 56)
    print("Difference Analysis")
    print("=" * 56)

    print(
        f"Unmatched Sheet1 rows : "
        f"{result.sheet1_unmatched_count}"
    )

    print(
        f"Unmatched Sheet2 rows : "
        f"{result.sheet2_unmatched_count}"
    )

    print(
        f"Combined analysis rows : "
        f"{len(result.items)}"
    )

    print(
        f"Target difference      : "
        f"{result.target_difference:.2f}"
    )

    print(
        f"Maximum items          : "
        f"{result.max_items}"
    )

    print(
        f"Search time limit      : "
        f"{result.time_limit_seconds:.0f} seconds"
    )

    print("-" * 56)

    if result.timed_out:
        print(
            "Search status          : "
            "TIME LIMIT REACHED"
        )

        print(
            "Result                 : "
            "No guaranteed minimum combination"
        )

    elif not result.found:
        print(
            "Search status          : "
            "COMPLETED"
        )

        print(
            "Result                 : "
            f"No exact combination within "
            f"{result.max_items} items"
        )

    elif result.minimum_item_count == 0:
        print(
            "Search status          : "
            "COMPLETED"
        )

        print(
            "Result                 : "
            "Remaining difference is zero"
        )

    else:
        print(
            "Search status          : "
            "COMPLETED"
        )

        print(
            f"Minimum combination    : "
            f"{result.minimum_item_count} item(s)"
        )

        print(
            f"Candidate signed total : "
            f"{result.selected_total:.2f}"
        )

        print()

        for item in result.selected_items:
            print(
                f"{item.source_sheet:<8} "
                f"Row {item.source_row:<6} "
                f"Original {item.original_amount:>12.2f} "
                f"Search {item.search_amount:>12.2f}"
            )

    print(
        f"Analysis time          : "
        f"{result.elapsed_seconds:.3f} seconds"
    )

    print("=" * 56)