from __future__ import annotations

from decimal import Decimal

from matcher import amount_sign, sum_records


MATCH_TYPE_PREFIX = "Final Yellow Match"
MAX_ONE_TO_MANY_SIZE = 6


def build_match_type(
    left_size: int,
    right_size: int,
) -> str:
    """生成清晰的匹配类型，例如 Final Yellow Match (3-to-1)。"""

    return (
        f"{MATCH_TYPE_PREFIX} "
        f"({left_size}-to-{right_size})"
    )


def mark_final_yellow_match(
    left_records,
    right_records,
) -> None:
    """
    标记最终黄色区匹配结果。

    这些记录：
    - 原本均为未匹配记录；
    - 本阶段不要求 Key word 相同；
    - 成功后由 Excel 输出层标成蓝色；
    - Match Type 保留实际几对几结构。
    """

    left_rows = [record.row for record in left_records]
    right_rows = [record.row for record in right_records]

    match_type = build_match_type(
        left_size=len(left_records),
        right_size=len(right_records),
    )

    for record in left_records:
        record.matched = True
        record.match_type = match_type
        record.partners = list(right_rows)
        record.keyword_match = False
        record.keyword_conflict = True
        record.review_required = False
        record.review_reason = ""

    for record in right_records:
        record.matched = True
        record.match_type = match_type
        record.partners = list(left_rows)
        record.keyword_match = False
        record.keyword_conflict = True
        record.review_required = False
        record.review_reason = ""


def unmatched_by_sign(records, sign: int):
    """取得指定金额方向的未匹配记录。"""

    return [
        record
        for record in records
        if (
            not record.matched
            and amount_sign(record.amount) == sign
        )
    ]


def find_exact_group(
    target_amount: Decimal,
    candidates,
    group_size: int,
):
    """
    在候选记录中寻找固定笔数的精确金额组合。

    使用 DFS + 剪枝，避免暴力枚举。
    """

    if group_size <= 0 or len(candidates) < group_size:
        return None

    target_value = abs(target_amount)

    if target_value == 0:
        return None

    sorted_candidates = sorted(
        candidates,
        key=lambda record: (
            abs(record.amount),
            record.row,
        ),
    )

    values = [
        abs(record.amount)
        for record in sorted_candidates
    ]

    candidate_count = len(sorted_candidates)
    selected_indexes = []

    def minimum_possible_sum(
        start_index: int,
        remaining_count: int,
    ):
        end_index = start_index + remaining_count

        if end_index > candidate_count:
            return None

        return sum(
            values[start_index:end_index],
            Decimal("0"),
        )

    def maximum_possible_sum(
        start_index: int,
        remaining_count: int,
    ):
        available_count = candidate_count - start_index

        if available_count < remaining_count:
            return None

        return sum(
            values[candidate_count - remaining_count:],
            Decimal("0"),
        )

    def search(
        start_index: int,
        remaining_count: int,
        current_sum: Decimal,
    ):
        if remaining_count == 0:
            if current_sum == target_value:
                return list(selected_indexes)
            return None

        if candidate_count - start_index < remaining_count:
            return None

        smallest_extra = minimum_possible_sum(
            start_index=start_index,
            remaining_count=remaining_count,
        )

        if smallest_extra is None:
            return None

        if current_sum + smallest_extra > target_value:
            return None

        largest_extra = maximum_possible_sum(
            start_index=start_index,
            remaining_count=remaining_count,
        )

        if largest_extra is None:
            return None

        if current_sum + largest_extra < target_value:
            return None

        final_start_index = candidate_count - remaining_count

        for index in range(
            start_index,
            final_start_index + 1,
        ):
            value = values[index]
            new_sum = current_sum + value

            if new_sum > target_value:
                break

            selected_indexes.append(index)

            result = search(
                start_index=index + 1,
                remaining_count=remaining_count - 1,
                current_sum=new_sum,
            )

            if result is not None:
                return result

            selected_indexes.pop()

        return None

    matched_indexes = search(
        start_index=0,
        remaining_count=group_size,
        current_sum=Decimal("0"),
    )

    if matched_indexes is None:
        return None

    return [
        sorted_candidates[index]
        for index in matched_indexes
    ]


def match_one_to_many_direction(
    single_records,
    many_records,
    single_is_left: bool,
    group_size: int,
) -> int:
    """
    执行最终黄色区的一个方向匹配。

    single_is_left=True:
        Sheet1 1笔 ↔ Sheet2 N笔

    single_is_left=False:
        Sheet1 N笔 ↔ Sheet2 1笔
    """

    matched_groups = 0

    for sign in (1, -1):
        singles = unmatched_by_sign(
            single_records,
            sign,
        )

        for single in singles:
            if single.matched:
                continue

            candidates = unmatched_by_sign(
                many_records,
                sign,
            )

            group = find_exact_group(
                target_amount=single.amount,
                candidates=candidates,
                group_size=group_size,
            )

            if group is None:
                continue

            if single_is_left:
                left_records = [single]
                right_records = group
            else:
                left_records = group
                right_records = [single]

            if sum_records(left_records) != sum_records(right_records):
                continue

            mark_final_yellow_match(
                left_records=left_records,
                right_records=right_records,
            )

            matched_groups += 1

    return matched_groups


def match_final_yellow_records(
    sheet1_records,
    sheet2_records,
):
    """
    最终黄色区匹配。

    规则：
    1. 只处理前面阶段结束后仍未匹配的记录；
    2. 不要求 Key word 相同；
    3. 不要求 Key word 为空；
    4. 金额必须完全相等；
    5. 只支持：
       - 1↔1
       - 1↔2 / 2↔1
       - 一直到 1↔6 / 6↔1
    6. 不执行 2↔2；
    7. 成功后统一标为 Final Yellow Match；
    8. Excel 中统一使用蓝色。
    """

    results = {}

    for group_size in range(
        1,
        MAX_ONE_TO_MANY_SIZE + 1,
    ):
        if group_size == 1:
            results["One-to-One"] = (
                match_one_to_many_direction(
                    single_records=sheet1_records,
                    many_records=sheet2_records,
                    single_is_left=True,
                    group_size=1,
                )
            )
            continue

        one_to_many_key = f"One-to-{group_size}"
        many_to_one_key = f"{group_size}-to-One"

        results[one_to_many_key] = (
            match_one_to_many_direction(
                single_records=sheet1_records,
                many_records=sheet2_records,
                single_is_left=True,
                group_size=group_size,
            )
        )

        results[many_to_one_key] = (
            match_one_to_many_direction(
                single_records=sheet2_records,
                many_records=sheet1_records,
                single_is_left=False,
                group_size=group_size,
            )
        )

    return results


def print_final_yellow_summary(results) -> None:
    """打印最终黄色区匹配汇总。"""

    print()
    print("=" * 56)
    print("Final Yellow Matching Summary")
    print("=" * 56)

    total_groups = 0

    for match_type, count in results.items():
        print(f"{match_type:<22}: {count}")
        total_groups += count

    print("-" * 56)
    print(f"Total groups           : {total_groups}")
    print("=" * 56)