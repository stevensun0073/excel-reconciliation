from __future__ import annotations

from decimal import Decimal
from itertools import combinations

from matcher import amount_sign, sum_records


MATCH_TYPE = "Keyword Difference"
MAX_ONE_TO_MANY_SIZE = 6


def mark_keyword_difference(
    left_records,
    right_records,
) -> None:
    """
    标记最终黄色区匹配结果。

    所有参与记录：
    - matched=True
    - match_type="Keyword Difference"
    - keyword_conflict=True
    - Keyword 单元格由 excel_io.py 标红
    """

    left_rows = [
        record.row
        for record in left_records
    ]

    right_rows = [
        record.row
        for record in right_records
    ]

    for record in left_records:
        record.matched = True
        record.match_type = MATCH_TYPE
        record.partners = list(right_rows)
        record.keyword_match = False
        record.keyword_conflict = True
        record.review_required = False
        record.review_reason = ""

    for record in right_records:
        record.matched = True
        record.match_type = MATCH_TYPE
        record.partners = list(left_rows)
        record.keyword_match = False
        record.keyword_conflict = True
        record.review_required = False
        record.review_reason = ""


def unmatched_by_sign(
    records,
    sign: int,
):
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

    金额必须完全相等。
    """

    if len(candidates) < group_size:
        return None

    target_value = abs(target_amount)

    sorted_candidates = sorted(
        candidates,
        key=lambda record: (
            abs(record.amount),
            record.row,
        )
    )

    for group in combinations(
        sorted_candidates,
        group_size,
    ):
        if sum(
            (abs(record.amount) for record in group),
            Decimal("0"),
        ) == target_value:
            return list(group)

    return None


def match_one_to_many_direction(
    single_records,
    many_records,
    single_is_left: bool,
    group_size: int,
) -> int:
    """
    执行一个方向的一对多匹配。

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

            if sum_records(
                left_records
            ) != sum_records(
                right_records
            ):
                continue

            mark_keyword_difference(
                left_records=left_records,
                right_records=right_records,
            )

            matched_groups += 1

    return matched_groups


def build_pair_index(
    records,
):
    """
    为未匹配记录建立两笔组合金额索引。

    返回：
        total_amount -> [pair, pair, ...]
    """

    index = {}

    for sign in (1, -1):
        sign_records = unmatched_by_sign(
            records,
            sign,
        )

        for pair in combinations(
            sign_records,
            2,
        ):
            total = sum_records(pair)

            index.setdefault(
                total,
                [],
            ).append(pair)

    return index


def match_two_to_two(
    sheet1_records,
    sheet2_records,
) -> int:
    """
    执行最终黄色区的 2↔2 精确金额匹配。

    规则：
    - 只处理未匹配记录；
    - 两边金额合计必须完全相等；
    - 两边内部金额方向必须一致；
    - 不要求 Key word 相同。
    """

    matched_groups = 0

    left_index = build_pair_index(
        sheet1_records
    )

    for sign in (1, -1):
        right_records = unmatched_by_sign(
            sheet2_records,
            sign,
        )

        for right_pair in combinations(
            right_records,
            2,
        ):
            if any(
                record.matched
                for record in right_pair
            ):
                continue

            total = sum_records(
                right_pair
            )

            candidate_left_pairs = (
                left_index.get(
                    total,
                    [],
                )
            )

            matched_left_pair = None

            for left_pair in candidate_left_pairs:
                if any(
                    record.matched
                    for record in left_pair
                ):
                    continue

                if sum_records(
                    left_pair
                ) != total:
                    continue

                matched_left_pair = list(
                    left_pair
                )
                break

            if matched_left_pair is None:
                continue

            mark_keyword_difference(
                left_records=matched_left_pair,
                right_records=list(right_pair),
            )

            matched_groups += 1

    return matched_groups


def match_final_yellow_records(
    sheet1_records,
    sheet2_records,
):
    """
    最终黄色区匹配。

    只处理前面所有阶段结束后仍未匹配的记录。

    规则：
    1. 不要求 Key word 相同；
    2. 不要求 Key word 为空；
    3. 金额必须完全相等；
    4. 支持：
       - 1↔1
       - 1↔2 / 2↔1
       - 1↔3 / 3↔1
       - 1↔4 / 4↔1
       - 1↔5 / 5↔1
       - 1↔6 / 6↔1
       - 2↔2
    5. 成功后统一记为 Keyword Difference；
    6. 所有参与记录的 Keyword 单元格标红；
    7. 已经匹配成功的记录完全不动。
    """

    results = {}

    # 先做笔数最少的一对多 / 多对一。
    for group_size in range(
        1,
        MAX_ONE_TO_MANY_SIZE + 1,
    ):
        left_key = f"One-to-{group_size}"
        right_key = f"{group_size}-to-One"

        results[left_key] = (
            match_one_to_many_direction(
                single_records=sheet1_records,
                many_records=sheet2_records,
                single_is_left=True,
                group_size=group_size,
            )
        )

        if group_size == 1:
            # 1↔1 已经在上面完成，反方向不重复执行。
            results[right_key] = 0
            continue

        results[right_key] = (
            match_one_to_many_direction(
                single_records=sheet2_records,
                many_records=sheet1_records,
                single_is_left=False,
                group_size=group_size,
            )
        )

    # 最后执行 2↔2，避免它先抢走可用于更简单匹配的记录。
    results["Two-to-Two"] = (
        match_two_to_two(
            sheet1_records=sheet1_records,
            sheet2_records=sheet2_records,
        )
    )

    return results


def print_final_yellow_summary(
    results,
) -> None:
    """打印最终黄色区匹配汇总。"""

    print()
    print("=" * 56)
    print("Final Yellow Matching Summary")
    print("=" * 56)

    total_groups = 0

    for match_type, count in results.items():
        print(
            f"{match_type:<22}: {count}"
        )
        total_groups += count

    print("-" * 56)
    print(
        f"Total groups           : "
        f"{total_groups}"
    )
    print("=" * 56)