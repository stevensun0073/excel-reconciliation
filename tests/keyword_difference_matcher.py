from decimal import Decimal
from itertools import combinations

from matcher import (
    amount_sign,
    get_sheet2_keyword,
    parse_sheet1_keywords,
    sum_records,
)


MATCH_TYPE_PREFIX = "Keyword-Difference"
MAX_BLANK_GROUP_SIZE = 3


def keyword_is_blank(record) -> bool:
    """判断一条记录的 Key word 是否为空。"""

    if record.source_sheet == "Sheet1":
        return not parse_sheet1_keywords(record)

    if record.source_sheet == "Sheet2":
        return not get_sheet2_keyword(record)

    return True


def find_blank_keyword_combination(
    target_amount: Decimal,
    records,
    max_group_size: int = MAX_BLANK_GROUP_SIZE,
):
    """
    只在同一张表中，从尚未匹配且 Key word 为空的记录里，
    寻找 1～3 笔合计等于 target_amount 的组合。
    """

    target_sign = amount_sign(target_amount)

    if target_sign == 0:
        return None

    candidates = [
        record
        for record in records
        if (
            not record.matched
            and keyword_is_blank(record)
            and amount_sign(record.amount) == target_sign
        )
    ]

    candidates.sort(
        key=lambda record: (
            abs(record.amount),
            record.row,
        )
    )

    largest_group_size = min(
        max_group_size,
        len(candidates),
    )

    for group_size in range(
        1,
        largest_group_size + 1,
    ):
        for group in combinations(
            candidates,
            group_size,
        ):
            if sum_records(group) == target_amount:
                return list(group)

    return None


def mark_keyword_difference_match(
    left_records,
    right_records,
) -> None:
    """
    标记新的 Key word 差额补齐匹配。

    所有参与记录均设置 keyword_conflict=True，
    由 Excel 输出层把所有 Key word 单元格标红。
    """

    left_rows = [
        record.row
        for record in left_records
    ]

    right_rows = [
        record.row
        for record in right_records
    ]

    match_type = (
        f"{MATCH_TYPE_PREFIX} "
        f"({len(left_records)}-to-{len(right_records)})"
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


def match_keyword_differences(
    sheet1_records,
    sheet2_records,
) -> int:
    """
    只处理原匹配流程结束后仍未匹配的记录。

    规则：
    1. 以 Sheet2 的 Key word 为基准；
    2. 只读取尚未匹配的记录；
    3. 对每个 Key word、每个金额方向分别处理；
    4. 汇总两边相同 Key word 的剩余金额；
    5. 差额只允许在金额较小的一侧补齐；
    6. 补差记录必须来自同一张表且 Key word 为空；
    7. 最多使用 3 条补差记录；
    8. 不允许 Sheet1、Sheet2 两边同时补差；
    9. 已匹配记录完全不变。
    """

    matched_groups = 0

    sheet2_keywords = sorted({
        get_sheet2_keyword(record)
        for record in sheet2_records
        if (
            not record.matched
            and get_sheet2_keyword(record)
        )
    })

    for keyword in sheet2_keywords:
        for sign in (1, -1):
            right_keyword_records = [
                record
                for record in sheet2_records
                if (
                    not record.matched
                    and get_sheet2_keyword(record) == keyword
                    and amount_sign(record.amount) == sign
                )
            ]

            if not right_keyword_records:
                continue

            left_keyword_records = [
                record
                for record in sheet1_records
                if (
                    not record.matched
                    and keyword in parse_sheet1_keywords(record)
                    and amount_sign(record.amount) == sign
                )
            ]

            if not left_keyword_records:
                left_total = Decimal("0")
            else:
                left_total = sum_records(
                    left_keyword_records
                )

            right_total = sum_records(
                right_keyword_records
            )

            if left_total == right_total:
                continue

            left_magnitude = abs(left_total)
            right_magnitude = abs(right_total)

            final_left_records = list(
                left_keyword_records
            )
            final_right_records = list(
                right_keyword_records
            )

            if left_magnitude < right_magnitude:
                difference = Decimal(sign) * (
                    right_magnitude - left_magnitude
                )

                supplement = (
                    find_blank_keyword_combination(
                        target_amount=difference,
                        records=sheet1_records,
                    )
                )

                if supplement is None:
                    continue

                final_left_records.extend(
                    supplement
                )

            else:
                difference = Decimal(sign) * (
                    left_magnitude - right_magnitude
                )

                supplement = (
                    find_blank_keyword_combination(
                        target_amount=difference,
                        records=sheet2_records,
                    )
                )

                if supplement is None:
                    continue

                final_right_records.extend(
                    supplement
                )

            if (
                not final_left_records
                or not final_right_records
            ):
                continue

            if sum_records(
                final_left_records
            ) != sum_records(
                final_right_records
            ):
                continue

            mark_keyword_difference_match(
                left_records=final_left_records,
                right_records=final_right_records,
            )

            matched_groups += 1

    return matched_groups