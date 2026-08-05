from __future__ import annotations

from collections import Counter, defaultdict
from decimal import Decimal

from matcher import amount_sign, sum_records


MAX_COMBINATION_SIZE = 6

BLUE_UNIQUE_PREFIX = "Blue Unique Amount"
BLUE_REFERENCE_PREFIX = "Blue Same Reference"

YELLOW_KEYWORD_PREFIX = "Yellow Keyword Group"
YELLOW_AMOUNT_PREFIX = "Yellow Amount Only"


def normalize(value) -> str:
    if value is None:
        return ""

    return str(value).strip().lower()


def get_extra_value(record, possible_names) -> str:
    """
    从 record.extra 中按不区分大小写的字段名读取值。
    """

    normalized_names = {
        normalize(name)
        for name in possible_names
    }

    for key, value in record.extra.items():
        if normalize(key) in normalized_names:
            return normalize(value)

    return ""


def get_key_word(record) -> str:
    return get_extra_value(
        record,
        {
            "Key word",
            "Keyword",
            "key word",
        },
    )


def get_bank_reference(record) -> str:
    return get_extra_value(
        record,
        {
            "银行流水号",
            "银行交易流水号",
            "银行参考号",
            "流水号",
            "Bank Reference",
            "Reference Number",
        },
    )


def get_business_type(record) -> str:
    return get_extra_value(
        record,
        {
            "Business Type",
            "business type",
            "业务类型",
        },
    )


def unmatched_records(records):
    return [
        record
        for record in records
        if not record.matched
    ]


def unmatched_by_sign(records, sign: int):
    return [
        record
        for record in records
        if (
            not record.matched
            and amount_sign(record.amount) == sign
        )
    ]


def mark_match(
    left_records,
    right_records,
    match_type: str,
    review_reason: str,
) -> None:
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
        record.match_type = match_type
        record.partners = list(right_rows)
        record.review_required = False
        record.review_reason = review_reason
        record.keyword_match = False
        record.keyword_conflict = True

    for record in right_records:
        record.matched = True
        record.match_type = match_type
        record.partners = list(left_rows)
        record.review_required = False
        record.review_reason = review_reason
        record.keyword_match = False
        record.keyword_conflict = True


# ============================================================
# 淡蓝色规则一：
# 双方唯一金额一对一
# ============================================================

def match_unique_amount_one_to_one(
    sheet1_records,
    sheet2_records,
) -> int:
    sheet1_unmatched = unmatched_records(
        sheet1_records
    )

    sheet2_unmatched = unmatched_records(
        sheet2_records
    )

    sheet1_counts = Counter(
        record.amount
        for record in sheet1_unmatched
    )

    sheet2_counts = Counter(
        record.amount
        for record in sheet2_unmatched
    )

    sheet2_unique_map = {
        record.amount: record
        for record in sheet2_unmatched
        if sheet2_counts[record.amount] == 1
    }

    matched_pairs = 0

    for sheet1_record in sheet1_unmatched:
        if sheet1_record.matched:
            continue

        amount = sheet1_record.amount

        if sheet1_counts[amount] != 1:
            continue

        if sheet2_counts.get(amount, 0) != 1:
            continue

        sheet2_record = sheet2_unique_map[amount]

        if sheet2_record.matched:
            continue

        mark_match(
            left_records=[sheet1_record],
            right_records=[sheet2_record],
            match_type=(
                f"{BLUE_UNIQUE_PREFIX} (1-to-1)"
            ),
            review_reason="双方唯一金额一对一",
        )

        matched_pairs += 1

    return matched_pairs


# ============================================================
# 淡蓝色规则二：
# Sheet1 同银行流水号全部记录 → Sheet2 单条
# ============================================================

def match_same_reference_group_to_one(
    sheet1_records,
    sheet2_records,
) -> dict:
    groups = defaultdict(list)

    for record in sheet1_records:
        if record.matched:
            continue

        reference = get_bank_reference(record)

        if not reference:
            continue

        groups[reference].append(record)

    matched_groups = 0
    matched_sheet1_rows = 0

    for reference, group_records in groups.items():
        if len(group_records) < 2:
            continue

        if any(record.matched for record in group_records):
            continue

        total = sum_records(group_records)

        candidates = [
            record
            for record in sheet2_records
            if (
                not record.matched
                and record.amount == total
            )
        ]

        # 必须只有一条唯一候选。
        if len(candidates) != 1:
            continue

        sheet2_record = candidates[0]

        mark_match(
            left_records=group_records,
            right_records=[sheet2_record],
            match_type=(
                f"{BLUE_REFERENCE_PREFIX} "
                f"({len(group_records)}-to-1)"
            ),
            review_reason="同银行流水号多对一",
        )

        matched_groups += 1
        matched_sheet1_rows += len(group_records)

    return {
        "groups": matched_groups,
        "sheet1_rows": matched_sheet1_rows,
    }


# ============================================================
# BANK → Charging 业务规则
# ============================================================

def sheet2_allowed_for_sheet1(
    sheet1_record,
    sheet2_record,
) -> bool:
    """
    Sheet1 Key word 为 BANK 时，
    只能匹配 Sheet2 Business Type 为 Charging。
    """

    if get_key_word(sheet1_record) != "bank":
        return True

    return (
        get_business_type(sheet2_record)
        == "charging"
    )


def sheet1_allowed_for_sheet2(
    sheet1_record,
    sheet2_record,
) -> bool:
    """
    Sheet2 Business Type 不是 Charging 时，
    Sheet1 BANK 不允许参加组合。
    """

    if get_key_word(sheet1_record) != "bank":
        return True

    return (
        get_business_type(sheet2_record)
        == "charging"
    )


# ============================================================
# 快速寻找第一个固定笔数组合
# ============================================================

def find_exact_group(
    target_amount: Decimal,
    candidates,
    group_size: int,
):
    """
    寻找第一个金额完全相等的固定笔数组合。

    找到第一个组合立即返回，
    不继续证明是否存在第二个解。
    """

    if group_size <= 0:
        return None

    if len(candidates) < group_size:
        return None

    target_value = abs(target_amount)

    if target_value == 0:
        return None

    target_sign = amount_sign(target_amount)

    filtered_candidates = [
        record
        for record in candidates
        if (
            not record.matched
            and amount_sign(record.amount)
            == target_sign
            and abs(record.amount)
            <= target_value
        )
    ]

    if len(filtered_candidates) < group_size:
        return None

    sorted_candidates = sorted(
        filtered_candidates,
        key=lambda record: (
            abs(record.amount),
            record.row,
        ),
    )

    # states[count][amount] = 第一个组合
    states = [
        {}
        for _ in range(group_size + 1)
    ]

    states[0][Decimal("0")] = tuple()

    for record in sorted_candidates:
        value = abs(record.amount)

        for count in range(
            group_size - 1,
            -1,
            -1,
        ):
            current_states = list(
                states[count].items()
            )

            for current_sum, combination in current_states:
                new_sum = current_sum + value

                if new_sum > target_value:
                    continue

                new_count = count + 1

                if new_sum not in states[new_count]:
                    states[new_count][new_sum] = (
                        combination + (record,)
                    )

                if (
                    new_count == group_size
                    and new_sum == target_value
                ):
                    return list(
                        states[new_count][new_sum]
                    )

    return None


# ============================================================
# 同一张表内 Key word 相同的组合优先
# ============================================================

def group_unmatched_by_keyword(records):
    groups = defaultdict(list)

    for record in records:
        if record.matched:
            continue

        key_word = get_key_word(record)

        if not key_word:
            continue

        groups[key_word].append(record)

    return groups


def match_keyword_many_to_one(
    sheet1_records,
    sheet2_records,
    group_size: int,
) -> int:
    """
    Sheet1 内相同 Key word 的多条记录，
    匹配 Sheet2 一条记录。

    两张表之间 Key word 不要求相同。
    """

    matched_groups = 0

    for sign in (1, -1):
        sheet2_targets = unmatched_by_sign(
            sheet2_records,
            sign,
        )

        for sheet2_record in sheet2_targets:
            if sheet2_record.matched:
                continue

            keyword_groups = (
                group_unmatched_by_keyword(
                    sheet1_records
                )
            )

            found_group = None

            for key_word in sorted(keyword_groups):
                candidates = [
                    record
                    for record
                    in keyword_groups[key_word]
                    if sheet1_allowed_for_sheet2(
                        record,
                        sheet2_record,
                    )
                ]

                group = find_exact_group(
                    target_amount=sheet2_record.amount,
                    candidates=candidates,
                    group_size=group_size,
                )

                if group is not None:
                    found_group = group
                    break

            if found_group is None:
                continue

            if (
                sum_records(found_group)
                != sheet2_record.amount
            ):
                continue

            mark_match(
                left_records=found_group,
                right_records=[sheet2_record],
                match_type=(
                    f"{YELLOW_KEYWORD_PREFIX} "
                    f"({group_size}-to-1)"
                ),
                review_reason=(
                    f"Sheet1同Key word "
                    f"{group_size}对1"
                ),
            )

            matched_groups += 1

    return matched_groups


def match_keyword_one_to_many(
    sheet1_records,
    sheet2_records,
    group_size: int,
) -> int:
    """
    Sheet2 内相同 Key word 的多条记录，
    匹配 Sheet1 一条记录。

    两张表之间 Key word 不要求相同。
    """

    matched_groups = 0

    for sign in (1, -1):
        sheet1_targets = unmatched_by_sign(
            sheet1_records,
            sign,
        )

        for sheet1_record in sheet1_targets:
            if sheet1_record.matched:
                continue

            keyword_groups = (
                group_unmatched_by_keyword(
                    sheet2_records
                )
            )

            found_group = None

            for key_word in sorted(keyword_groups):
                candidates = [
                    record
                    for record
                    in keyword_groups[key_word]
                    if sheet2_allowed_for_sheet1(
                        sheet1_record,
                        record,
                    )
                ]

                group = find_exact_group(
                    target_amount=sheet1_record.amount,
                    candidates=candidates,
                    group_size=group_size,
                )

                if group is not None:
                    found_group = group
                    break

            if found_group is None:
                continue

            if (
                sheet1_record.amount
                != sum_records(found_group)
            ):
                continue

            mark_match(
                left_records=[sheet1_record],
                right_records=found_group,
                match_type=(
                    f"{YELLOW_KEYWORD_PREFIX} "
                    f"(1-to-{group_size})"
                ),
                review_reason=(
                    f"Sheet2同Key word "
                    f"1对{group_size}"
                ),
            )

            matched_groups += 1

    return matched_groups


# ============================================================
# 纯金额组合
# ============================================================

def match_amount_one_to_many(
    sheet1_records,
    sheet2_records,
    group_size: int,
) -> int:
    matched_groups = 0

    for sign in (1, -1):
        sheet1_targets = unmatched_by_sign(
            sheet1_records,
            sign,
        )

        for sheet1_record in sheet1_targets:
            if sheet1_record.matched:
                continue

            candidates = [
                record
                for record in unmatched_by_sign(
                    sheet2_records,
                    sign,
                )
                if sheet2_allowed_for_sheet1(
                    sheet1_record,
                    record,
                )
            ]

            group = find_exact_group(
                target_amount=sheet1_record.amount,
                candidates=candidates,
                group_size=group_size,
            )

            if group is None:
                continue

            if (
                sheet1_record.amount
                != sum_records(group)
            ):
                continue

            mark_match(
                left_records=[sheet1_record],
                right_records=group,
                match_type=(
                    f"{YELLOW_AMOUNT_PREFIX} "
                    f"(1-to-{group_size})"
                ),
                review_reason=(
                    f"纯金额1对{group_size}"
                ),
            )

            matched_groups += 1

    return matched_groups


def match_amount_many_to_one(
    sheet1_records,
    sheet2_records,
    group_size: int,
) -> int:
    matched_groups = 0

    for sign in (1, -1):
        sheet2_targets = unmatched_by_sign(
            sheet2_records,
            sign,
        )

        for sheet2_record in sheet2_targets:
            if sheet2_record.matched:
                continue

            candidates = [
                record
                for record in unmatched_by_sign(
                    sheet1_records,
                    sign,
                )
                if sheet1_allowed_for_sheet2(
                    record,
                    sheet2_record,
                )
            ]

            group = find_exact_group(
                target_amount=sheet2_record.amount,
                candidates=candidates,
                group_size=group_size,
            )

            if group is None:
                continue

            if (
                sum_records(group)
                != sheet2_record.amount
            ):
                continue

            mark_match(
                left_records=group,
                right_records=[sheet2_record],
                match_type=(
                    f"{YELLOW_AMOUNT_PREFIX} "
                    f"({group_size}-to-1)"
                ),
                review_reason=(
                    f"纯金额{group_size}对1"
                ),
            )

            matched_groups += 1

    return matched_groups


# ============================================================
# 正式剩余记录流程
# ============================================================

def match_remaining_records(
    sheet1_records,
    sheet2_records,
):
    results = {
        "blue_unique_one_to_one": 0,
        "blue_reference_groups": 0,
        "blue_reference_sheet1_rows": 0,
        "keyword_one_to_many": {},
        "keyword_many_to_one": {},
        "amount_one_to_many": {},
        "amount_many_to_one": {},
        "remaining_sheet1": 0,
        "remaining_sheet2": 0,
    }

    # 淡蓝色一：唯一金额一对一。
    results["blue_unique_one_to_one"] = (
        match_unique_amount_one_to_one(
            sheet1_records,
            sheet2_records,
        )
    )

    # 淡蓝色二：同银行流水号多对一。
    reference_result = (
        match_same_reference_group_to_one(
            sheet1_records,
            sheet2_records,
        )
    )

    results["blue_reference_groups"] = (
        reference_result["groups"]
    )

    results["blue_reference_sheet1_rows"] = (
        reference_result["sheet1_rows"]
    )

    # 淡黄色第一轮：
    # 同一张表内 Key word 相同优先。
    for group_size in range(
        2,
        MAX_COMBINATION_SIZE + 1,
    ):
        results["keyword_one_to_many"][
            group_size
        ] = match_keyword_one_to_many(
            sheet1_records,
            sheet2_records,
            group_size,
        )

        results["keyword_many_to_one"][
            group_size
        ] = match_keyword_many_to_one(
            sheet1_records,
            sheet2_records,
            group_size,
        )

    # 淡黄色第二轮：
    # 剩余记录纯金额组合。
    for group_size in range(
        2,
        MAX_COMBINATION_SIZE + 1,
    ):
        results["amount_one_to_many"][
            group_size
        ] = match_amount_one_to_many(
            sheet1_records,
            sheet2_records,
            group_size,
        )

        results["amount_many_to_one"][
            group_size
        ] = match_amount_many_to_one(
            sheet1_records,
            sheet2_records,
            group_size,
        )

    results["remaining_sheet1"] = len(
        unmatched_records(sheet1_records)
    )

    results["remaining_sheet2"] = len(
        unmatched_records(sheet2_records)
    )

    return results


def print_remaining_match_summary(
    results,
) -> None:
    print()
    print("=" * 68)
    print("Remaining Records Matching Summary")
    print("=" * 68)

    print()
    print("Blue Rules")

    print(
        "  Unique amount 1-to-1       : "
        f"{results['blue_unique_one_to_one']}"
    )

    print(
        "  Same reference groups      : "
        f"{results['blue_reference_groups']}"
    )

    print(
        "  Same reference Sheet1 rows : "
        f"{results['blue_reference_sheet1_rows']}"
    )

    print()
    print("Yellow Round 1 - Same Key Word Groups")

    for size in range(
        2,
        MAX_COMBINATION_SIZE + 1,
    ):
        print(
            f"  Keyword 1-to-{size:<2} : "
            f"{results['keyword_one_to_many'][size]}"
        )

        print(
            f"  Keyword {size}-to-1  : "
            f"{results['keyword_many_to_one'][size]}"
        )

    print()
    print("Yellow Round 2 - Amount Only")

    for size in range(
        2,
        MAX_COMBINATION_SIZE + 1,
    ):
        print(
            f"  Amount 1-to-{size:<2}  : "
            f"{results['amount_one_to_many'][size]}"
        )

        print(
            f"  Amount {size}-to-1   : "
            f"{results['amount_many_to_one'][size]}"
        )

    print()
    print("Business Rule")

    print(
        "  Sheet1 BANK can only match "
        "Sheet2 Charging"
    )

    print()
    print("Remaining")

    print(
        f"  Sheet1 : "
        f"{results['remaining_sheet1']}"
    )

    print(
        f"  Sheet2 : "
        f"{results['remaining_sheet2']}"
    )

    print("=" * 68)