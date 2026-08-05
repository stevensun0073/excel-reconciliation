from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal
import re

from matcher import amount_sign, sum_records


MAX_COMBINATION_SIZE = 6

BLUE_UNIQUE_PREFIX = "Blue Unique Amount"
BLUE_REFERENCE_PREFIX = "Blue Same Reference"
BLUE_REVERSAL_PREFIX = "Blue Same Reference Reversal"

YELLOW_KEYWORD_PREFIX = "Yellow Keyword Group"
YELLOW_AMOUNT_PREFIX = "Yellow Amount Only"

SPECIAL_MULTIPLE = Decimal("872")
SPECIAL_MULTIPLE_REASON = "872的整数倍"


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


def get_raw_extra_value(record, possible_names):
    """
    从 record.extra 中按不区分大小写的字段名读取原始值。
    """

    normalized_names = {
        normalize(name)
        for name in possible_names
    }

    for key, value in record.extra.items():
        if normalize(key) in normalized_names:
            return value

    return None


def extract_reference_date(reference: str) -> str:
    """
    从银行流水号中提取 YYYYMMDD 日期。

    例如：
    650090000017880|20260528|051312...
    返回：
    20260528
    """

    if not reference:
        return ""

    match = re.search(
        r"(?<!\d)(20\d{6})(?!\d)",
        str(reference),
    )

    if match is None:
        return ""

    date_text = match.group(1)

    try:
        datetime.strptime(
            date_text,
            "%Y%m%d",
        )
    except ValueError:
        return ""

    return date_text


def get_transaction_date(record) -> str:
    """
    从 Sheet2 Transaction Time 中取得 YYYYMMDD 日期。
    """

    value = get_raw_extra_value(
        record,
        {
            "Transaction Time",
            "transaction time",
            "交易时间",
            "交易日期",
            "日期",
        },
    )

    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.strftime("%Y%m%d")

    if isinstance(value, date):
        return value.strftime("%Y%m%d")

    text = str(value).strip()

    patterns = (
        r"(?<!\d)(20\d{2})[-/\.](\d{1,2})[-/\.](\d{1,2})(?!\d)",
        r"(?<!\d)(20\d{6})(?!\d)",
    )

    match = re.search(
        patterns[0],
        text,
    )

    if match is not None:
        year, month, day = match.groups()

        try:
            parsed = datetime(
                int(year),
                int(month),
                int(day),
            )
        except ValueError:
            return ""

        return parsed.strftime("%Y%m%d")

    match = re.search(
        patterns[1],
        text,
    )

    if match is None:
        return ""

    date_text = match.group(1)

    try:
        datetime.strptime(
            date_text,
            "%Y%m%d",
        )
    except ValueError:
        return ""

    return date_text


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


def mark_internal_sheet1_group(
    records,
    match_type: str,
    review_reason: str,
) -> None:
    """
    标记Sheet1内部同组关系。

    每条记录的Partner Rows列出同组内其余所有行号。
    """

    group_rows = [
        record.row
        for record in records
    ]

    for record in records:
        record.matched = True
        record.match_type = match_type
        record.partners = [
            row
            for row in group_rows
            if row != record.row
        ]
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
#
# 若 Sheet2 有多条相同金额候选：
# 用银行流水号中的日期与 Transaction Time 日期核对。
# 只有日期能够唯一锁定一条候选时才匹配。
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
    date_resolved_groups = 0

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

        if not candidates:
            continue

        sheet2_record = None
        review_reason = "同银行流水号多对一"

        if len(candidates) == 1:
            sheet2_record = candidates[0]

        else:
            reference_date = extract_reference_date(
                reference
            )

            if not reference_date:
                continue

            date_candidates = [
                record
                for record in candidates
                if (
                    get_transaction_date(record)
                    == reference_date
                )
            ]

            if len(date_candidates) != 1:
                continue

            sheet2_record = date_candidates[0]
            review_reason = (
                "同银行流水号多对一；"
                "重复金额按流水号日期唯一确定"
            )
            date_resolved_groups += 1

        mark_match(
            left_records=group_records,
            right_records=[sheet2_record],
            match_type=(
                f"{BLUE_REFERENCE_PREFIX} "
                f"({len(group_records)}-to-1)"
            ),
            review_reason=review_reason,
        )

        matched_groups += 1
        matched_sheet1_rows += len(group_records)

    return {
        "groups": matched_groups,
        "sheet1_rows": matched_sheet1_rows,
        "date_resolved_groups": date_resolved_groups,
    }


# ============================================================
# 淡蓝色规则三：
# Sheet1内部正负冲销
#
# 只处理Sheet1，不处理Sheet2。
#
# 分组规则：
# 1. 相同有效银行流水号归为同一组；
# 2. 流水号为空、None、False、"false"、"none"，
#    全部视为同一个流水号组；
# 3. 不同有效流水号绝不互相冲销。
#
# 执行位置：
# 黄色组合全部完成以后。
#
# 匹配规则：
# 1. 同组全部剩余记录合计严格等于0，整组冲销；
# 2. 若整组不为0，继续匹配金额绝对值完全相同的
#    一正一负记录。
# ============================================================

def normalize_reversal_reference(reference: str) -> str:
    normalized_reference = normalize(reference)

    if normalized_reference in {
        "",
        "false",
        "none",
    }:
        return "__NO_REFERENCE__"

    return normalized_reference


def match_same_reference_reversals(
    sheet1_records,
) -> dict:
    groups = defaultdict(list)

    for record in sheet1_records:
        if record.matched:
            continue

        reference = normalize_reversal_reference(
            get_bank_reference(record)
        )

        groups[reference].append(record)

    matched_zero_sum_groups = 0
    matched_pairs = 0
    matched_sheet1_rows = 0

    for reference in sorted(groups):
        group_records = sorted(
            (
                record
                for record in groups[reference]
                if not record.matched
            ),
            key=lambda record: record.row,
        )

        if len(group_records) < 2:
            continue

        group_total = sum_records(
            group_records
        )

        # 同组全部剩余记录合计严格为0：
        # 整组标记为Sheet1内部冲销。
        if group_total == Decimal("0"):
            if reference == "__NO_REFERENCE__":
                review_reason = (
                    "Sheet1空流水号记录整组合计为零"
                )
            else:
                review_reason = (
                    "Sheet1同银行流水号整组合计为零"
                )

            mark_internal_sheet1_group(
                records=group_records,
                match_type=(
                    f"{BLUE_REVERSAL_PREFIX} "
                    f"({len(group_records)}-row group)"
                ),
                review_reason=review_reason,
            )

            matched_zero_sum_groups += 1
            matched_sheet1_rows += len(
                group_records
            )
            continue

        # 整组不为0时，保留原有的一正一负精确冲销。
        positives = defaultdict(list)
        negatives = defaultdict(list)

        for record in group_records:
            if record.matched:
                continue

            if record.amount > 0:
                positives[
                    record.amount
                ].append(record)

            elif record.amount < 0:
                negatives[
                    abs(record.amount)
                ].append(record)

        for amount in sorted(positives):
            positive_records = positives[
                amount
            ]

            negative_records = negatives.get(
                amount,
                [],
            )

            pair_count = min(
                len(positive_records),
                len(negative_records),
            )

            for index in range(
                pair_count
            ):
                positive_record = (
                    positive_records[index]
                )

                negative_record = (
                    negative_records[index]
                )

                if (
                    positive_record.matched
                    or negative_record.matched
                ):
                    continue

                if (
                    positive_record.amount
                    + negative_record.amount
                    != Decimal("0")
                ):
                    continue

                if reference == "__NO_REFERENCE__":
                    review_reason = (
                        "Sheet1空流水号正负冲销"
                    )
                else:
                    review_reason = (
                        "Sheet1同银行流水号正负冲销"
                    )

                mark_internal_sheet1_group(
                    records=[
                        positive_record,
                        negative_record,
                    ],
                    match_type=(
                        f"{BLUE_REVERSAL_PREFIX} "
                        "(1-to-1)"
                    ),
                    review_reason=review_reason,
                )

                matched_pairs += 1
                matched_sheet1_rows += 2

    return {
        "zero_sum_groups": (
            matched_zero_sum_groups
        ),
        "pairs": matched_pairs,
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
# 最终特殊标记：
# 对最后仍未匹配、且金额绝对值为872整数倍的记录，
# 保持未匹配状态，仅写入复核说明。
# ============================================================

def mark_remaining_872_multiples(records) -> int:
    marked_rows = 0

    for record in records:
        if record.matched:
            continue

        amount = abs(record.amount)

        # 0不作为872的整数倍进行特殊标记。
        if amount == Decimal("0"):
            continue

        if amount % SPECIAL_MULTIPLE != Decimal("0"):
            continue

        record.review_required = True
        record.review_reason = SPECIAL_MULTIPLE_REASON
        marked_rows += 1

    return marked_rows


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
        "blue_reference_date_resolved": 0,
        "blue_reversal_zero_sum_groups": 0,
        "blue_reversal_pairs": 0,
        "blue_reversal_sheet1_rows": 0,
        "blue_unique_final_one_to_one": 0,
        "keyword_one_to_many": {},
        "keyword_many_to_one": {},
        "amount_one_to_many": {},
        "amount_many_to_one": {},
        "multiple_872_sheet1": 0,
        "multiple_872_sheet2": 0,
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

    results["blue_reference_date_resolved"] = (
        reference_result["date_resolved_groups"]
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

    # 淡蓝色三：
    # 黄色组合全部完成后，
    # 仅对Sheet1剩余记录执行内部正负冲销。
    reversal_result = (
        match_same_reference_reversals(
            sheet1_records,
        )
    )

    results[
        "blue_reversal_zero_sum_groups"
    ] = reversal_result[
        "zero_sum_groups"
    ]

    results["blue_reversal_pairs"] = (
        reversal_result["pairs"]
    )

    results["blue_reversal_sheet1_rows"] = (
        reversal_result["sheet1_rows"]
    )

    # 淡蓝色收尾：
    # 黄色组合完成后，部分重复金额可能已经变成
    # 双方剩余唯一金额，因此再执行一次唯一金额一对一。
    results["blue_unique_final_one_to_one"] = (
        match_unique_amount_one_to_one(
            sheet1_records,
            sheet2_records,
        )
    )

    # 最终特殊规则：
    # 对仍未匹配且金额为872整数倍的记录进行特殊着色和说明。
    # 这些记录仍然属于未匹配记录。
    results["multiple_872_sheet1"] = (
        mark_remaining_872_multiples(
            sheet1_records,
        )
    )

    results["multiple_872_sheet2"] = (
        mark_remaining_872_multiples(
            sheet2_records,
        )
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

    print(
        "  Date-resolved groups       : "
        f"{results['blue_reference_date_resolved']}"
    )

    print(
        "  Reversal zero-sum groups   : "
        f"{results['blue_reversal_zero_sum_groups']}"
    )

    print(
        "  Reversal pairs             : "
        f"{results['blue_reversal_pairs']}"
    )

    print(
        "  Reversal Sheet1 rows       : "
        f"{results['blue_reversal_sheet1_rows']}"
    )

    print(
        "  Final unique 1-to-1        : "
        f"{results['blue_unique_final_one_to_one']}"
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
    print("Final Special Marking")

    print(
        "  Sheet1 multiples of 872    : "
        f"{results['multiple_872_sheet1']}"
    )

    print(
        "  Sheet2 multiples of 872    : "
        f"{results['multiple_872_sheet2']}"
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