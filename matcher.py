import re
from decimal import Decimal
from itertools import combinations
from typing import Callable, Optional


MAX_GROUP_SIZE = 10

NUMBER_NAMES = {
    1: "One",
    2: "Two",
    3: "Three",
    4: "Four",
    5: "Five",
    6: "Six",
    7: "Seven",
    8: "Eight",
    9: "Nine",
    10: "Ten",
}

MANY_TO_MANY_PATTERNS = [
    (2, 2),
    (2, 3),
    (3, 2),
    (2, 4),
    (4, 2),
    (2, 5),
    (5, 2),
    (2, 6),
    (6, 2),
]


def amount_equal(left: Decimal, right: Decimal) -> bool:
    """
    判断两个金额是否完全相等。

    银行对账不允许金额误差。
    """
    return left == right


def amount_sign(amount: Decimal) -> int:
    if amount > 0:
        return 1

    if amount < 0:
        return -1

    return 0


def sum_records(records) -> Decimal:
    return sum(
        (record.amount for record in records),
        Decimal("0"),
    )


def build_match_type(
    left_size: int,
    right_size: int,
) -> str:
    return (
        f"{NUMBER_NAMES[left_size]}"
        f"-to-"
        f"{NUMBER_NAMES[right_size]}"
    )


def normalize_keyword(value) -> str:
    """把一个 Key word 标准化，比较时不区分大小写。"""

    if value is None:
        return ""

    return str(value).strip().casefold()


def parse_sheet1_keywords(record) -> set[str]:
    """
    Sheet1 的 Key word 可能有多个，用分号分隔。

    例如：
        ETONG; ALFEM; ENGINEERING
    """

    extra = (
        record.extra
        if isinstance(record.extra, dict)
        else {}
    )

    raw_value = extra.get("Key word")

    if raw_value is None:
        return set()

    return {
        normalize_keyword(part)
        for part in str(raw_value).split(";")
        if normalize_keyword(part)
    }


def get_sheet2_keyword(record) -> str:
    """
    Sheet2 的 Key word 应只有一个。
    """

    extra = (
        record.extra
        if isinstance(record.extra, dict)
        else {}
    )

    return normalize_keyword(
        extra.get("Key word")
    )


def one_to_one_keywords_match(
    sheet1_record,
    sheet2_record,
) -> bool:
    """
    一对一关键词规则：

    只要 Sheet2 的单个 Key word 出现在
    Sheet1 的 Key word 集合中，就算完全匹配。

    任一方为空时返回 False。
    """

    sheet1_keywords = parse_sheet1_keywords(
        sheet1_record
    )

    sheet2_keyword = get_sheet2_keyword(
        sheet2_record
    )

    if not sheet1_keywords:
        return False

    if not sheet2_keyword:
        return False

    return sheet2_keyword in sheet1_keywords


def extract_first_three_words(value) -> list[str]:
    if value is None:
        return []

    words = re.findall(
        r"[A-Z0-9]+",
        str(value).upper(),
    )

    return words[:3]


def build_sheet2_keyword_set(record) -> set[str]:
    extra = (
        record.extra
        if isinstance(record.extra, dict)
        else {}
    )

    recipient_words = extract_first_three_words(
        extra.get("Recipient's Account Name")
    )

    description_words = extract_first_three_words(
        extra.get("Description")
    )

    return set(
        recipient_words
        + description_words
    )


def sheet2_group_has_common_words(
    records,
    minimum_common_words: int = 2,
) -> bool:
    if len(records) < 2:
        return False

    keyword_sets = [
        build_sheet2_keyword_set(record)
        for record in records
    ]

    if any(
        not keyword_set
        for keyword_set in keyword_sets
    ):
        return False

    common_words = set.intersection(
        *keyword_sets
    )

    return (
        len(common_words)
        >= minimum_common_words
    )


class Matcher:
    """
    银行日记账与银行对账单匹配引擎。

    本版本只升级一对一匹配：

    1. 金额必须完全相同；
    2. 优先在相同金额候选中寻找 Key word 匹配的记录；
    3. 如果存在 Key word 匹配：
       两边记录标记 keyword_match=True；
    4. 如果相同金额候选均不匹配 Key word：
       仍按金额确认一对一，
       两边记录标记 keyword_conflict=True；
    5. 任一边 Key word 为空，也属于关键词冲突。

    一对多和多对一先按 Key word 业务索引分池，再搜索金额组合。
多对多暂时关闭，待本阶段验证后再启用。
    """

    def __init__(
        self,
        sheet1_records,
        sheet2_records,
    ):
        self.sheet1 = sheet1_records
        self.sheet2 = sheet2_records

        # Key word 业务索引。
        #
        # Sheet1 一条记录可能有多个关键词，因此可以进入多个业务池。
        # Sheet2 一条记录只有一个关键词，因此只进入一个业务池。
        #
        # 索引只保存记录引用；是否已匹配、金额方向等条件，
        # 在实际搜索时动态检查。
        self.sheet1_keyword_index = (
            self.build_sheet1_keyword_index()
        )
        self.sheet2_keyword_index = (
            self.build_sheet2_keyword_index()
        )
    def build_sheet1_keyword_index(self):
        """
        建立 Sheet1 的 Key word 索引。

        示例：
            ETONG; ALFEM

        会同时进入：
            ETONG -> [record]
            ALFEM -> [record]
        """
        index = {}

        for record in self.sheet1:
            for keyword in parse_sheet1_keywords(
                record
            ):
                index.setdefault(
                    keyword,
                    [],
                ).append(record)

        return index

    def build_sheet2_keyword_index(self):
        """
        建立 Sheet2 的 Key word 索引。

        Sheet2 每条记录只应有一个 Key word。
        空白关键词不进入索引。
        """
        index = {}

        for record in self.sheet2:
            keyword = get_sheet2_keyword(
                record
            )

            if not keyword:
                continue

            index.setdefault(
                keyword,
                [],
            ).append(record)

        return index

    @staticmethod
    def filter_candidates(
        target,
        records,
    ):
        """
        从业务池中筛选：

        1. 尚未匹配；
        2. 与目标金额方向一致；
        3. 零金额不参加组合。

        最后按金额绝对值从小到大排序，
        供递归组合搜索使用。
        """
        target_sign = amount_sign(
            target.amount
        )

        if target_sign == 0:
            return []

        candidates = [
            record
            for record in records
            if (
                not record.matched
                and amount_sign(record.amount)
                == target_sign
            )
        ]

        candidates.sort(
            key=lambda record: abs(record.amount)
        )

        return candidates

    @staticmethod
    def mark_match(
        left_records,
        right_records,
        match_type: str,
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

            if hasattr(record, "review_required"):
                record.review_required = False

            if hasattr(record, "review_reason"):
                record.review_reason = ""

        for record in right_records:
            record.matched = True
            record.match_type = match_type
            record.partners = list(left_rows)

            if hasattr(record, "review_required"):
                record.review_required = False

            if hasattr(record, "review_reason"):
                record.review_reason = ""

    @staticmethod
    def set_one_to_one_keyword_result(
        left,
        right,
        keyword_matched: bool,
    ) -> None:
        """
        把一对一关键词核验结果写入两边记录。
        """

        for record in (left, right):
            record.keyword_match = keyword_matched
            record.keyword_conflict = (
                not keyword_matched
            )

            if keyword_matched:
                record.match_type = "One-to-One"
                record.review_required = False
                record.review_reason = ""
            else:
                record.match_type = "One-to-One (KW)"
                record.review_required = False
                record.review_reason = ""

    def run(self):
        """
        按顺序执行匹配。

        1. 一对一：
           - 先匹配相同金额、相同 Key word；
           - 同金额组剩余数量相等时，剩余记录统一记为
             Keyword Difference；
           - 剩余数量不等时，保留给后续阶段。
        2. 一对多 / 多对一，最多 1↔10；
        3. 多对多只启用：
           2↔2
           2↔3 / 3↔2
           2↔4 / 4↔2
           2↔5 / 5↔2
           2↔6 / 6↔2
        4. 不执行 3↔3 及以上两边都至少为 3 的组合。
        """
        results = {}

        (
            one_to_one_count,
            keyword_difference_count,
        ) = self.one_to_one_match()

        results["One-to-One"] = one_to_one_count
        results["Keyword Difference"] = (
            keyword_difference_count
        )

        for group_size in range(
            2,
            MAX_GROUP_SIZE + 1,
        ):
            one_to_many_type = build_match_type(
                left_size=1,
                right_size=group_size,
            )

            many_to_one_type = build_match_type(
                left_size=group_size,
                right_size=1,
            )

            results[one_to_many_type] = (
                self.one_to_many_match(
                    group_size=group_size,
                )
            )

            results[many_to_one_type] = (
                self.many_to_one_match(
                    group_size=group_size,
                )
            )

        for left_size, right_size in (
            MANY_TO_MANY_PATTERNS
        ):
            match_type = build_match_type(
                left_size=left_size,
                right_size=right_size,
            )

            results[match_type] = (
                self.many_to_many_match(
                    left_size=left_size,
                    right_size=right_size,
                )
            )

        self.print_summary(results)

        return results

    def one_to_one_match(self) -> tuple[int, int]:
        """
        执行同金额的一对一匹配。

        每个金额组按以下规则处理：

        1. 先优先匹配 Key word 相同的记录；
        2. 匹配完相同 Key word 后：
           - 如果两边剩余数量相同，则按行号顺序一一配对，
             统一记为 Keyword Difference，并把两边
             keyword_conflict 设为 True；
           - 如果两边剩余数量不同，则全部保留给后续
             一对多、多对一或其他匹配阶段；
        3. 已经匹配的记录不会再次参与。

        返回：
            (正常 One-to-One 组数, Keyword Difference 组数)
        """
        one_to_one_groups = 0
        keyword_difference_groups = 0

        sheet1_by_amount = {}
        sheet2_by_amount = {}

        for record in self.sheet1:
            if not record.matched:
                sheet1_by_amount.setdefault(
                    record.amount,
                    [],
                ).append(record)

        for record in self.sheet2:
            if not record.matched:
                sheet2_by_amount.setdefault(
                    record.amount,
                    [],
                ).append(record)

        common_amounts = sorted(
            set(sheet1_by_amount)
            & set(sheet2_by_amount)
        )

        for amount in common_amounts:
            left_group = sorted(
                sheet1_by_amount[amount],
                key=lambda record: record.row,
            )
            right_group = sorted(
                sheet2_by_amount[amount],
                key=lambda record: record.row,
            )

            # 第一步：同金额组内，优先匹配 Key word 相同的记录。
            for right in right_group:
                if right.matched:
                    continue

                matched_left = None

                for left in left_group:
                    if left.matched:
                        continue

                    if one_to_one_keywords_match(
                        left,
                        right,
                    ):
                        matched_left = left
                        break

                if matched_left is None:
                    continue

                self.mark_match(
                    left_records=[matched_left],
                    right_records=[right],
                    match_type="One-to-One",
                )

                self.set_one_to_one_keyword_result(
                    left=matched_left,
                    right=right,
                    keyword_matched=True,
                )

                one_to_one_groups += 1

            # 第二步：只有两边剩余数量相同，才直接一一配对。
            remaining_left = [
                record
                for record in left_group
                if not record.matched
            ]
            remaining_right = [
                record
                for record in right_group
                if not record.matched
            ]

            if not remaining_left:
                continue

            if len(remaining_left) != len(remaining_right):
                continue

            for left, right in zip(
                remaining_left,
                remaining_right,
            ):
                self.mark_match(
                    left_records=[left],
                    right_records=[right],
                    match_type="Keyword Difference",
                )

                for record in (left, right):
                    record.keyword_match = False
                    record.keyword_conflict = True
                    record.review_required = False
                    record.review_reason = ""

                keyword_difference_groups += 1

        return (
            one_to_one_groups,
            keyword_difference_groups,
        )

    @staticmethod
    def get_same_direction_candidates(
        target,
        records,
    ):
        target_sign = amount_sign(
            target.amount
        )

        if target_sign == 0:
            return []

        candidates = [
            record
            for record in records
            if (
                not record.matched
                and amount_sign(record.amount)
                == target_sign
            )
        ]

        candidates.sort(
            key=lambda record: abs(record.amount)
        )

        return candidates

    @staticmethod
    def find_combination(
        target,
        candidates,
        group_size: int,
        combination_validator: Optional[
            Callable[[list], bool]
        ] = None,
    ) -> Optional[list]:
        if group_size < 2:
            return None

        if len(candidates) < group_size:
            return None

        target_value = abs(target.amount)

        candidate_values = [
            abs(record.amount)
            for record in candidates
        ]

        selected_indexes = []

        def search(
            start_index: int,
            remaining_count: int,
            current_sum: Decimal,
        ):
            if remaining_count == 0:
                if not amount_equal(
                    current_sum,
                    target_value,
                ):
                    return None

                selected_records = [
                    candidates[index]
                    for index in selected_indexes
                ]

                if (
                    combination_validator
                    is not None
                    and not combination_validator(
                        selected_records
                    )
                ):
                    return None

                return list(
                    selected_indexes
                )

            available_count = (
                len(candidates)
                - start_index
            )

            if available_count < remaining_count:
                return None

            if current_sum > target_value:
                return None

            smallest_possible = (
                current_sum
                + sum(
                    candidate_values[
                        start_index:
                        start_index
                        + remaining_count
                    ],
                    Decimal("0"),
                )
            )

            if smallest_possible > target_value:
                return None

            largest_possible = (
                current_sum
                + sum(
                    candidate_values[
                        len(candidate_values)
                        - remaining_count:
                    ],
                    Decimal("0"),
                )
            )

            if largest_possible < target_value:
                return None

            last_possible_index = (
                len(candidates)
                - remaining_count
            )

            for index in range(
                start_index,
                last_possible_index + 1,
            ):
                value = candidate_values[index]
                new_sum = current_sum + value

                if new_sum > target_value:
                    break

                selected_indexes.append(index)

                result = search(
                    start_index=index + 1,
                    remaining_count=(
                        remaining_count - 1
                    ),
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
            candidates[index]
            for index in matched_indexes
        ]

    def one_to_many_match(
        self,
        group_size: int,
    ) -> int:
        """
        Sheet1 一笔对应 Sheet2 多笔，最多 1↔10。

        执行顺序：
        1. 读取目标 Sheet1 的关键词集合；
        2. 直接从 Sheet2 Key word 索引取得同业务候选；
        3. 再筛选未匹配、同金额方向的记录；
        4. 最后搜索金额完全相等的固定笔数组合。

        每条 Sheet2 的单个 Key word 都必须包含在
        目标 Sheet1 的 Key word 集合中。
        """
        matched_groups = 0

        match_type = build_match_type(
            left_size=1,
            right_size=group_size,
        )

        for left in self.sheet1:
            if left.matched:
                continue

            left_keywords = parse_sheet1_keywords(
                left
            )

            if not left_keywords:
                continue

            # 从多个关键词业务池中合并候选，并按对象身份去重。
            indexed_records = []
            seen_ids = set()

            for keyword in left_keywords:
                for right in self.sheet2_keyword_index.get(
                    keyword,
                    [],
                ):
                    record_id = id(right)

                    if record_id in seen_ids:
                        continue

                    seen_ids.add(record_id)
                    indexed_records.append(right)

            candidates = self.filter_candidates(
                target=left,
                records=indexed_records,
            )

            combination = self.find_combination(
                target=left,
                candidates=candidates,
                group_size=group_size,
            )

            if combination is None:
                continue

            self.mark_match(
                left_records=[left],
                right_records=combination,
                match_type=match_type,
            )

            matched_groups += 1

        return matched_groups
    def many_to_one_match(
        self,
        group_size: int,
    ) -> int:
        """
        Sheet1 多笔对应 Sheet2 一笔，最多 10↔1。

        执行顺序：
        1. 读取目标 Sheet2 的单个 Key word；
        2. 直接从 Sheet1 Key word 索引取得同业务候选；
        3. 再筛选未匹配、同金额方向的记录；
        4. 最后搜索金额完全相等的固定笔数组合。

        每条候选 Sheet1 的关键词集合中，
        都包含该 Sheet2 Key word。
        """
        matched_groups = 0

        match_type = build_match_type(
            left_size=group_size,
            right_size=1,
        )

        for right in self.sheet2:
            if right.matched:
                continue

            right_keyword = get_sheet2_keyword(
                right
            )

            if not right_keyword:
                continue

            indexed_records = (
                self.sheet1_keyword_index.get(
                    right_keyword,
                    [],
                )
            )

            candidates = self.filter_candidates(
                target=right,
                records=indexed_records,
            )

            combination = self.find_combination(
                target=right,
                candidates=candidates,
                group_size=group_size,
            )

            if combination is None:
                continue

            self.mark_match(
                left_records=combination,
                right_records=[right],
                match_type=match_type,
            )

            matched_groups += 1

        return matched_groups
    @staticmethod
    def groups_share_keyword(
        left_group,
        right_group,
    ) -> bool:
        """
        判断一个多对多组合是否属于同一业务关键词池。

        规则：
        1. 每条 Sheet2 必须有且只有一个 Key word；
        2. 每条 Sheet2 的 Key word 至少出现在一条 Sheet1 的关键词集合中；
        3. 每条 Sheet1 至少包含所选 Sheet2 的一个 Key word；
        4. 任一方关键词为空时不允许进入多对多匹配。
        """
        right_keywords = [
            get_sheet2_keyword(record)
            for record in right_group
        ]

        if any(
            not keyword
            for keyword in right_keywords
        ):
            return False

        right_keyword_set = set(
            right_keywords
        )

        left_keyword_sets = [
            parse_sheet1_keywords(record)
            for record in left_group
        ]

        if any(
            not keyword_set
            for keyword_set in left_keyword_sets
        ):
            return False

        # 每个 Sheet2 关键词都必须在至少一条 Sheet1 中出现。
        for keyword in right_keyword_set:
            if not any(
                keyword in keyword_set
                for keyword_set in left_keyword_sets
            ):
                return False

        # 每条 Sheet1 都必须至少关联到一个 Sheet2 关键词。
        for keyword_set in left_keyword_sets:
            if not (
                keyword_set
                & right_keyword_set
            ):
                return False

        return True

    @staticmethod
    def get_unmatched_records_by_sign(
        records,
        sign: int,
    ):
        return [
            record
            for record in records
            if (
                not record.matched
                and amount_sign(record.amount)
                == sign
            )
        ]

    @staticmethod
    def build_combination_index(
        records,
        group_size: int,
    ):
        index = {}

        for group in combinations(
            records,
            group_size,
        ):
            total = sum_records(group)

            index.setdefault(
                total,
                [],
            ).append(group)

        return index

    def many_to_many_match(
        self,
        left_size: int,
        right_size: int,
    ) -> int:
        """
        执行关键词优先的多对多匹配。

        当前范围：
            2↔2
            2↔3 / 3↔2
            2↔4 / 4↔2
            2↔5 / 5↔2
            2↔6 / 6↔2

        优化后的算法：

        1. 按单一 Key word 建立业务池；
        2. 正数与负数分别处理；
        3. 两边所有候选都必须属于同一个 Key word；
        4. 始终先为“两笔的一侧”建立金额索引；
        5. 另一侧只枚举一次固定笔数组合；
        6. 用金额合计直接查找，不做两层组合嵌套；
        7. 金额必须完全相等；
        8. 匹配成功后立即锁定记录。

        这样 2↔6 或 6↔2 时，不再出现：
            所有两笔组合 × 所有六笔组合
        的巨大笛卡尔积。
        """
        match_type = build_match_type(
            left_size=left_size,
            right_size=right_size,
        )

        matched_groups = 0

        common_keywords = sorted(
            set(self.sheet1_keyword_index)
            & set(self.sheet2_keyword_index)
        )

        print(
            f"  Searching {match_type} "
            f"across {len(common_keywords)} keyword pool(s)..."
        )

        for keyword in common_keywords:
            for sign in (1, -1):
                left_records = [
                    record
                    for record in self.sheet1_keyword_index.get(
                        keyword,
                        [],
                    )
                    if (
                        not record.matched
                        and amount_sign(record.amount) == sign
                    )
                ]

                right_records = [
                    record
                    for record in self.sheet2_keyword_index.get(
                        keyword,
                        [],
                    )
                    if (
                        not record.matched
                        and amount_sign(record.amount) == sign
                    )
                ]

                if len(left_records) < left_size:
                    continue

                if len(right_records) < right_size:
                    continue

                # --------------------------------------------------
                # 情况一：左边是两笔。
                # 为 Sheet1 两笔组合建立金额索引，
                # 再枚举 Sheet2 的 right_size 笔组合。
                # --------------------------------------------------
                if left_size == 2:
                    left_pair_index = {}

                    for left_group in combinations(
                        left_records,
                        2,
                    ):
                        total = sum_records(
                            left_group
                        )

                        left_pair_index.setdefault(
                            total,
                            [],
                        ).append(left_group)

                    for right_group in combinations(
                        right_records,
                        right_size,
                    ):
                        if any(
                            record.matched
                            for record in right_group
                        ):
                            continue

                        right_total = sum_records(
                            right_group
                        )

                        candidate_left_groups = (
                            left_pair_index.get(
                                right_total,
                                [],
                            )
                        )

                        matched_left_group = None

                        for left_group in candidate_left_groups:
                            if any(
                                record.matched
                                for record in left_group
                            ):
                                continue

                            # 双重保险：两边合计必须完全一致。
                            if not amount_equal(
                                sum_records(left_group),
                                right_total,
                            ):
                                continue

                            matched_left_group = (
                                left_group
                            )
                            break

                        if matched_left_group is None:
                            continue

                        self.mark_match(
                            left_records=list(
                                matched_left_group
                            ),
                            right_records=list(
                                right_group
                            ),
                            match_type=match_type,
                        )

                        matched_groups += 1

                # --------------------------------------------------
                # 情况二：右边是两笔。
                # 为 Sheet2 两笔组合建立金额索引，
                # 再枚举 Sheet1 的 left_size 笔组合。
                # --------------------------------------------------
                elif right_size == 2:
                    right_pair_index = {}

                    for right_group in combinations(
                        right_records,
                        2,
                    ):
                        total = sum_records(
                            right_group
                        )

                        right_pair_index.setdefault(
                            total,
                            [],
                        ).append(right_group)

                    for left_group in combinations(
                        left_records,
                        left_size,
                    ):
                        if any(
                            record.matched
                            for record in left_group
                        ):
                            continue

                        left_total = sum_records(
                            left_group
                        )

                        candidate_right_groups = (
                            right_pair_index.get(
                                left_total,
                                [],
                            )
                        )

                        matched_right_group = None

                        for right_group in candidate_right_groups:
                            if any(
                                record.matched
                                for record in right_group
                            ):
                                continue

                            if not amount_equal(
                                left_total,
                                sum_records(right_group),
                            ):
                                continue

                            matched_right_group = (
                                right_group
                            )
                            break

                        if matched_right_group is None:
                            continue

                        self.mark_match(
                            left_records=list(
                                left_group
                            ),
                            right_records=list(
                                matched_right_group
                            ),
                            match_type=match_type,
                        )

                        matched_groups += 1

                else:
                    raise ValueError(
                        "Optimized many-to-many currently "
                        "requires one side to have exactly 2 records."
                    )

        print(
            f"  Finished {match_type}: "
            f"{matched_groups} group(s)"
        )

        return matched_groups
    def print_summary(
        self,
        results,
    ) -> None:
        remaining_sheet1 = sum(
            1
            for record in self.sheet1
            if not record.matched
        )

        remaining_sheet2 = sum(
            1
            for record in self.sheet2
            if not record.matched
        )

        matched_sheet1 = (
            len(self.sheet1)
            - remaining_sheet1
        )

        matched_sheet2 = (
            len(self.sheet2)
            - remaining_sheet2
        )

        print()
        print("=" * 56)
        print("Matching Summary")
        print("=" * 56)

        for match_type, count in results.items():
            print(
                f"{match_type:<22}: {count}"
            )

        print("-" * 56)

        print(
            f"Total Sheet1 rows      : "
            f"{len(self.sheet1)}"
        )

        print(
            f"Total Sheet2 rows      : "
            f"{len(self.sheet2)}"
        )

        print(
            f"Matched Sheet1 rows    : "
            f"{matched_sheet1}"
        )

        print(
            f"Matched Sheet2 rows    : "
            f"{matched_sheet2}"
        )

        print(
            f"Remaining Sheet1 rows  : "
            f"{remaining_sheet1}"
        )

        print(
            f"Remaining Sheet2 rows  : "
            f"{remaining_sheet2}"
        )

        print("=" * 56)