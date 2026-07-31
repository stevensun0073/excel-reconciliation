import re
from decimal import Decimal
from itertools import combinations
from typing import Callable, Optional


MAX_GROUP_SIZE = 6

NUMBER_NAMES = {
    1: "One",
    2: "Two",
    3: "Three",
    4: "Four",
    5: "Five",
    6: "Six",
}

# 多对多匹配的执行顺序。
#
# 必须优先执行较简单的组合，避免本来可以按 2↔2
# 解释的记录被提前放入 2↔4 等更复杂的组合。
MANY_TO_MANY_PATTERNS = [
    (2, 2),
    (2, 3),
    (3, 2),
    (2, 4),
    (4, 2),
]


def amount_equal(left: Decimal, right: Decimal) -> bool:
    """
    判断两个金额是否完全相等。

    银行对账不允许金额误差。
    例如 100.00 与 99.99 不属于相同金额。
    """
    return left == right


def amount_sign(amount: Decimal) -> int:
    """
    返回金额方向。

    正数：1
    负数：-1
    零：0
    """
    if amount > 0:
        return 1

    if amount < 0:
        return -1

    return 0


def sum_records(records) -> Decimal:
    """计算一组记录的金额合计。"""
    return sum(
        (record.amount for record in records),
        Decimal("0"),
    )


def build_match_type(
    left_size: int,
    right_size: int,
) -> str:
    """生成 One-to-Two、Two-to-Four 等匹配名称。"""
    return (
        f"{NUMBER_NAMES[left_size]}"
        f"-to-"
        f"{NUMBER_NAMES[right_size]}"
    )


def extract_first_three_words(value) -> list[str]:
    """
    从一个字段中提取前三个单词。

    处理规则：
    1. 转为大写；
    2. 标点符号作为分隔符；
    3. 保留英文字母和数字；
    4. 最多取前三个单词。
    """
    if value is None:
        return []

    words = re.findall(
        r"[A-Z0-9]+",
        str(value).upper(),
    )

    return words[:3]


def build_sheet2_keyword_set(record) -> set[str]:
    """
    为一条 Sheet2 记录建立关键词集合。

    分别取以下两列的前三个单词：

    - Recipient's Account Name
    - Description

    然后合并成一个集合。
    """
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
    """
    判断一组 Sheet2 记录是否至少共享两个单词。

    每条记录都使用：

    - Recipient's Account Name 的前三个单词；
    - Description 的前三个单词。

    只有至少两个单词同时出现在组合中的每一条记录里，
    才允许 Sheet1 一笔对应 Sheet2 多笔。
    """
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

    当前自动匹配范围：

    一对一：
        1↔1

    一对多和多对一：
        1↔2 至 1↔6
        2↔1 至 6↔1

    有限多对多：
        2↔2
        2↔3
        3↔2
        2↔4
        4↔2

    所有组合匹配必须满足：

    1. 两边金额合计完全相等；
    2. 同一组合内所有金额方向一致；
    3. 正数只能与正数匹配；
    4. 负数只能与负数匹配；
    5. 每条记录只能匹配一次；
    6. 优先处理笔数较少的组合。
    """

    def __init__(
        self,
        sheet1_records,
        sheet2_records,
    ):
        self.sheet1 = sheet1_records
        self.sheet2 = sheet2_records

    @staticmethod
    def mark_match(
        left_records,
        right_records,
        match_type: str,
    ) -> None:
        """
        将左右两组记录登记为已匹配。

        每条记录的 partners 保存对方工作表中的行号。
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

    def run(self):
        """按照由简单到复杂的顺序执行全部匹配。"""
        results = {}

        # 第一层：一对一
        results["One-to-One"] = (
            self.one_to_one_match()
        )

        # 第二层：一对多及多对一
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

        # 第三层：有限多对多
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

    def one_to_one_match(self) -> int:
        """
        执行一对一匹配。

        只有金额完全相等，才确认匹配。
        """
        matched_groups = 0

        for left in self.sheet1:
            if left.matched:
                continue

            for right in self.sheet2:
                if right.matched:
                    continue

                if not amount_equal(
                    left.amount,
                    right.amount,
                ):
                    continue

                self.mark_match(
                    left_records=[left],
                    right_records=[right],
                    match_type="One-to-One",
                )

                matched_groups += 1
                break

        return matched_groups

    @staticmethod
    def get_same_direction_candidates(
        target,
        records,
    ):
        """
        取得与目标金额同方向的未匹配候选记录。

        正数目标只能使用正数候选；
        负数目标只能使用负数候选；
        零金额不参加组合匹配。
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
    def find_combination(
        target,
        candidates,
        group_size: int,
        combination_validator: Optional[
            Callable[[list], bool]
        ] = None,
    ) -> Optional[list]:
        """
        从候选记录中寻找固定笔数的金额组合。

        例如 group_size=3：

        寻找三条候选记录，使其金额合计
        与目标金额完全相等。

        采用递归回溯及金额范围剪枝。
        """
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
                    # 金额相等但文字条件不合格时，
                    # 继续搜索其他金额组合。
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
        Sheet1 一笔对应 Sheet2 多笔。

        支持：

        1↔2
        1↔3
        1↔4
        1↔5
        1↔6

        除金额完全相等之外，参与组合的每条 Sheet2 记录：

        1. 取 Recipient's Account Name 的前三个单词；
        2. 取 Description 的前三个单词；
        3. 合并后，所有记录必须至少共享两个单词。
        """
        matched_groups = 0

        match_type = build_match_type(
            left_size=1,
            right_size=group_size,
        )

        for left in self.sheet1:
            if left.matched:
                continue

            candidates = (
                self.get_same_direction_candidates(
                    target=left,
                    records=self.sheet2,
                )
            )

            combination = self.find_combination(
                target=left,
                candidates=candidates,
                group_size=group_size,
                combination_validator=(
                    sheet2_group_has_common_words
                ),
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
        Sheet1 多笔对应 Sheet2 一笔。

        支持：

        2↔1
        3↔1
        4↔1
        5↔1
        6↔1
        """
        matched_groups = 0

        match_type = build_match_type(
            left_size=group_size,
            right_size=1,
        )

        for right in self.sheet2:
            if right.matched:
                continue

            candidates = (
                self.get_same_direction_candidates(
                    target=right,
                    records=self.sheet1,
                )
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
    def get_unmatched_records_by_sign(
        records,
        sign: int,
    ):
        """
        取得指定方向的未匹配记录。

        sign=1：正数
        sign=-1：负数
        """
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
        """
        为固定笔数组合建立精确金额索引。

        索引格式：

        {
            Decimal金额合计: [
                记录组合1,
                记录组合2,
                ...
            ]
        }

        不进行四舍五入，也不允许一分钱误差。
        """
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
        执行有限多对多匹配。

        当前调用范围：

        2↔2
        2↔3
        3↔2
        2↔4
        4↔2

        算法：

        1. 正数和负数分别处理；
        2. 为右侧组合建立精确金额索引；
        3. 遍历左侧组合；
        4. 使用 Decimal 合计直接查找；
        5. 最后再次进行精确金额核对；
        6. 已使用记录不能再次使用。
        """
        match_type = build_match_type(
            left_size=left_size,
            right_size=right_size,
        )

        matched_groups = 0

        # 正数和负数分别匹配，
        # 保证不会出现方向混合。
        for sign in (1, -1):
            left_records = (
                self.get_unmatched_records_by_sign(
                    records=self.sheet1,
                    sign=sign,
                )
            )

            right_records = (
                self.get_unmatched_records_by_sign(
                    records=self.sheet2,
                    sign=sign,
                )
            )

            if len(left_records) < left_size:
                continue

            if len(right_records) < right_size:
                continue

            # 右侧组合只建立一次精确金额索引。
            right_index = (
                self.build_combination_index(
                    records=right_records,
                    group_size=right_size,
                )
            )

            used_left_ids = set()
            used_right_ids = set()

            for left_group in combinations(
                left_records,
                left_size,
            ):
                left_ids = {
                    id(record)
                    for record in left_group
                }

                if left_ids & used_left_ids:
                    continue

                left_total = sum_records(
                    left_group
                )

                candidate_groups = (
                    right_index.get(
                        left_total,
                        [],
                    )
                )

                matched_right_group = None

                for right_group in candidate_groups:
                    right_ids = {
                        id(record)
                        for record in right_group
                    }

                    if right_ids & used_right_ids:
                        continue

                    right_total = sum_records(
                        right_group
                    )

                    if not amount_equal(
                        left_total,
                        right_total,
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

                used_left_ids.update(
                    id(record)
                    for record in left_group
                )

                used_right_ids.update(
                    id(record)
                    for record
                    in matched_right_group
                )

                matched_groups += 1

        return matched_groups

    def print_summary(
        self,
        results,
    ) -> None:
        """打印各类匹配数量和剩余记录数量。"""
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