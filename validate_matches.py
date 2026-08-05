"""
validate_matches.py

验证匹配关系是否一致。

支持两种关系：

1. Sheet1与Sheet2之间的匹配；
2. Sheet1内部冲销组。

检查Partner Rows是否存在并保持双向一致。
"""

from decimal import Decimal


BLUE_REVERSAL_PREFIX = (
    "Blue Same Reference Reversal"
)


def is_sheet1_reversal(record) -> bool:
    return record.match_type.startswith(
        BLUE_REVERSAL_PREFIX
    )


def validate(
    sheet1_records,
    sheet2_records,
):

    sheet1 = {
        record.row: record
        for record in sheet1_records
    }

    sheet2 = {
        record.row: record
        for record in sheet2_records
    }

    errors = 0
    checked_reversal_groups = set()

    print()
    print("=" * 60)
    print("Match Validation")
    print("=" * 60)

    # --------------------------------------------------------
    # Sheet1：
    # 普通匹配检查Sheet2；
    # 内部冲销检查Sheet1。
    # --------------------------------------------------------
    for record in sheet1_records:

        if not record.matched:
            continue

        if is_sheet1_reversal(record):
            group_rows = frozenset(
                [record.row]
                + record.partners
            )

            if len(group_rows) < 2:
                print(
                    f"[ERROR] "
                    f"Sheet1 Row {record.row} "
                    f"内部冲销组少于2行"
                )
                errors += 1
                continue

            if (
                group_rows
                in checked_reversal_groups
            ):
                continue

            checked_reversal_groups.add(
                group_rows
            )

            group_records = []

            for row in sorted(group_rows):
                if row not in sheet1:
                    print(
                        f"[ERROR] "
                        f"Sheet1冲销伙伴 "
                        f"Row {row} 不存在"
                    )
                    errors += 1
                    continue

                group_records.append(
                    sheet1[row]
                )

            if (
                len(group_records)
                != len(group_rows)
            ):
                continue

            expected_rows = set(
                group_rows
            )

            for group_record in group_records:
                if not is_sheet1_reversal(
                    group_record
                ):
                    print(
                        f"[ERROR] "
                        f"Sheet1 Row "
                        f"{group_record.row} "
                        f"冲销类型不一致"
                    )
                    errors += 1

                actual_rows = set(
                    [group_record.row]
                    + group_record.partners
                )

                if (
                    actual_rows
                    != expected_rows
                ):
                    print(
                        f"[ERROR] "
                        f"Sheet1 Row "
                        f"{group_record.row} "
                        f"冲销组Partner Rows不一致"
                    )
                    errors += 1

            group_total = sum(
                (
                    group_record.amount
                    for group_record
                    in group_records
                ),
                Decimal("0"),
            )

            if group_total != Decimal("0"):
                print(
                    f"[ERROR] "
                    f"Sheet1内部冲销组 "
                    f"{sorted(group_rows)} "
                    f"合计不是0："
                    f"{group_total}"
                )
                errors += 1

            continue

        for partner_row in record.partners:

            if partner_row not in sheet2:
                print(
                    f"[ERROR] "
                    f"Sheet2 Row "
                    f"{partner_row} 不存在"
                )
                errors += 1
                continue

            partner = sheet2[
                partner_row
            ]

            if (
                record.row
                not in partner.partners
            ):
                print(
                    f"[ERROR] "
                    f"Sheet1 Row {record.row}"
                    f" -> Sheet2 Row "
                    f"{partner_row}"
                    f" 不是双向关系"
                )
                errors += 1

    # --------------------------------------------------------
    # Sheet2 -> Sheet1
    # --------------------------------------------------------
    for record in sheet2_records:

        if not record.matched:
            continue

        for partner_row in record.partners:

            if partner_row not in sheet1:
                print(
                    f"[ERROR] "
                    f"Sheet1 Row "
                    f"{partner_row} 不存在"
                )
                errors += 1
                continue

            partner = sheet1[
                partner_row
            ]

            if is_sheet1_reversal(
                partner
            ):
                print(
                    f"[ERROR] "
                    f"Sheet2 Row {record.row} "
                    f"错误指向Sheet1内部"
                    f"冲销记录 Row "
                    f"{partner_row}"
                )
                errors += 1
                continue

            if (
                record.row
                not in partner.partners
            ):
                print(
                    f"[ERROR] "
                    f"Sheet2 Row {record.row}"
                    f" -> Sheet1 Row "
                    f"{partner_row}"
                    f" 不是双向关系"
                )
                errors += 1

    print("-" * 60)

    if errors == 0:
        print("Validation Passed")
    else:
        print(
            f"Validation Failed : "
            f"{errors} error(s)"
        )

    print("=" * 60)

    return errors == 0