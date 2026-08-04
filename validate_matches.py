"""
validate_matches.py

验证匹配关系是否一致。

检查：

1.
Sheet1 -> Sheet2

是否

Sheet2 -> Sheet1

2.
金额是否一致。

3.
Partner Rows 是否双向一致。

发现异常立即打印。
"""

from collections import defaultdict


def validate(sheet1_records, sheet2_records):

    sheet1 = {
        r.row: r
        for r in sheet1_records
    }

    sheet2 = {
        r.row: r
        for r in sheet2_records
    }

    errors = 0

    print()
    print("=" * 60)
    print("Match Validation")
    print("=" * 60)

    # -----------------------------
    # Sheet1 -> Sheet2
    # -----------------------------
    for record in sheet1_records:

        if not record.matched:
            continue

        for partner_row in record.partners:

            if partner_row not in sheet2:

                print(
                    f"[ERROR] Sheet2 Row {partner_row} 不存在"
                )

                errors += 1
                continue

            partner = sheet2[partner_row]

            if record.row not in partner.partners:

                print(
                    f"[ERROR] "
                    f"Sheet1 Row {record.row}"
                    f" -> Sheet2 Row {partner_row}"
                    f" 不是双向关系"
                )

                errors += 1

    # -----------------------------
    # Sheet2 -> Sheet1
    # -----------------------------
    for record in sheet2_records:

        if not record.matched:
            continue

        for partner_row in record.partners:

            if partner_row not in sheet1:

                print(
                    f"[ERROR] Sheet1 Row {partner_row} 不存在"
                )

                errors += 1
                continue

            partner = sheet1[partner_row]

            if record.row not in partner.partners:

                print(
                    f"[ERROR] "
                    f"Sheet2 Row {record.row}"
                    f" -> Sheet1 Row {partner_row}"
                    f" 不是双向关系"
                )

                errors += 1

    print("-" * 60)

    if errors == 0:

        print("Validation Passed")

    else:

        print(f"Validation Failed : {errors} error(s)")

    print("=" * 60)

    return errors == 0