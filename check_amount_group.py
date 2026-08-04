from excel_io import load_excel
from matcher import (
    Matcher,
    get_sheet2_keyword,
    parse_sheet1_keywords,
)


INPUT_FILE = "data.xlsx"
TARGET_AMOUNT_TEXT = "-10000000"


def format_sheet1_keyword(record):
    keywords = sorted(
        parse_sheet1_keywords(record)
    )
    return "; ".join(keywords) or "<blank>"


def format_sheet2_keyword(record):
    return (
        get_sheet2_keyword(record)
        or "<blank>"
    )


def print_records(title, records, keyword_getter):
    print()
    print(title)
    print("-" * 88)

    if not records:
        print("No records found.")
        return

    for record in sorted(
        records,
        key=lambda item: item.row,
    ):
        print(
            f"Row {record.row:<5} "
            f"Amount={record.amount!r:<18} "
            f"Matched={str(record.matched):<5} "
            f"Type={record.match_type or '<blank>':<24} "
            f"Partners={record.partners!s:<18} "
            f"Keyword={keyword_getter(record)}"
        )


def main():
    _, sheet1_records, sheet2_records = load_excel(
        INPUT_FILE
    )

    target_amount = next(
        record.amount
        for record in sheet1_records + sheet2_records
        if str(record.amount) == TARGET_AMOUNT_TEXT
    )

    before_sheet1 = [
        record
        for record in sheet1_records
        if record.amount == target_amount
    ]

    before_sheet2 = [
        record
        for record in sheet2_records
        if record.amount == target_amount
    ]

    print("=" * 88)
    print(f"Target amount: {target_amount!r}")
    print("=" * 88)

    print_records(
        "BEFORE MATCHER — Sheet1",
        before_sheet1,
        format_sheet1_keyword,
    )

    print_records(
        "BEFORE MATCHER — Sheet2",
        before_sheet2,
        format_sheet2_keyword,
    )

    matcher = Matcher(
        sheet1_records,
        sheet2_records,
    )
    matcher.run()

    after_sheet1 = [
        record
        for record in sheet1_records
        if record.amount == target_amount
    ]

    after_sheet2 = [
        record
        for record in sheet2_records
        if record.amount == target_amount
    ]

    print()
    print("=" * 88)
    print("AFTER MATCHER")
    print("=" * 88)

    print_records(
        "AFTER MATCHER — Sheet1",
        after_sheet1,
        format_sheet1_keyword,
    )

    print_records(
        "AFTER MATCHER — Sheet2",
        after_sheet2,
        format_sheet2_keyword,
    )

    remaining_sheet1 = [
        record
        for record in after_sheet1
        if not record.matched
    ]

    remaining_sheet2 = [
        record
        for record in after_sheet2
        if not record.matched
    ]

    print()
    print("=" * 88)
    print("GROUP COUNTS")
    print("=" * 88)
    print(
        f"Before: Sheet1={len(before_sheet1)}, "
        f"Sheet2={len(before_sheet2)}"
    )
    print(
        f"After unmatched: Sheet1={len(remaining_sheet1)}, "
        f"Sheet2={len(remaining_sheet2)}"
    )


if __name__ == "__main__":
    main()