from difference_analyzer import (
    analyze_difference,
    print_difference_summary,
)
from excel_io import load_excel, save_results
from matcher import Matcher


INPUT_FILE = "data.xlsx"
OUTPUT_FILE = "result_reconciliation.xlsx"


def main():
    print("=" * 48)
    print(" Excel Reconciliation Tool")
    print("=" * 48)

    workbook, sheet1_records, sheet2_records = load_excel(
        INPUT_FILE
    )

    print(f"Input file     : {INPUT_FILE}")
    print(f"Sheet1 records : {len(sheet1_records)}")
    print(f"Sheet2 records : {len(sheet2_records)}")
    print("Matching rule  : Exact amount only")
    print()

    matcher = Matcher(
        sheet1_records,
        sheet2_records,
    )
    matcher.run()

    difference_result = analyze_difference(
        sheet1_records=sheet1_records,
        sheet2_records=sheet2_records,
    )

    print_difference_summary(
        difference_result
    )

    save_results(
        workbook=workbook,
        sheet1_records=sheet1_records,
        sheet2_records=sheet2_records,
        output_filename=OUTPUT_FILE,
        difference_result=difference_result,
    )

    print()
    print(f"Result saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()