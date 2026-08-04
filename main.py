from difference_analyzer import (
    analyze_difference,
    print_difference_summary,
)
from excel_io import load_excel, save_results
from final_yellow_matcher import (
    match_final_yellow_records,
    print_final_yellow_summary,
)
from keyword_difference_matcher import (
    match_keyword_differences,
)
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

    # 第一阶段：原有匹配流程。
    matcher = Matcher(
        sheet1_records,
        sheet2_records,
    )
    matcher.run()

    # 第二阶段：相同 Key word 的金额差额补齐。
    keyword_difference_groups = (
        match_keyword_differences(
            sheet1_records=sheet1_records,
            sheet2_records=sheet2_records,
        )
    )

    print()
    print(
        "Keyword-Difference groups: "
        f"{keyword_difference_groups}"
    )

    # 第三阶段：最终黄色区匹配。
    # 不要求 Key word 相同，也不要求 Key word 为空。
    # 只处理尚未匹配记录，支持 1↔1、1↔2...1↔6
    # 及对称方向，并支持 2↔2。
    final_yellow_results = (
        match_final_yellow_records(
            sheet1_records=sheet1_records,
            sheet2_records=sheet2_records,
        )
    )

    print_final_yellow_summary(
        final_yellow_results
    )

    # 最后只分析仍未匹配的黄色记录。
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