from difference_analyzer import (
    analyze_difference,
    print_difference_summary,
)
from excel_io import load_excel, save_results
from matcher import Matcher
from remaining_matcher import (
    match_remaining_records,
    print_remaining_match_summary,
)
from validate_matches import validate


INPUT_FILE = "data.xlsx"
OUTPUT_FILE = "result_reconciliation.xlsx"


def main():
    print("=" * 48)
    print(" Excel Reconciliation Tool")
    print("=" * 48)

    (
        workbook,
        sheet1_records,
        sheet2_records,
    ) = load_excel(INPUT_FILE)

    print(f"Input file     : {INPUT_FILE}")
    print(
        f"Sheet1 records : "
        f"{len(sheet1_records)}"
    )
    print(
        f"Sheet2 records : "
        f"{len(sheet2_records)}"
    )
    print("Matching rule  : Exact amount only")
    print()

    # ========================================================
    # 第一阶段：
    # 原有绿色和棕色规则
    # 完全保持不变
    # ========================================================

    matcher = Matcher(
        sheet1_records,
        sheet2_records,
    )

    matcher.run()

    # ========================================================
    # 第二阶段：
    # 剩余记录规则
    #
    # 淡蓝色：
    # 1. 唯一金额一对一
    # 2. Sheet1同流水号多对一
    #
    # 淡黄色：
    # 3. 同一张表内相同Key word优先
    # 4. 剩余记录纯金额组合
    # 5. 支持1↔2至1↔6
    #
    # 业务限制：
    # Sheet1 BANK只能匹配Sheet2 Charging
    #
    # 深黄色：
    # 最终仍未匹配
    # ========================================================

    remaining_result = (
        match_remaining_records(
            sheet1_records=sheet1_records,
            sheet2_records=sheet2_records,
        )
    )

    print_remaining_match_summary(
        remaining_result
    )

    # ========================================================
    # 第三阶段：
    # 验证匹配关系
    # ========================================================

    validation_passed = validate(
        sheet1_records,
        sheet2_records,
    )

    if not validation_passed:
        raise RuntimeError(
            "Match validation failed. "
            "Result file was not generated."
        )

    # ========================================================
    # 第四阶段：
    # 分析最终剩余差额
    # ========================================================

    difference_result = analyze_difference(
        sheet1_records=sheet1_records,
        sheet2_records=sheet2_records,
    )

    print_difference_summary(
        difference_result
    )

    # ========================================================
    # 第五阶段：
    # 写入正式结果
    # ========================================================

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