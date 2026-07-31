from decimal import Decimal, InvalidOperation
from typing import TYPE_CHECKING

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from models import Record

if TYPE_CHECKING:
    from difference_analyzer import DifferenceAnalysisResult


HEADER_ROW = 2
START_ROW = 3

SHEET_CONFIG = {
    "Sheet1": {
        "amount_column": 3,       # C列
        "match_type_column": 4,   # D列
        "partner_column": 5,      # E列
        "review_column": 6,       # F列
    },
    "Sheet2": {
        "amount_column": 6,       # F列
        "match_type_column": 7,   # G列
        "partner_column": 8,      # H列
        "review_column": 9,       # I列
    },
}


# 一对一：柔和绿色
ONE_TO_ONE_FILL = PatternFill(
    start_color="B7D7A8",
    end_color="B7D7A8",
    fill_type="solid",
)

# 所有组合匹配：柔和蓝色
COMBINATION_FILL = PatternFill(
    start_color="BDD7EE",
    end_color="BDD7EE",
    fill_type="solid",
)

# 未匹配：金黄色
UNMATCHED_FILL = PatternFill(
    start_color="FFE699",
    end_color="FFE699",
    fill_type="solid",
)

# Difference Analyzer：浅珊瑚红
DIFFERENCE_CANDIDATE_FILL = PatternFill(
    start_color="F4CCCC",
    end_color="F4CCCC",
    fill_type="solid",
)




def parse_amount(value):
    """把 Excel 金额安全转换为 Decimal。"""

    if value is None:
        return None

    try:
        cleaned_value = str(value).replace(",", "").strip()

        if not cleaned_value:
            return None

        return Decimal(cleaned_value)

    except (InvalidOperation, AttributeError, ValueError):
        return None


def get_sheet_config(sheet_name):
    """取得指定工作表的读取和输出配置。"""

    if sheet_name not in SHEET_CONFIG:
        raise ValueError(
            f"没有为工作表 {sheet_name} 设置读取规则。"
        )

    return SHEET_CONFIG[sheet_name]


def read_headers(sheet, last_data_column):
    """读取第2行中的原始字段名称。"""

    headers = {}

    for column in range(1, last_data_column + 1):
        value = sheet.cell(
            row=HEADER_ROW,
            column=column,
        ).value

        if value is None or str(value).strip() == "":
            header = f"Column {get_column_letter(column)}"
        else:
            header = str(value).strip()

        headers[column] = header

    return headers


def read_sheet(sheet):
    """按照工作表配置读取金额及原始辅助信息。"""

    config = get_sheet_config(sheet.title)
    amount_column = config["amount_column"]

    headers = read_headers(
        sheet=sheet,
        last_data_column=amount_column,
    )

    records = []

    for row in range(START_ROW, sheet.max_row + 1):
        amount = parse_amount(
            sheet.cell(
                row=row,
                column=amount_column,
            ).value
        )

        if amount is None:
            continue

        extra = {}

        for column in range(1, amount_column + 1):
            if column == amount_column:
                continue

            header = headers[column]

            value = sheet.cell(
                row=row,
                column=column,
            ).value

            extra[header] = value

        records.append(
            Record(
                row=row,
                amount=amount,
                source_sheet=sheet.title,
                extra=extra,
            )
        )

    return records


def load_excel(filename):
    """打开工作簿并读取 Sheet1、Sheet2。"""

    workbook = load_workbook(filename)

    for sheet_name in ("Sheet1", "Sheet2"):
        if sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"找不到工作表 {sheet_name}。"
            )

    return (
        workbook,
        read_sheet(workbook["Sheet1"]),
        read_sheet(workbook["Sheet2"]),
    )


def get_fill(record):
    """
    根据最终状态返回颜色。

    绿色：One-to-One
    蓝色：所有其它正式匹配
    黄色：未匹配
    红色：Difference Candidate（最后覆盖）
    """

    if not record.matched:
        return UNMATCHED_FILL

    if record.match_type == "One-to-One":
        return ONE_TO_ONE_FILL

    # 除一对一之外，其它所有正式匹配统一蓝色
    return COMBINATION_FILL


def get_result_last_column(sheet):
    """取得该工作表结果区域的最后一列。"""

    config = get_sheet_config(sheet.title)
    return config["review_column"]


def fill_row(sheet, row, fill):
    """给一整行原始数据及匹配结果着色。"""

    last_column = get_result_last_column(sheet)

    for column in range(1, last_column + 1):
        sheet.cell(
            row=row,
            column=column,
        ).fill = fill


def clear_old_results(sheet):
    """清除上一次运行留下的匹配结果和颜色。"""

    config = get_sheet_config(sheet.title)

    match_type_column = config["match_type_column"]
    partner_column = config["partner_column"]
    review_column = config["review_column"]

    headers = {
        match_type_column: "Match Type",
        partner_column: "Partner Rows",
        review_column: "Review",
    }

    for column, header in headers.items():
        cell = sheet.cell(
            row=HEADER_ROW,
            column=column,
        )

        cell.value = header
        cell.font = Font(bold=True)

    last_column = get_result_last_column(sheet)

    for row in range(START_ROW, sheet.max_row + 1):
        for column in range(1, last_column + 1):
            sheet.cell(
                row=row,
                column=column,
            ).fill = PatternFill()

        sheet.cell(
            row=row,
            column=match_type_column,
        ).value = None

        sheet.cell(
            row=row,
            column=partner_column,
        ).value = None

        sheet.cell(
            row=row,
            column=review_column,
        ).value = None


def write_records(sheet, records):
    """
    写入匹配结果。

    已匹配记录按照匹配类型着色；
    未匹配记录统一标为黄色。
    """

    config = get_sheet_config(sheet.title)

    match_type_column = config["match_type_column"]
    partner_column = config["partner_column"]
    review_column = config["review_column"]

    for record in records:
        fill = get_fill(record)

        if fill is not None:
            fill_row(
                sheet=sheet,
                row=record.row,
                fill=fill,
            )

        if not record.matched:
            continue

        sheet.cell(
            row=record.row,
            column=match_type_column,
        ).value = record.match_type

        sheet.cell(
            row=record.row,
            column=partner_column,
        ).value = ", ".join(
            str(row)
            for row in record.partners
        )

        sheet.cell(
            row=record.row,
            column=review_column,
        ).value = record.review_reason


def apply_difference_candidate_fill(
    workbook,
    difference_result,
):
    """
    把差额分析选中的候选记录标成红色。

    红色优先于未匹配记录原来的黄色。
    """

    if difference_result is None:
        return

    if not difference_result.selected_items:
        return

    for item in difference_result.selected_items:
        if item.source_sheet not in workbook.sheetnames:
            continue

        sheet = workbook[item.source_sheet]

        fill_row(
            sheet=sheet,
            row=item.source_row,
            fill=DIFFERENCE_CANDIDATE_FILL,
        )


def set_result_column_widths(sheet):
    """设置新增结果列的宽度。"""

    config = get_sheet_config(sheet.title)

    widths = {
        config["match_type_column"]: 18,
        config["partner_column"]: 20,
        config["review_column"]: 32,
    }

    for column, width in widths.items():
        column_letter = get_column_letter(column)
        sheet.column_dimensions[column_letter].width = width


def save_results(
    workbook,
    sheet1_records,
    sheet2_records,
    output_filename,
    difference_result=None,
):
    """把匹配结果和差额候选写入工作簿。"""

    sheet1 = workbook["Sheet1"]
    sheet2 = workbook["Sheet2"]

    clear_old_results(sheet1)
    clear_old_results(sheet2)

    write_records(
        sheet=sheet1,
        records=sheet1_records,
    )

    write_records(
        sheet=sheet2,
        records=sheet2_records,
    )

    # 最后标红，使红色覆盖原来的未匹配黄色。
    apply_difference_candidate_fill(
        workbook=workbook,
        difference_result=difference_result,
    )

    set_result_column_widths(sheet1)
    set_result_column_widths(sheet2)

    workbook.save(output_filename)