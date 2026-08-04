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
        "key_word_column": 4,     # D列
        "match_type_column": 5,   # E列
        "partner_column": 6,      # F列
        "review_column": 7,       # G列
    },
    "Sheet2": {
        "amount_column": 6,       # F列
        "key_word_column": 7,     # G列
        "match_type_column": 8,   # H列
        "partner_column": 9,      # I列
        "review_column": 10,      # J列
    },
}


# ============================================================
# Final Color Rules
#
# 绿色：
#     Key word相同并成功匹配。
#
# 浅棕色：
#     Key word不同，但按前置、较可靠业务规则匹配成功。
#     包括：
#     1. 唯一同金额的一对一，但Key word不同；
#     2. 重复金额组先消除相同Key word后，
#        两边剩余数量相同；
#     3. 相同业务Key word存在金额差额，
#        由单侧最多3条空白Key word记录补齐。
#
# 蓝色：
#     最终黄色区中，不要求Key word相同而匹配成功；
#     只支持1↔1到1↔6及其反向。
#
# 黄色：
#     最终仍未匹配。
# ============================================================


MATCHED_KEYWORD_SAME_FILL = PatternFill(
    start_color="93C47D",
    end_color="93C47D",
    fill_type="solid",
)

MATCHED_KEYWORD_DIFFERENT_FILL = PatternFill(
    start_color="E3D5CA",
    end_color="E3D5CA",
    fill_type="solid",
)

FINAL_YELLOW_MATCH_FILL = PatternFill(
    start_color="BDD7EE",
    end_color="BDD7EE",
    fill_type="solid",
)

UNMATCHED_FILL = PatternFill(
    start_color="FFE699",
    end_color="FFE699",
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
    """按照工作表配置读取金额、Key word及原始辅助信息。"""

    config = get_sheet_config(sheet.title)
    amount_column = config["amount_column"]
    key_word_column = config["key_word_column"]

    headers = read_headers(
        sheet=sheet,
        last_data_column=key_word_column,
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

        for column in range(1, key_word_column + 1):
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
    根据最终状态返回整行颜色。

    未匹配：
        黄色

    Final Yellow Match：
        蓝色

    其他已匹配且 keyword_conflict=True：
        浅棕色

    其他已匹配记录：
        绿色
    """

    if not record.matched:
        return UNMATCHED_FILL

    if record.match_type.startswith(
        "Final Yellow Match"
    ):
        return FINAL_YELLOW_MATCH_FILL

    if record.keyword_conflict:
        return MATCHED_KEYWORD_DIFFERENT_FILL

    return MATCHED_KEYWORD_SAME_FILL


def get_result_last_column(sheet):
    """取得该工作表结果区域的最后一列。"""

    config = get_sheet_config(sheet.title)
    return config["review_column"]


def fill_row(sheet, row, fill):
    """给整行原始数据及匹配结果着色。"""

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
    """写入匹配结果，并按最终四色规则着色。"""

    config = get_sheet_config(sheet.title)

    match_type_column = config["match_type_column"]
    partner_column = config["partner_column"]
    review_column = config["review_column"]

    for record in records:
        fill_row(
            sheet=sheet,
            row=record.row,
            fill=get_fill(record),
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


def set_result_column_widths(sheet):
    """设置结果列宽。"""

    config = get_sheet_config(sheet.title)

    widths = {
        config["key_word_column"]: 24,
        config["match_type_column"]: 30,
        config["partner_column"]: 24,
        config["review_column"]: 24,
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
    """
    把匹配结果写入工作簿。

    difference_result 参数继续保留，以兼容 main.py；
    但它不再改变任何行的颜色。
    """

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

    set_result_column_widths(sheet1)
    set_result_column_widths(sheet2)

    workbook.save(output_filename)