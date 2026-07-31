from decimal import Decimal, InvalidOperation

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from models import Record


START_ROW = 3
AMOUNT_COLUMN = 2
MATCH_TYPE_COLUMN = 3
PARTNER_COLUMN = 4
REVIEW_COLUMN = 5

ONE_TO_ONE_FILL = PatternFill(
    start_color="D4EDDA",
    end_color="D4EDDA",
    fill_type="solid",
)

COMBINATION_FILL = PatternFill(
    start_color="F5EBE6",
    end_color="F5EBE6",
    fill_type="solid",
)

REVIEW_FILL = PatternFill(
    start_color="FFF2CC",
    end_color="FFF2CC",
    fill_type="solid",
)


def parse_amount(value):
    """把 Excel 金额安全转换为 Decimal。"""

    if value is None:
        return None

    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, AttributeError):
        return None


def read_sheet(sheet):
    """读取工作表 B 列中的交易金额。"""

    records = []

    for row in range(START_ROW, sheet.max_row + 1):
        amount = parse_amount(
            sheet.cell(row=row, column=AMOUNT_COLUMN).value
        )

        if amount is None:
            continue

        records.append(
            Record(
                row=row,
                amount=amount,
            )
        )

    return records


def load_excel(filename):
    workbook = load_workbook(filename)

    for sheet_name in ("Sheet1", "Sheet2"):
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"找不到工作表 {sheet_name}。")

    return (
        workbook,
        read_sheet(workbook["Sheet1"]),
        read_sheet(workbook["Sheet2"]),
    )


def get_fill(record):
    """人工复核黄色优先于普通匹配颜色。"""

    if record.review_required:
        return REVIEW_FILL

    if record.match_type == "One-to-One":
        return ONE_TO_ONE_FILL

    if record.match_type in {"One-to-Two", "Two-to-One"}:
        return COMBINATION_FILL

    return None


def clear_old_results(sheet):
    """清除上一次运行留下的结果和颜色。"""

    headers = {
        MATCH_TYPE_COLUMN: "Match Type",
        PARTNER_COLUMN: "Partner Rows",
        REVIEW_COLUMN: "Review",
    }

    for column, header in headers.items():
        cell = sheet.cell(row=2, column=column)
        cell.value = header
        cell.font = Font(bold=True)

    for row in range(START_ROW, sheet.max_row + 1):
        for column in (
            AMOUNT_COLUMN,
            MATCH_TYPE_COLUMN,
            PARTNER_COLUMN,
            REVIEW_COLUMN,
        ):
            sheet.cell(row=row, column=column).fill = PatternFill()

        sheet.cell(row=row, column=MATCH_TYPE_COLUMN).value = None
        sheet.cell(row=row, column=PARTNER_COLUMN).value = None
        sheet.cell(row=row, column=REVIEW_COLUMN).value = None


def write_records(sheet, records):
    """把匹配结果、对应行号和复核信息写入 Excel。"""

    for record in records:
        if not record.matched:
            continue

        cells = [
            sheet.cell(row=record.row, column=AMOUNT_COLUMN),
            sheet.cell(row=record.row, column=MATCH_TYPE_COLUMN),
            sheet.cell(row=record.row, column=PARTNER_COLUMN),
            sheet.cell(row=record.row, column=REVIEW_COLUMN),
        ]

        cells[1].value = record.match_type
        cells[2].value = ", ".join(
            str(row) for row in record.partners
        )
        cells[3].value = record.review_reason

        fill = get_fill(record)

        if fill is not None:
            for cell in cells:
                cell.fill = fill


def save_results(
    workbook,
    sheet1_records,
    sheet2_records,
    output_filename,
):
    sheet1 = workbook["Sheet1"]
    sheet2 = workbook["Sheet2"]

    clear_old_results(sheet1)
    clear_old_results(sheet2)

    write_records(sheet1, sheet1_records)
    write_records(sheet2, sheet2_records)

    for sheet in (sheet1, sheet2):
        sheet.column_dimensions["C"].width = 18
        sheet.column_dimensions["D"].width = 20
        sheet.column_dimensions["E"].width = 32

    workbook.save(output_filename)
