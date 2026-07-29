from decimal import Decimal
from openpyxl import load_workbook

from models import Record


START_ROW = 3          # 从第三行开始读取
AMOUNT_COLUMN = 2      # B列


def read_sheet(sheet):
    """读取一个 Sheet，返回 Record 列表"""

    records = []

    row = START_ROW

    while True:

        value = sheet.cell(row=row, column=AMOUNT_COLUMN).value

        # B列为空，说明结束
        if value is None:
            break

        record = Record(
            row=row,
            amount=Decimal(str(value))
        )

        records.append(record)

        row += 1

    return records


def load_excel(filename):

    workbook = load_workbook(filename)

    sheet1 = workbook["Sheet1"]

    sheet2 = workbook["Sheet2"]

    sheet1_records = read_sheet(sheet1)

    sheet2_records = read_sheet(sheet2)

    return sheet1_records, sheet2_records