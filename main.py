from excel_io import load_excel

TOLERANCE = 0.01


def main():

    sheet1_records, sheet2_records = load_excel("data.xlsx")

    print("=" * 40)
    print(" Excel Reconciliation Tool")
    print("=" * 40)
    print()

    print(f"Sheet1 Records : {len(sheet1_records)}")
    print(f"Sheet2 Records : {len(sheet2_records)}")
    print()

    print(f"Tolerance : ±{TOLERANCE:.2f}")
    print()

    print("Ready for Matching...")
    print("=" * 40)

    # ===== 临时测试：显示前5条记录 =====
    print("\nSheet1 前5条：")
    for record in sheet1_records[:5]:
        print(record)

    print("\nSheet2 前5条：")
    for record in sheet2_records[:5]:
        print(record)


if __name__ == "__main__":
    main()