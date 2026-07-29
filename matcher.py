from decimal import Decimal


TOLERANCE = Decimal("0.01")


def amount_equal(a, b):
    """金额比较（允许±0.01误差）"""
    return abs(a - b) <= TOLERANCE


class Matcher:

    def __init__(self, sheet1_records, sheet2_records):
        self.sheet1 = sheet1_records
        self.sheet2 = sheet2_records

    def run(self):
        """运行所有匹配规则（目前只有一对一）"""
        self.one_to_one_match()

    def one_to_one_match(self):

        matched_count = 0

        for left in self.sheet1:

            if left.matched:
                continue

            for right in self.sheet2:

                if right.matched:
                    continue

                if amount_equal(left.amount, right.amount):

                    left.matched = True
                    right.matched = True

                    left.match_type = "One-to-One"
                    right.match_type = "One-to-One"

                    left.partners.append(right.row)
                    right.partners.append(left.row)

                    matched_count += 1

                    break

        print()
        print("========== One-to-One Match ==========")
        print(f"Matched : {matched_count}")
        print(f"Remaining Sheet1 : {sum(not r.matched for r in self.sheet1)}")
        print(f"Remaining Sheet2 : {sum(not r.matched for r in self.sheet2)}")
        print("======================================")