from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any


@dataclass
class Record:
    """Represents one transaction record."""

    row: int
    amount: Decimal

    # 原始工作表名称，例如 Sheet1 / Sheet2
    source_sheet: str = ""

    # 保存原始 Excel 中的辅助信息
    # 例如流水号、摘要、日期、账户名、Key word 等
    extra: dict[str, Any] = field(default_factory=dict)

    matched: bool = False
    match_type: str = ""

    partners: list[int] = field(default_factory=list)

    review_required: bool = False
    review_reason: str = ""

    # 一对一匹配的关键词核验结果
    keyword_match: bool = False
    keyword_conflict: bool = False
