from dataclasses import dataclass, field
from decimal import Decimal

@dataclass
class Record:
    """Represents one transaction record."""

    row: int
    amount: Decimal

    matched: bool = False
    match_type: str = ""

    partners: list[int] = field(default_factory=list)