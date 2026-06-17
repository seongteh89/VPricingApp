from decimal import Decimal, InvalidOperation
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class SourceType(str, Enum):
    EXCEL = "excel"
    PDF = "pdf"


class WarningLevel(str, Enum):
    INFO = "info"
    REVIEW = "review"
    BLOCKER = "blocker"


def normalize_text(value: Any) -> str:
    return " ".join(str(value or "").strip().lower().split())


def parse_money(value: Any) -> Decimal:
    if value is None:
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int | float):
        return Decimal(str(value))
    text = str(value).replace("$", "").replace(",", "").strip()
    if text in {"", "-", "-"}:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        cleaned = "".join(ch for ch in text if ch.isdigit() or ch in ".-")
        return Decimal(cleaned or "0")


class Money(BaseModel):
    material: Decimal = Decimal("0")
    labor: Decimal = Decimal("0")
    total: Decimal = Decimal("0")

    @property
    def calculated_total(self) -> Decimal:
        return self.material + self.labor


class QuoteLine(BaseModel):
    package: str
    section: str | None = None
    section_description: str | None = None
    item_no: str | None = None
    description: str
    quantity: Decimal | None = None
    revised_quantity: Decimal | None = None
    unit: str | None = None
    material_rate: Decimal | None = None
    labor_rate: Decimal | None = None
    material_total: Decimal = Decimal("0")
    labor_total: Decimal = Decimal("0")
    total: Decimal = Decimal("0")
    source_row: int | None = None
    is_vendor_added: bool = False
    confidence: str = "high"
    notes: str | None = None

    @property
    def match_key(self) -> str:
        return "|".join(
            [
                normalize_text(self.package),
                normalize_text(self.section),
                normalize_text(self.item_no),
                normalize_text(self.description),
            ]
        )


class OriginalBQ(BaseModel):
    project_name: str
    source_filename: str
    packages: dict[str, list[QuoteLine]]


class VendorQuote(BaseModel):
    vendor_name: str
    source_filename: str
    source_type: SourceType
    revision: str | None = None
    is_official: bool = False
    packages: dict[str, list[QuoteLine]] = Field(default_factory=dict)
    package_totals: dict[str, Money] = Field(default_factory=dict)

    @property
    def grand_total(self) -> Decimal:
        return sum((money.total for money in self.package_totals.values()), Decimal("0"))


class ImportWarning(BaseModel):
    level: WarningLevel
    code: str
    vendor_name: str | None = None
    package: str | None = None
    row_label: str | None = None
    message: str


class ComparisonModel(BaseModel):
    project_name: str
    original: OriginalBQ
    vendors: list[VendorQuote]
    warnings: list[ImportWarning] = Field(default_factory=list)
