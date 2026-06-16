from decimal import Decimal

from vpricing_app.models import (
    Money,
    QuoteLine,
    SourceType,
    VendorQuote,
    normalize_text,
    parse_money,
)


def test_parse_money_accepts_currency_strings():
    assert parse_money("$ 1,234.50") == Decimal("1234.50")
    assert parse_money("-") == Decimal("0")
    assert parse_money(None) == Decimal("0")


def test_normalize_text_collapses_spacing_and_case():
    assert normalize_text("  Fire   Alarm\nPanel ") == "fire alarm panel"


def test_quote_line_key_uses_package_section_and_description():
    line = QuoteLine(
        package="PACKAGE 1",
        section="A",
        item_no="1",
        description="Sub Alarm Panel",
        quantity=Decimal("5"),
        revised_quantity=Decimal("5"),
        unit="Nos",
    )

    assert line.match_key == "package 1|a|1|sub alarm panel"


def test_vendor_quote_total_sums_packages():
    quote = VendorQuote(
        vendor_name="Rich",
        source_filename="rich.pdf",
        source_type=SourceType.PDF,
        package_totals={"Package 1": Money(total=Decimal("100")), "Package 2": Money(total=Decimal("50"))},
    )

    assert quote.grand_total == Decimal("150")
