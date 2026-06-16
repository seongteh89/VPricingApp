from decimal import Decimal

from vpricing_app.models import Money, OriginalBQ, QuoteLine, SourceType, VendorQuote
from vpricing_app.services.comparison_service import build_new_comparison
from vpricing_app.services.matcher import build_comparison_rows
from vpricing_app.services.validator import validate_comparison


def _line(description: str, item_no: str, revised_quantity: str = "1") -> QuoteLine:
    return QuoteLine(
        package="Package 1",
        section="A",
        item_no=item_no,
        description=description,
        quantity=Decimal("1"),
        revised_quantity=Decimal(revised_quantity),
        unit="Lot",
        total=Decimal("10"),
    )


def test_matcher_marks_vendor_added_items():
    original = OriginalBQ(
        project_name="Sample",
        source_filename="original.xlsx",
        packages={"Package 1": [_line("Original Item", "1")]},
    )
    vendor = VendorQuote(
        vendor_name="Rich",
        source_filename="rich.pdf",
        source_type=SourceType.PDF,
        packages={"Package 1": [_line("Original Item", "1"), _line("Extra Item", "2")]},
    )

    rows = build_comparison_rows(original, [vendor])

    assert len(rows["Package 1"]) == 2
    assert rows["Package 1"][1].original_line is None
    assert rows["Package 1"][1].vendor_lines["Rich"].is_vendor_added is True


def test_validator_flags_quantity_change_and_total_mismatch():
    original = OriginalBQ(
        project_name="Sample",
        source_filename="original.xlsx",
        packages={"Package 1": [_line("Original Item", "1", revised_quantity="1")]},
    )
    changed = _line("Original Item", "1", revised_quantity="2")
    vendor = VendorQuote(
        vendor_name="Rensar",
        source_filename="rensar.xlsx",
        source_type=SourceType.EXCEL,
        packages={"Package 1": [changed]},
        package_totals={"Package 1": Money(material=Decimal("100"), labor=Decimal("50"), total=Decimal("140"))},
    )
    rows = build_comparison_rows(original, [vendor])

    warnings = validate_comparison(original, [vendor], rows)

    codes = {warning.code for warning in warnings}
    assert "protected_field_changed" in codes
    assert "package_total_mismatch" in codes


def test_build_new_comparison_maps_vendor_package_heading_to_original_sheet():
    original = OriginalBQ(
        project_name="Sample",
        source_filename="original.xlsx",
        packages={"PACKAGE 1 with Meter Run": [_line("Original Item", "1")]},
    )
    vendor_line = _line("Original Item", "1")
    vendor_line.package = "Package 1 - Buildings"
    vendor = VendorQuote(
        vendor_name="Rich",
        source_filename="rich.pdf",
        source_type=SourceType.PDF,
        packages={"Package 1 - Buildings": [vendor_line]},
    )

    comparison = build_new_comparison(original, [vendor])

    assert "PACKAGE 1 with Meter Run" in comparison.vendors[0].packages
    assert comparison.vendors[0].packages["PACKAGE 1 with Meter Run"][0].description == "Original Item"


def test_matcher_does_not_match_different_descriptions_by_item_number_only():
    original = OriginalBQ(
        project_name="Sample",
        source_filename="original.xlsx",
        packages={"Package 1": [_line("Original Item", "3")]},
    )
    vendor = VendorQuote(
        vendor_name="E-Tech",
        source_filename="etech.xlsx",
        source_type=SourceType.EXCEL,
        packages={"Package 1": [_line("Different Vendor Item", "3")]},
    )

    rows = build_comparison_rows(original, [vendor])

    assert len(rows["Package 1"]) == 2
    assert rows["Package 1"][0].original_line.description == "Original Item"
    assert "E-Tech" not in rows["Package 1"][0].vendor_lines
    assert rows["Package 1"][1].vendor_lines["E-Tech"].is_vendor_added is True
