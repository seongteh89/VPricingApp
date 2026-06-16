from decimal import Decimal

from vpricing_app.models import SourceType
from vpricing_app.parsers.vendor_excel import parse_vendor_excel


def test_parse_vendor_excel_extracts_rates_totals_and_added_rows(sample_vendor_excel):
    quote = parse_vendor_excel(sample_vendor_excel, vendor_name="Rensar")

    assert quote.vendor_name == "Rensar"
    assert quote.source_type == SourceType.EXCEL
    assert quote.package_totals["PACKAGE 1 with Meter Run"].total == Decimal("320")
    lines = quote.packages["PACKAGE 1 with Meter Run"]
    assert len(lines) == 3
    assert lines[1].description == "Manual Call Point"
    assert lines[1].revised_quantity == Decimal("12")
    assert lines[1].labor_rate == Decimal("10")
    assert lines[2].description == "Vendor Extra Item"
