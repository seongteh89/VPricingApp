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


def test_parse_vendor_excel_accepts_request_for_quotation_sheet(sample_vendor_rfq):
    quote = parse_vendor_excel(sample_vendor_rfq, vendor_name="T-Tech")

    assert quote.package_totals["Request For Quotation"].total == Decimal("37090.80")
    lines = quote.packages["Request For Quotation"]
    assert len(lines) == 2
    assert lines[0].description == "PBB : Automatic Fire Sprinkler System"
    assert lines[0].quantity == Decimal("1")
    assert lines[0].revised_quantity == Decimal("12")
    assert lines[0].material_rate == Decimal("600")
    assert lines[0].material_total == Decimal("7200")
    assert lines[0].labor_total == Decimal("15054.48")
    assert lines[0].total == Decimal("22254.48")
