from decimal import Decimal

from vpricing_app.models import SourceType
from vpricing_app.parsers.vendor_pdf import parse_vendor_pdf


def test_parse_vendor_pdf_extracts_official_quote(sample_vendor_pdf):
    quote = parse_vendor_pdf(sample_vendor_pdf, vendor_name="Rich")

    assert quote.vendor_name == "Rich"
    assert quote.source_type == SourceType.PDF
    assert quote.is_official is True
    assert quote.package_totals["PACKAGE 1 with Meter Run"].total == Decimal("205")
    lines = quote.packages["PACKAGE 1 with Meter Run"]
    assert [line.description for line in lines] == [
        "Sub Alarm Panel",
        "Manual Call Point",
        "Vendor Added Testing",
    ]
    assert lines[0].confidence == "high"


def test_parse_vendor_pdf_extracts_bq_layout_with_package_heading(sample_vendor_pdf_bq_layout):
    quote = parse_vendor_pdf(sample_vendor_pdf_bq_layout, vendor_name="Rich")

    assert quote.package_totals["Package 1 - Buildings"].total == Decimal("75")
    lines = quote.packages["Package 1 - Buildings"]
    assert len(lines) == 1
    assert lines[0].section == "A"
    assert lines[0].description == "Sub Alarm Panel"
    assert lines[0].labor_rate == Decimal("15")


def test_parse_vendor_pdf_total_uses_last_three_non_empty_values(sample_vendor_pdf_bq_layout, monkeypatch):
    class FakePage:
        def extract_tables(self):
            return [
                [
                    ["Package 1 - Buildings", "", "", "", ""],
                    ["Size:", "DESCRIPTION", "Quantity", "Revised Quantity", "Unit", "Labor Rate", "Total Cost"],
                    ["1", "Sub Alarm Panel", "5", "5", "Nos", "15", "75"],
                    ["GRAND TOTAL", "", "", "", "$ 97,238.10", "$ 590,331.75", "$ 668,117.75", ""],
                ]
            ]

    class FakePdf:
        pages = [FakePage()]

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return None

    monkeypatch.setattr("vpricing_app.parsers.vendor_pdf.pdfplumber.open", lambda _: FakePdf())

    quote = parse_vendor_pdf(sample_vendor_pdf_bq_layout, vendor_name="Rich")

    total = quote.package_totals["Package 1 - Buildings"]
    assert total.material == Decimal("97238.10")
    assert total.labor == Decimal("590331.75")
    assert total.total == Decimal("668117.75")
