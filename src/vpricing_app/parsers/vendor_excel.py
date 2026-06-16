from pathlib import Path

from openpyxl import load_workbook

from vpricing_app.models import Money, QuoteLine, SourceType, VendorQuote, parse_money
from vpricing_app.parsers.original_bq import _is_package_sheet


def parse_vendor_excel(path: str | Path, vendor_name: str) -> VendorQuote:
    wb = load_workbook(path, data_only=True)
    quote = VendorQuote(
        vendor_name=vendor_name,
        source_filename=Path(path).name,
        source_type=SourceType.EXCEL,
        is_official=False,
    )

    for ws in wb.worksheets:
        if not _is_package_sheet(ws.title):
            continue
        header_row = None
        for row_number in range(1, ws.max_row + 1):
            if str(ws.cell(row_number, 2).value or "").strip().upper() == "DESCRIPTION":
                header_row = row_number
                break
        if header_row is None:
            continue

        current_section: str | None = None
        lines: list[QuoteLine] = []
        for row_number in range(header_row + 1, ws.max_row + 1):
            item_no = ws.cell(row_number, 1).value
            description = ws.cell(row_number, 2).value
            if str(item_no or "").strip().upper() == "GRAND TOTAL":
                quote.package_totals[ws.title] = Money(
                    material=parse_money(ws.cell(row_number, 8).value),
                    labor=parse_money(ws.cell(row_number, 9).value),
                    total=parse_money(ws.cell(row_number, 10).value),
                )
                break
            if not description:
                continue
            description_text = str(description).strip()
            if isinstance(item_no, str) and len(item_no.strip()) == 1 and item_no.strip().isalpha():
                current_section = item_no.strip()
                continue
            lines.append(
                QuoteLine(
                    package=ws.title,
                    section=current_section,
                    item_no=str(item_no).strip() if item_no is not None else None,
                    description=description_text,
                    quantity=parse_money(ws.cell(row_number, 3).value),
                    revised_quantity=parse_money(ws.cell(row_number, 4).value),
                    unit=str(ws.cell(row_number, 5).value or "").strip() or None,
                    material_rate=parse_money(ws.cell(row_number, 6).value),
                    labor_rate=parse_money(ws.cell(row_number, 7).value),
                    material_total=parse_money(ws.cell(row_number, 8).value),
                    labor_total=parse_money(ws.cell(row_number, 9).value),
                    total=parse_money(ws.cell(row_number, 10).value),
                    source_row=row_number,
                )
            )
        quote.packages[ws.title] = lines

    return quote
