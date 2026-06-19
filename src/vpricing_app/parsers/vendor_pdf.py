from pathlib import Path

import pdfplumber

from vpricing_app.models import Money, QuoteLine, SourceType, VendorQuote, parse_money


HEADER_ALIASES = {
    "package": "package",
    "section": "section",
    "item": "item_no",
    "size": "item_no",
    "size:": "item_no",
    "description": "description",
    "qty": "quantity",
    "quantity": "quantity",
    "rev qty": "revised_quantity",
    "revised quantity": "revised_quantity",
    "unit": "unit",
    "mat rate": "material_rate",
    "material rate": "material_rate",
    "lab rate": "labor_rate",
    "labor rate": "labor_rate",
    "mat total": "material_total",
    "material total": "material_total",
    "total material cost": "material_total",
    "lab total": "labor_total",
    "labor total": "labor_total",
    "total labor cost": "labor_total",
    "total": "total",
    "total cost": "total",
}


def _clean_cell(value: object) -> str:
    return " ".join(str(value or "").replace("\n", " ").split())


def _header_map(row: list[object]) -> dict[int, str]:
    result: dict[int, str] = {}
    for index, value in enumerate(row):
        key = _clean_cell(value).lower()
        if key in HEADER_ALIASES:
            result[index] = HEADER_ALIASES[key]
    return result


def parse_vendor_pdf(path: str | Path, vendor_name: str) -> VendorQuote:
    quote = VendorQuote(
        vendor_name=vendor_name,
        source_filename=Path(path).name,
        source_type=SourceType.PDF,
        is_official=True,
    )
    current_header: dict[int, str] = {}
    current_package: str | None = None
    current_section: str | None = None
    current_section_description: str | None = None

    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for row in table:
                    cells = [_clean_cell(cell) for cell in row]
                    if not any(cells):
                        continue
                    possible_header = _header_map(cells)
                    if "description" in possible_header.values() and "total" in possible_header.values():
                        current_header = possible_header
                        continue
                    if cells[0].lower().startswith("package ") and (len(cells) == 1 or not cells[1]):
                        current_package = cells[0]
                        continue
                    if cells[0].upper() == "GRAND TOTAL" and current_package:
                        values = [cell for cell in cells[1:] if cell]
                        material, labor, total = (values[-3:] if len(values) >= 3 else ["", "", ""])
                        quote.package_totals[current_package] = Money(
                            material=parse_money(material),
                            labor=parse_money(labor),
                            total=parse_money(total),
                        )
                        continue
                    if not current_header:
                        continue

                    mapped = {name: cells[index] for index, name in current_header.items() if index < len(cells)}
                    description = mapped.get("description", "")
                    package = mapped.get("package") or current_package
                    if not package or not description:
                        continue
                    item_no = mapped.get("item_no") or None
                    if item_no and len(item_no.strip()) == 1 and item_no.strip().isalpha():
                        current_package = package
                        current_section = item_no.strip()
                        current_section_description = description
                        continue
                    current_package = package
                    line = QuoteLine(
                        package=package,
                        section=mapped.get("section") or current_section,
                        section_description=current_section_description,
                        item_no=item_no,
                        description=description,
                        quantity=parse_money(mapped.get("quantity")),
                        revised_quantity=parse_money(mapped.get("revised_quantity")),
                        unit=mapped.get("unit") or None,
                        material_rate=parse_money(mapped.get("material_rate")),
                        labor_rate=parse_money(mapped.get("labor_rate")),
                        material_total=parse_money(mapped.get("material_total")),
                        labor_total=parse_money(mapped.get("labor_total")),
                        total=parse_money(mapped.get("total")),
                        confidence="high" if mapped.get("total") else "medium",
                    )
                    quote.packages.setdefault(package, []).append(line)

    return quote
