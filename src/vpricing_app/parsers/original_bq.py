from pathlib import Path

from openpyxl import load_workbook

from vpricing_app.models import OriginalBQ, QuoteLine, parse_money


def _is_package_sheet(name: str) -> bool:
    return name.upper().startswith("PACKAGE")


def _project_name(value: object) -> str:
    text = str(value or "").strip()
    if "PROJECT NAME" in text.upper() and ":" in text:
        return text.split(":", 1)[1].strip()
    return text


def parse_original_bq(path: str | Path) -> OriginalBQ:
    wb = load_workbook(path, data_only=True)
    packages: dict[str, list[QuoteLine]] = {}
    project_name = Path(path).stem

    for ws in wb.worksheets:
        if not _is_package_sheet(ws.title):
            continue
        if ws["A1"].value:
            project_name = _project_name(ws["A1"].value)

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
            if not description:
                continue
            description_text = str(description).strip()
            if description_text.upper() == "GRAND TOTAL":
                break
            if isinstance(item_no, str) and len(item_no.strip()) == 1 and item_no.strip().isalpha():
                current_section = item_no.strip()
                continue
            if item_no is None and not ws.cell(row_number, 3).value and not ws.cell(row_number, 4).value:
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
                    source_row=row_number,
                )
            )
        packages[ws.title] = lines

    return OriginalBQ(project_name=project_name, source_filename=Path(path).name, packages=packages)
