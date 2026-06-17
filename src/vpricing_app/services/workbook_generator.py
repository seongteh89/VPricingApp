import json
import re
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from vpricing_app.models import ComparisonModel, Money, QuoteLine, VendorQuote
from vpricing_app.parsers.original_bq import (
    _all_header_columns,
    _find_header_row,
    _first_header_column,
    _header_columns,
)
from vpricing_app.services.filename_builder import build_output_filename
from vpricing_app.services.matcher import ComparisonRow, build_comparison_rows


HEADER_FILL = PatternFill("solid", fgColor="FFFFFFFF")
HEADER_FONT = Font(color="000000", bold=True)
ADDED_FILL = PatternFill("solid", fgColor="FFFFF2CC")
LOW_FILL = PatternFill("solid", fgColor="FFC6EFCE")
REVISED_QTY_FILL = PatternFill("solid", fgColor="FF7CC7E3")
MATERIAL_FILL = PatternFill("solid", fgColor="FF808080")
LABOR_FILL = PatternFill("solid", fgColor="FFFFFF00")
TOTAL_FILL = PatternFill("solid", fgColor="FFD9D9D9")
THIN_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
MONEY_FORMAT = r'_-"$"* #,##0.00_-;\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'
METADATA_MARKER = "comparison_model_json_v1"
METADATA_CHUNK_SIZE = 30000


def _as_float(value: Decimal | int | float | None) -> float:
    return float(value or 0)


def _summary_package_label(package_name: str) -> str:
    match = re.search(r"package\s*(\d+)", package_name, flags=re.IGNORECASE)
    return f"package {match.group(1)}" if match else package_name


def _package_total(vendor: VendorQuote, package_name: str) -> Money:
    if package_name in vendor.package_totals:
        return vendor.package_totals[package_name]
    lines = vendor.packages.get(package_name, [])
    return Money(
        material=sum((line.material_total for line in lines), Decimal("0")),
        labor=sum((line.labor_total for line in lines), Decimal("0")),
        total=sum((line.total for line in lines), Decimal("0")),
    )


def _write_headers(ws, vendors: list[str]) -> None:
    base_headers = ["Item", "Description", "Quantity", "Revised Quantity", "Unit"]
    vendor_headers = [
        "Revised Quantity",
        "Unit",
        "Material Rate",
        "Labor Rate",
        "Material Total",
        "Labor Total",
        "Total Cost",
        "Review Flag",
    ]
    headers = base_headers[:]
    for vendor in vendors:
        headers.extend([f"{vendor} {header}" for header in vendor_headers])
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def _vendor_block_start(vendor_index: int) -> int:
    return 6 + vendor_index * 8


def _style_money_cell(cell, fill: PatternFill | None = None) -> None:
    cell.number_format = MONEY_FORMAT
    cell.border = THIN_BORDER
    if fill:
        cell.fill = fill


def _style_data_row(ws, row_number: int, vendor_count: int) -> None:
    for cell in ws[row_number]:
        cell.border = THIN_BORDER
    ws.cell(row_number, 3).fill = REVISED_QTY_FILL
    ws.cell(row_number, 4).fill = REVISED_QTY_FILL
    for vendor_index in range(vendor_count):
        start_col = _vendor_block_start(vendor_index)
        ws.cell(row_number, start_col).fill = REVISED_QTY_FILL
        ws.cell(row_number, start_col + 2).fill = MATERIAL_FILL
        ws.cell(row_number, start_col + 3).fill = LABOR_FILL
        ws.cell(row_number, start_col + 4).fill = MATERIAL_FILL
        for col in range(start_col + 2, start_col + 7):
            _style_money_cell(ws.cell(row_number, col))


def _reference_line(row: ComparisonRow) -> QuoteLine | None:
    if row.original_line:
        return row.original_line
    return next(iter(row.vendor_lines.values()), None)


def _section_key(row: ComparisonRow) -> tuple[str, str] | None:
    line = _reference_line(row)
    if not line or not line.section:
        return None
    return line.section, line.section_description or ""


def _write_section_subtotal_row(ws, section_key: tuple[str, str], vendor_count: int) -> int:
    section, description = section_key
    values = [section, description, "", "", ""]
    for _ in range(vendor_count):
        values.extend(["", "", "", "", "", "", "", ""])
    ws.append(values)
    row_number = ws.max_row
    for cell in ws[row_number]:
        cell.font = Font(bold=True)
        cell.fill = TOTAL_FILL
        cell.border = THIN_BORDER
    return row_number


def _write_section_subtotal_formulas(
    ws,
    subtotal_row: int | None,
    first_item_row: int | None,
    last_item_row: int | None,
    vendor_count: int,
) -> None:
    if subtotal_row is None or first_item_row is None or last_item_row is None:
        return
    for vendor_index in range(vendor_count):
        start_col = _vendor_block_start(vendor_index)
        for col in (start_col + 4, start_col + 5, start_col + 6):
            column_letter = get_column_letter(col)
            cell = ws.cell(subtotal_row, col)
            cell.value = f"=SUM({column_letter}{first_item_row}:{column_letter}{last_item_row})"
            cell.font = Font(bold=True)
            _style_money_cell(cell, TOTAL_FILL)


def _write_summary(summary, model: ComparisonModel) -> None:
    summary.append(["Vendor", "Grand Total"])
    for cell in summary[1]:
        cell.font = HEADER_FONT
        cell.border = THIN_BORDER
    for vendor in model.vendors:
        summary.append([vendor.vendor_name, _as_float(vendor.grand_total)])
        summary.cell(summary.max_row, 2).number_format = MONEY_FORMAT

    block_start = len(model.vendors) + 3
    summary.cell(block_start, 2).value = "show Total Amount of each Vendor"
    header_row = block_start + 1
    for index, vendor in enumerate(model.vendors, start=3):
        cell = summary.cell(header_row, index)
        cell.value = vendor.vendor_name
        cell.font = HEADER_FONT
        cell.border = THIN_BORDER

    package_start = header_row + 1
    for row_offset, package_name in enumerate(model.original.packages):
        row_number = package_start + row_offset
        summary.cell(row_number, 2).value = _summary_package_label(package_name)
        for col_offset, vendor in enumerate(model.vendors, start=3):
            cell = summary.cell(row_number, col_offset)
            cell.value = _as_float(_package_total(vendor, package_name).total)
            cell.number_format = MONEY_FORMAT

    total_row = package_start + len(model.original.packages)
    summary.cell(total_row, 2).value = "Total Package Amount"
    summary.cell(total_row, 2).font = Font(bold=True)
    for col_offset, vendor in enumerate(model.vendors, start=3):
        total = sum((_package_total(vendor, package_name).total for package_name in model.original.packages), Decimal("0"))
        cell = summary.cell(total_row, col_offset)
        cell.value = _as_float(total)
        cell.font = Font(bold=True)
        cell.number_format = MONEY_FORMAT
        cell.fill = TOTAL_FILL


def _write_package_total_row(ws, package_name: str, vendors: list[VendorQuote]) -> None:
    values = ["GRAND TOTAL", "", "", "", ""]
    for vendor in vendors:
        total = _package_total(vendor, package_name)
        values.extend(["", "", "", "", _as_float(total.material), _as_float(total.labor), _as_float(total.total), ""])
    ws.append(values)
    row_number = ws.max_row
    for cell in ws[row_number]:
        cell.font = Font(bold=True)
        cell.fill = TOTAL_FILL
        cell.border = THIN_BORDER
    for vendor_index in range(len(vendors)):
        start_col = _vendor_block_start(vendor_index)
        for col in (start_col + 4, start_col + 5, start_col + 6):
            _style_money_cell(ws.cell(row_number, col), TOTAL_FILL)


def _autosize(ws) -> None:
    for column in ws.columns:
        width = max(len(str(cell.value or "")) for cell in column)
        ws.column_dimensions[get_column_letter(column[0].column)].width = min(max(width + 2, 10), 45)


def _safe_sheet_title(base_title: str, existing_titles: set[str]) -> str:
    cleaned = re.sub(r"[\[\]\*\?/\\:]", " ", base_title).strip() or "Sheet"
    title = cleaned[:31]
    counter = 1
    while title in existing_titles:
        suffix = f" {counter}"
        title = f"{cleaned[: 31 - len(suffix)]}{suffix}"
        counter += 1
    existing_titles.add(title)
    return title


def _vendor_template_title(package_name: str, vendor_name: str, existing_titles: set[str]) -> str:
    package_number = _package_number(package_name)
    if package_number is not None:
        return _safe_sheet_title(f"Package {package_number} - {vendor_name}", existing_titles)
    preferred = f"{package_name} - {vendor_name}"
    if len(preferred) <= 31:
        return _safe_sheet_title(preferred, existing_titles)
    abbreviated_package = re.sub(r"\bRequest\s+For\s+Quotation\b", "RFQ", package_name, flags=re.IGNORECASE)
    abbreviated_package = re.sub(r"\bQuotation\b", "Quote", abbreviated_package, flags=re.IGNORECASE)
    preferred = f"{abbreviated_package} - {vendor_name}"
    return _safe_sheet_title(preferred, existing_titles)


def _package_number(name: str) -> str | None:
    match = re.search(r"package\s*(\d+)", name, flags=re.IGNORECASE)
    return match.group(1) if match else None


def _worksheet_for_package(wb, package_name: str):
    if package_name in wb.sheetnames:
        return wb[package_name]
    package_number = _package_number(package_name)
    if package_number is not None:
        for ws in wb.worksheets:
            if _package_number(ws.title) == package_number:
                return ws
    return None


def _worksheet_for_vendor_package(wb, package_name: str, vendor_name: str):
    exact_titles = [f"{package_name} - {vendor_name}"]
    package_number = _package_number(package_name)
    if package_number is not None:
        exact_titles.append(f"Package {package_number} - {vendor_name}")
    for title in exact_titles:
        if title in wb.sheetnames:
            return wb[title]

    normalized_vendor = vendor_name.casefold()
    for ws in wb.worksheets:
        if normalized_vendor not in ws.title.casefold():
            continue
        if package_number is None or _package_number(ws.title) == package_number:
            return ws
    return None


def _pricing_columns(ws) -> list[int]:
    header_row = _find_header_row(ws)
    if header_row is None:
        return []
    headers = _header_columns(ws, header_row)
    pricing_candidates = [
        _first_header_column(headers, "revised quantity"),
        _first_header_column(headers, "unit"),
        _first_header_column(headers, "material rate"),
        _first_header_column(headers, "labor rate"),
        _first_header_column(headers, "total material cost", "material total"),
        _first_header_column(headers, "total labor cost", "labor total"),
        _first_header_column(headers, "total cost", "total"),
        *_all_header_columns(headers, "number of servicing"),
        *_all_header_columns(headers, "unit price"),
        *_all_header_columns(headers, "amount"),
    ]
    columns = [column for column in pricing_candidates if column is not None]
    if not columns:
        return []
    return list(range(min(columns), max(columns) + 1))


def _pricing_column_map(ws) -> dict[str, int | list[int] | None]:
    header_row = _find_header_row(ws)
    if header_row is None:
        return {}
    headers = _header_columns(ws, header_row)
    return {
        "revised_quantity": _first_header_column(headers, "revised quantity", "number of servicing"),
        "unit": _first_header_column(headers, "unit"),
        "material_rate": _first_header_column(headers, "material rate"),
        "labor_rate": _first_header_column(headers, "labor rate"),
        "material_total": _first_header_column(headers, "total material cost", "material total"),
        "labor_total": _first_header_column(headers, "total labor cost", "labor total"),
        "total": _first_header_column(headers, "total cost", "total"),
        "service_cols": _all_header_columns(headers, "number of servicing"),
        "unit_price_cols": _all_header_columns(headers, "unit price"),
        "amount_cols": _all_header_columns(headers, "amount"),
    }


def _write_cell_if_column(ws, row_number: int, column: int | None, value) -> None:
    if column is not None and value is not None:
        ws.cell(row_number, column).value = _as_float(value) if isinstance(value, Decimal) else value


def _write_parsed_pricing(target_ws, target_row: int, line: QuoteLine, pricing_map: dict[str, int | list[int] | None]) -> None:
    service_cols = pricing_map.get("service_cols") or []
    unit_price_cols = pricing_map.get("unit_price_cols") or []
    amount_cols = pricing_map.get("amount_cols") or []
    if isinstance(service_cols, list) and isinstance(unit_price_cols, list) and isinstance(amount_cols, list) and amount_cols:
        _write_cell_if_column(target_ws, target_row, service_cols[0] if service_cols else None, line.revised_quantity)
        _write_cell_if_column(target_ws, target_row, unit_price_cols[0] if unit_price_cols else None, line.material_rate)
        _write_cell_if_column(target_ws, target_row, amount_cols[0], line.material_total)
        if len(unit_price_cols) > 1:
            _write_cell_if_column(target_ws, target_row, unit_price_cols[1], line.labor_rate)
        if len(amount_cols) > 1:
            _write_cell_if_column(target_ws, target_row, amount_cols[1], line.labor_total)
        if len(amount_cols) > 2:
            _write_cell_if_column(target_ws, target_row, amount_cols[-1], line.total)
        return

    _write_cell_if_column(target_ws, target_row, pricing_map.get("revised_quantity"), line.revised_quantity)
    _write_cell_if_column(target_ws, target_row, pricing_map.get("unit"), line.unit)
    _write_cell_if_column(target_ws, target_row, pricing_map.get("material_rate"), line.material_rate)
    _write_cell_if_column(target_ws, target_row, pricing_map.get("labor_rate"), line.labor_rate)
    _write_cell_if_column(target_ws, target_row, pricing_map.get("material_total"), line.material_total)
    _write_cell_if_column(target_ws, target_row, pricing_map.get("labor_total"), line.labor_total)
    _write_cell_if_column(target_ws, target_row, pricing_map.get("total"), line.total)


def _copy_pricing_cell(source_ws, source_row: int, target_ws, target_row: int, column: int) -> None:
    source = source_ws.cell(source_row, column)
    target = target_ws.cell(target_row, column)
    if isinstance(source.value, str) and source.value.startswith("="):
        try:
            target.value = Translator(source.value, origin=source.coordinate).translate_formula(target.coordinate)
        except Exception:
            target.value = source.value
    else:
        target.value = source.value


def _copy_vendor_pricing(source_ws, target_ws, rows: list[ComparisonRow], vendor: VendorQuote, pricing_columns: list[int]) -> None:
    pricing_map = _pricing_column_map(target_ws)
    for row in rows:
        vendor_line = row.vendor_lines.get(vendor.vendor_name)
        if vendor_line is None:
            continue
        target_row = row.original_line.source_row if row.original_line and row.original_line.source_row else vendor_line.source_row
        if target_row is None:
            continue
        if source_ws is not None and pricing_columns:
            if vendor_line.source_row is None:
                _write_parsed_pricing(target_ws, target_row, vendor_line, pricing_map)
                continue
            for column in pricing_columns:
                _copy_pricing_cell(source_ws, vendor_line.source_row, target_ws, target_row, column)
        else:
            _write_parsed_pricing(target_ws, target_row, vendor_line, pricing_map)


def _replace_sheet(wb, title: str):
    if title in wb.sheetnames:
        del wb[title]
    return wb.create_sheet(title)


def _write_change_log_sheet(wb, model: ComparisonModel) -> None:
    change_log = _replace_sheet(wb, "Change Log")
    change_log.append(["Level", "Code", "Vendor", "Package", "Row", "Message"])
    for warning in model.warnings:
        change_log.append(
            [warning.level.value, warning.code, warning.vendor_name, warning.package, warning.row_label, warning.message]
        )
    _autosize(change_log)


def _write_import_warnings_sheet(wb, model: ComparisonModel) -> None:
    warnings = _replace_sheet(wb, "Import Warnings")
    warnings.append(["Level", "Code", "Vendor", "Package", "Message"])
    for warning in model.warnings:
        warnings.append([warning.level.value, warning.code, warning.vendor_name, warning.package, warning.message])
    _autosize(warnings)


def _write_metadata_sheet(wb, model: ComparisonModel) -> None:
    metadata = _replace_sheet(wb, "System Metadata")
    metadata.sheet_state = "hidden"
    metadata["A1"] = METADATA_MARKER
    raw_metadata = model.model_dump_json()
    for index in range(0, len(raw_metadata), METADATA_CHUNK_SIZE):
        row = index // METADATA_CHUNK_SIZE + 2
        metadata.cell(row=row, column=1).value = raw_metadata[index : index + METADATA_CHUNK_SIZE]


def _write_support_sheets(wb, model: ComparisonModel) -> None:
    _write_change_log_sheet(wb, model)
    _write_import_warnings_sheet(wb, model)
    _write_metadata_sheet(wb, model)


def _export_template_preserved_workbook(
    model: ComparisonModel,
    output_path: Path,
    original_template_path: str | Path,
) -> Path:
    wb = load_workbook(original_template_path)
    existing_titles = set(wb.sheetnames)
    rows_by_package = build_comparison_rows(model.original, model.vendors)

    for package_name, rows in rows_by_package.items():
        template_ws = _worksheet_for_package(wb, package_name)
        if template_ws is None:
            continue
        pricing_columns = _pricing_columns(template_ws)
        for vendor in model.vendors:
            vendor_ws = wb.copy_worksheet(template_ws)
            vendor_ws.title = _vendor_template_title(package_name, vendor.vendor_name, existing_titles)
            source_ws = None
            if vendor.source_path:
                vendor_wb = load_workbook(vendor.source_path, data_only=False)
                source_ws = _worksheet_for_package(vendor_wb, package_name)
            _copy_vendor_pricing(source_ws, vendor_ws, rows, vendor, pricing_columns)

    for package_name in rows_by_package:
        if package_name in wb.sheetnames:
            del wb[package_name]

    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    summary = wb.create_sheet("Summary", 0)
    _write_summary(summary, model)
    _autosize(summary)
    _write_support_sheets(wb, model)
    wb.save(output_path)
    return output_path


def export_updated_template_workbook(
    model: ComparisonModel,
    existing_comparison_path: str | Path,
    revised_vendor: VendorQuote,
    output_dir: str | Path,
    revision: str | None = None,
) -> Path:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = build_output_filename(model.project_name, [vendor.vendor_name for vendor in model.vendors], revision)
    output_path = output_dir / filename

    wb = load_workbook(existing_comparison_path)
    rows_by_package = build_comparison_rows(model.original, model.vendors)
    source_wb = load_workbook(revised_vendor.source_path, data_only=False) if revised_vendor.source_path else None
    updated_any_sheet = False

    for package_name, rows in rows_by_package.items():
        target_ws = _worksheet_for_vendor_package(wb, package_name, revised_vendor.vendor_name)
        if target_ws is None:
            continue
        source_ws = _worksheet_for_package(source_wb, package_name) if source_wb is not None else None
        _copy_vendor_pricing(source_ws, target_ws, rows, revised_vendor, _pricing_columns(target_ws))
        updated_any_sheet = True

    if not updated_any_sheet:
        return export_comparison_workbook(model, output_dir, revision=revision)

    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    summary = wb.create_sheet("Summary", 0)
    _write_summary(summary, model)
    _autosize(summary)
    _write_support_sheets(wb, model)
    wb.save(output_path)
    return output_path


def export_comparison_workbook(
    model: ComparisonModel,
    output_dir: str | Path,
    revision: str | None = None,
    original_template_path: str | Path | None = None,
) -> Path:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = build_output_filename(model.project_name, [vendor.vendor_name for vendor in model.vendors], revision)
    output_path = output_dir / filename

    if original_template_path is not None:
        return _export_template_preserved_workbook(model, output_path, original_template_path)

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    _write_summary(summary, model)
    _autosize(summary)

    rows_by_package = build_comparison_rows(model.original, model.vendors)
    for package_name, rows in rows_by_package.items():
        ws = wb.create_sheet(package_name[:31])
        _write_headers(ws, [vendor.vendor_name for vendor in model.vendors])
        current_section_key: tuple[str, str] | None = None
        subtotal_row: int | None = None
        first_item_row: int | None = None
        last_item_row: int | None = None
        for row in rows:
            section_key = _section_key(row)
            if section_key != current_section_key:
                _write_section_subtotal_formulas(
                    ws, subtotal_row, first_item_row, last_item_row, len(model.vendors)
                )
                current_section_key = section_key
                subtotal_row = None
                first_item_row = None
                last_item_row = None
                if section_key is not None:
                    subtotal_row = _write_section_subtotal_row(ws, section_key, len(model.vendors))

            original = row.original_line
            first_vendor_line = next(iter(row.vendor_lines.values()), None)
            values = [
                original.item_no if original else "",
                original.description if original else first_vendor_line.description,
                float(original.quantity) if original and original.quantity is not None else "",
                float(original.revised_quantity) if original and original.revised_quantity is not None else "",
                original.unit if original else "",
            ]
            vendor_totals = []
            for vendor in model.vendors:
                vendor_line = row.vendor_lines.get(vendor.vendor_name)
                if vendor_line is None:
                    values.extend(["N/A", "N/A", "", "", "", "", "", "Missing"])
                    vendor_totals.append(None)
                    continue
                values.extend(
                    [
                        float(vendor_line.revised_quantity) if vendor_line.revised_quantity is not None else "",
                        vendor_line.unit or "",
                        float(vendor_line.material_rate) if vendor_line.material_rate is not None else "",
                        float(vendor_line.labor_rate) if vendor_line.labor_rate is not None else "",
                        float(vendor_line.material_total),
                        float(vendor_line.labor_total),
                        float(vendor_line.total),
                        "Vendor Added Item - Review Required" if vendor_line.is_vendor_added else "",
                    ]
                )
                vendor_totals.append(vendor_line.total)
            ws.append(values)
            _style_data_row(ws, ws.max_row, len(model.vendors))
            if original is None:
                for cell in ws[ws.max_row]:
                    cell.fill = ADDED_FILL
            numeric_totals = [total for total in vendor_totals if total is not None]
            if numeric_totals:
                lowest = min(numeric_totals)
                for index, total in enumerate(vendor_totals):
                    if total == lowest:
                        total_col = 5 + index * 8 + 7
                        ws.cell(ws.max_row, total_col).fill = LOW_FILL
            if current_section_key is not None:
                first_item_row = first_item_row or ws.max_row
                last_item_row = ws.max_row
        _write_section_subtotal_formulas(ws, subtotal_row, first_item_row, last_item_row, len(model.vendors))
        _write_package_total_row(ws, package_name, model.vendors)
        _autosize(ws)

    _write_support_sheets(wb, model)

    wb.save(output_path)
    return output_path


def load_comparison_metadata(path: str | Path) -> ComparisonModel:
    wb = load_workbook(path, data_only=True)
    sheet = wb["System Metadata"]
    if sheet["A1"].value == METADATA_MARKER:
        chunks = []
        row = 2
        while sheet.cell(row=row, column=1).value:
            chunks.append(sheet.cell(row=row, column=1).value)
            row += 1
        raw = "".join(chunks)
    else:
        raw = sheet["A1"].value
    return ComparisonModel.model_validate(json.loads(raw))
