import json
import re
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from vpricing_app.models import ComparisonModel, Money, VendorQuote
from vpricing_app.services.filename_builder import build_output_filename
from vpricing_app.services.matcher import build_comparison_rows


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


def export_comparison_workbook(model: ComparisonModel, output_dir: str | Path, revision: str | None = None) -> Path:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = build_output_filename(model.project_name, [vendor.vendor_name for vendor in model.vendors], revision)
    output_path = output_dir / filename

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    _write_summary(summary, model)
    _autosize(summary)

    rows_by_package = build_comparison_rows(model.original, model.vendors)
    for package_name, rows in rows_by_package.items():
        ws = wb.create_sheet(package_name[:31])
        _write_headers(ws, [vendor.vendor_name for vendor in model.vendors])
        for row in rows:
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
        _write_package_total_row(ws, package_name, model.vendors)
        _autosize(ws)

    change_log = wb.create_sheet("Change Log")
    change_log.append(["Level", "Code", "Vendor", "Package", "Row", "Message"])
    for warning in model.warnings:
        change_log.append(
            [warning.level.value, warning.code, warning.vendor_name, warning.package, warning.row_label, warning.message]
        )
    _autosize(change_log)

    warnings = wb.create_sheet("Import Warnings")
    warnings.append(["Level", "Code", "Vendor", "Package", "Message"])
    for warning in model.warnings:
        warnings.append([warning.level.value, warning.code, warning.vendor_name, warning.package, warning.message])
    _autosize(warnings)

    metadata = wb.create_sheet("System Metadata")
    metadata.sheet_state = "hidden"
    metadata["A1"] = METADATA_MARKER
    raw_metadata = model.model_dump_json()
    for index in range(0, len(raw_metadata), METADATA_CHUNK_SIZE):
        row = index // METADATA_CHUNK_SIZE + 2
        metadata.cell(row=row, column=1).value = raw_metadata[index : index + METADATA_CHUNK_SIZE]

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
