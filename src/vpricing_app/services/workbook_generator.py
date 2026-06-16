import json
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from vpricing_app.models import ComparisonModel
from vpricing_app.services.filename_builder import build_output_filename
from vpricing_app.services.matcher import build_comparison_rows


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
ADDED_FILL = PatternFill("solid", fgColor="FFF2CC")
LOW_FILL = PatternFill("solid", fgColor="C6EFCE")
METADATA_MARKER = "comparison_model_json_v1"
METADATA_CHUNK_SIZE = 30000


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
    summary.append(["Vendor", "Grand Total"])
    for vendor in model.vendors:
        summary.append([vendor.vendor_name, float(vendor.grand_total)])
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
