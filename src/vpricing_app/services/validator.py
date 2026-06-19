from decimal import Decimal

from vpricing_app.models import ImportWarning, OriginalBQ, VendorQuote, WarningLevel, normalize_text
from vpricing_app.services.matcher import ComparisonRow


def _different(left: object, right: object) -> bool:
    return normalize_text(left) != normalize_text(right)


def validate_comparison(
    original: OriginalBQ,
    vendors: list[VendorQuote],
    rows_by_package: dict[str, list[ComparisonRow]],
) -> list[ImportWarning]:
    warnings: list[ImportWarning] = []

    for package_name, rows in rows_by_package.items():
        for row in rows:
            for vendor in vendors:
                vendor_line = row.vendor_lines.get(vendor.vendor_name)
                if vendor_line is None:
                    if row.original_line is not None:
                        warnings.append(
                            ImportWarning(
                                level=WarningLevel.REVIEW,
                                code="missing_original_item",
                                vendor_name=vendor.vendor_name,
                                package=package_name,
                                row_label=row.original_line.description,
                                message=f"{vendor.vendor_name} did not price original item: {row.original_line.description}",
                            )
                        )
                    continue
                if vendor_line.is_vendor_added:
                    warnings.append(
                        ImportWarning(
                            level=WarningLevel.REVIEW,
                            code="vendor_added_item",
                            vendor_name=vendor.vendor_name,
                            package=package_name,
                            row_label=vendor_line.description,
                            message=f"{vendor.vendor_name} added item not found in original BQ: {vendor_line.description}",
                        )
                    )
                if row.original_line is not None:
                    protected_pairs = [
                        ("quantity", row.original_line.quantity, vendor_line.quantity),
                        ("revised quantity", row.original_line.revised_quantity, vendor_line.revised_quantity),
                        ("unit", row.original_line.unit, vendor_line.unit),
                        ("description", row.original_line.description, vendor_line.description),
                    ]
                    for label, original_value, vendor_value in protected_pairs:
                        if _different(original_value, vendor_value):
                            warnings.append(
                                ImportWarning(
                                    level=WarningLevel.REVIEW,
                                    code="protected_field_changed",
                                    vendor_name=vendor.vendor_name,
                                    package=package_name,
                                    row_label=row.original_line.description,
                                    message=f"{vendor.vendor_name} changed {label}: {original_value} -> {vendor_value}",
                                )
                            )
                if vendor_line.confidence != "high":
                    warnings.append(
                        ImportWarning(
                            level=WarningLevel.REVIEW,
                            code="pdf_extraction_review",
                            vendor_name=vendor.vendor_name,
                            package=package_name,
                            row_label=vendor_line.description,
                            message=f"{vendor.vendor_name} row requires PDF extraction review: {vendor_line.description}",
                        )
                    )

    for vendor in vendors:
        for package_name, money in vendor.package_totals.items():
            if abs(money.calculated_total - money.total) > Decimal("0.01"):
                warnings.append(
                    ImportWarning(
                        level=WarningLevel.REVIEW,
                        code="package_total_mismatch",
                        vendor_name=vendor.vendor_name,
                        package=package_name,
                        message=(
                            f"{vendor.vendor_name} {package_name} material plus labor "
                            f"does not equal total: {money.calculated_total} vs {money.total}"
                        ),
                    )
                )

    return warnings
