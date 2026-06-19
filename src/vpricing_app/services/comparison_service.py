import re

from vpricing_app.models import ComparisonModel, OriginalBQ, VendorQuote
from vpricing_app.services.matcher import build_comparison_rows
from vpricing_app.services.validator import validate_comparison


def _package_number(name: str) -> str | None:
    match = re.search(r"package\s*(\d+)", name, flags=re.IGNORECASE)
    return match.group(1) if match else None


def _align_vendor_packages(original: OriginalBQ, vendor: VendorQuote) -> VendorQuote:
    package_map = {
        number: original_name
        for original_name in original.packages
        if (number := _package_number(original_name)) is not None
    }
    aligned = vendor.model_copy(deep=True)
    aligned_packages = {}
    for package_name, lines in aligned.packages.items():
        target_name = package_map.get(_package_number(package_name), package_name)
        for line in lines:
            line.package = target_name
        aligned_packages.setdefault(target_name, []).extend(lines)
    aligned.packages = aligned_packages

    aligned_totals = {}
    for package_name, total in aligned.package_totals.items():
        target_name = package_map.get(_package_number(package_name), package_name)
        aligned_totals[target_name] = total
    aligned.package_totals = aligned_totals
    return aligned


def build_new_comparison(original: OriginalBQ, vendors: list[VendorQuote]) -> ComparisonModel:
    aligned_vendors = [_align_vendor_packages(original, vendor) for vendor in vendors]
    rows = build_comparison_rows(original, aligned_vendors)
    warnings = validate_comparison(original, aligned_vendors, rows)
    return ComparisonModel(project_name=original.project_name, original=original, vendors=aligned_vendors, warnings=warnings)


def replace_vendor_quote(existing: ComparisonModel, revised_quote: VendorQuote) -> ComparisonModel:
    vendors: list[VendorQuote] = []
    replaced = False
    for vendor in existing.vendors:
        if vendor.vendor_name.lower() == revised_quote.vendor_name.lower():
            vendors.append(revised_quote)
            replaced = True
        else:
            vendors.append(vendor)
    if not replaced:
        vendors.append(revised_quote)
    return build_new_comparison(existing.original, vendors)
