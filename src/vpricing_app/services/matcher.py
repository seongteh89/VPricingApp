from pydantic import BaseModel

from vpricing_app.models import OriginalBQ, QuoteLine, VendorQuote, normalize_text


class ComparisonRow(BaseModel):
    package: str
    section: str | None = None
    original_line: QuoteLine | None = None
    vendor_lines: dict[str, QuoteLine]


def _candidate_keys(line: QuoteLine) -> list[str]:
    return [
        line.match_key,
        "|".join([normalize_text(line.package), normalize_text(line.section), normalize_text(line.description)]),
    ]


def build_comparison_rows(original: OriginalBQ, vendors: list[VendorQuote]) -> dict[str, list[ComparisonRow]]:
    result: dict[str, list[ComparisonRow]] = {}

    for package_name, original_lines in original.packages.items():
        rows: list[ComparisonRow] = [
            ComparisonRow(package=package_name, section=line.section, original_line=line, vendor_lines={})
            for line in original_lines
        ]
        index: dict[str, ComparisonRow] = {}
        for row in rows:
            if row.original_line:
                for key in _candidate_keys(row.original_line):
                    index.setdefault(key, row)

        for vendor in vendors:
            for vendor_line in vendor.packages.get(package_name, []):
                matched = None
                for key in _candidate_keys(vendor_line):
                    if key in index:
                        matched = index[key]
                        break
                if matched is None:
                    vendor_line.is_vendor_added = True
                    matched = ComparisonRow(
                        package=package_name,
                        section=vendor_line.section,
                        original_line=None,
                        vendor_lines={},
                    )
                    rows.append(matched)
                matched.vendor_lines[vendor.vendor_name] = vendor_line

        result[package_name] = rows

    return result
