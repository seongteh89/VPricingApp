import re


INVALID_FILENAME_CHARS = r'[<>:"/\\|?*#@]'


def _clean_part(value: str, max_length: int) -> str:
    cleaned = re.sub(INVALID_FILENAME_CHARS, " ", value)
    cleaned = " ".join(cleaned.split())
    return cleaned[:max_length].rstrip()


def build_output_filename(project_name: str, vendor_names: list[str], revision: str | None = None) -> str:
    project = _clean_part(project_name or "Project", 45)
    vendor_set = _clean_part(" vs ".join(vendor_names) or "Vendors", 80)
    if revision:
        rev = _clean_part(revision, 30)
        return f"Quote Comparison - {project} - {vendor_set} - Rev {rev}.xlsx"
    return f"Quote Comparison - {project} - {vendor_set}.xlsx"
