from pathlib import Path
from tempfile import TemporaryDirectory, gettempdir

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from vpricing_app.parsers.original_bq import parse_original_bq
from vpricing_app.parsers.vendor_excel import parse_vendor_excel
from vpricing_app.parsers.vendor_pdf import parse_vendor_pdf
from vpricing_app.services.comparison_service import build_new_comparison, replace_vendor_quote
from vpricing_app.services.workbook_generator import (
    export_comparison_workbook,
    export_updated_template_workbook,
    load_comparison_metadata,
)

app = FastAPI(title="V Pricing Quote Comparison")


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


async def _save_upload(upload: UploadFile, folder: Path) -> Path:
    path = folder / upload.filename
    path.write_bytes(await upload.read())
    return path


async def _parse_vendor_uploads(vendors: list[UploadFile], vendor_names: str, folder: Path):
    names = [name.strip() for name in vendor_names.split(",") if name.strip()]
    vendor_models = []
    for upload, vendor_name in zip(vendors, names, strict=False):
        vendor_path = await _save_upload(upload, folder)
        if vendor_path.suffix.lower() == ".pdf":
            vendor_models.append(parse_vendor_pdf(vendor_path, vendor_name))
        else:
            vendor_models.append(parse_vendor_excel(vendor_path, vendor_name))
    return vendor_models


def _preview_payload(comparison):
    return {
        "project_name": comparison.project_name,
        "vendors": [vendor.vendor_name for vendor in comparison.vendors],
        "warnings": [warning.model_dump(mode="json") for warning in comparison.warnings],
    }


@app.post("/api/comparisons/preview")
async def preview_comparison(
    original: UploadFile = File(...),
    vendors: list[UploadFile] = File(...),
    vendor_names: str = Form(...),
):
    with TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        original_path = await _save_upload(original, tmp_path)
        original_model = parse_original_bq(original_path)
        vendor_models = await _parse_vendor_uploads(vendors, vendor_names, tmp_path)
        comparison = build_new_comparison(original_model, vendor_models)
        return _preview_payload(comparison)


@app.post("/api/comparisons/export")
async def export_comparison(
    original: UploadFile = File(...),
    vendors: list[UploadFile] = File(...),
    vendor_names: str = Form(...),
) -> FileResponse:
    with TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        original_path = await _save_upload(original, tmp_path)
        original_model = parse_original_bq(original_path)
        vendor_models = await _parse_vendor_uploads(vendors, vendor_names, tmp_path)
        comparison = build_new_comparison(original_model, vendor_models)
        output_dir = Path(gettempdir()) / "vpricing_exports"
        output_path = export_comparison_workbook(comparison, output_dir, original_template_path=original_path)
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=output_path.name,
        )


@app.post("/api/comparisons/update/preview")
async def preview_update_comparison(
    existing_comparison: UploadFile = File(...),
    revised_vendor: UploadFile = File(...),
    vendor_name: str = Form(...),
    revision: str | None = Form(None),
):
    with TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        existing_path = await _save_upload(existing_comparison, tmp_path)
        revised_path = await _save_upload(revised_vendor, tmp_path)
        existing_model = load_comparison_metadata(existing_path)
        if revised_path.suffix.lower() == ".pdf":
            revised_model = parse_vendor_pdf(revised_path, vendor_name)
        else:
            revised_model = parse_vendor_excel(revised_path, vendor_name)
        revised_model.revision = revision
        updated_model = replace_vendor_quote(existing_model, revised_model)
        return _preview_payload(updated_model)


@app.post("/api/comparisons/update")
async def update_comparison(
    existing_comparison: UploadFile = File(...),
    revised_vendor: UploadFile = File(...),
    vendor_name: str = Form(...),
    revision: str | None = Form(None),
) -> FileResponse:
    with TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        existing_path = await _save_upload(existing_comparison, tmp_path)
        revised_path = await _save_upload(revised_vendor, tmp_path)
        existing_model = load_comparison_metadata(existing_path)
        if revised_path.suffix.lower() == ".pdf":
            revised_model = parse_vendor_pdf(revised_path, vendor_name)
        else:
            revised_model = parse_vendor_excel(revised_path, vendor_name)
        revised_model.revision = revision
        updated_model = replace_vendor_quote(existing_model, revised_model)
        output_dir = Path(gettempdir()) / "vpricing_exports"
        output_path = export_updated_template_workbook(
            updated_model,
            existing_path,
            revised_model,
            output_dir,
            revision=revision,
        )
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=output_path.name,
        )


app.mount("/", StaticFiles(directory="src/vpricing_app/static", html=True, check_dir=False), name="static")
