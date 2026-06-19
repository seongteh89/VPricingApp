from fastapi.testclient import TestClient

from vpricing_app.main import app
from vpricing_app.parsers.original_bq import parse_original_bq
from vpricing_app.parsers.vendor_excel import parse_vendor_excel
from vpricing_app.services.comparison_service import build_new_comparison
from vpricing_app.services.workbook_generator import export_comparison_workbook


def test_health_endpoint_returns_ok():
    client = TestClient(app)
    response = client.get("/api/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_index_page_contains_new_and_update_upload_controls():
    client = TestClient(app)
    response = client.get("/")

    assert response.status_code == 200
    assert 'id="original"' in response.text
    assert 'id="vendors"' in response.text
    assert 'id="existing-comparison"' in response.text
    assert 'id="revised-vendor-name"' in response.text
    assert "Insert vendor names, e.g. T-Tech, YB" in response.text
    assert "Insert revised vendor name" in response.text
    assert "Insert revision, e.g. R2 or 2026-06-16" in response.text


def test_new_comparison_endpoint_exports_workbook(sample_original_bq, sample_vendor_excel):
    client = TestClient(app)
    with sample_original_bq.open("rb") as original_file, sample_vendor_excel.open("rb") as vendor_file:
        response = client.post(
            "/api/comparisons/export",
            files=[
                (
                    "original",
                    (
                        "Sample_original.xlsx",
                        original_file,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                ),
                (
                    "vendors",
                    (
                        "Rensar quote.xlsx",
                        vendor_file,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                ),
            ],
            data={"vendor_names": "Rensar"},
        )

    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert "Quote%20Comparison" in response.headers["content-disposition"]
    assert "Rensar" in response.headers["content-disposition"]


def test_new_comparison_preview_returns_review_warnings(sample_original_bq, sample_vendor_excel):
    client = TestClient(app)
    with sample_original_bq.open("rb") as original_file, sample_vendor_excel.open("rb") as vendor_file:
        response = client.post(
            "/api/comparisons/preview",
            files=[
                (
                    "original",
                    (
                        "Sample_original.xlsx",
                        original_file,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                ),
                (
                    "vendors",
                    (
                        "Rensar quote.xlsx",
                        vendor_file,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                ),
            ],
            data={"vendor_names": "Rensar"},
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["project_name"] == "#Fire#Sample Project"
    assert payload["vendors"] == ["Rensar"]
    assert "protected_field_changed" in {warning["code"] for warning in payload["warnings"]}


def test_update_comparison_endpoint_replaces_vendor(tmp_path, sample_original_bq, sample_vendor_excel):
    original = parse_original_bq(sample_original_bq)
    vendor = parse_vendor_excel(sample_vendor_excel, "Rensar")
    existing = export_comparison_workbook(build_new_comparison(original, [vendor]), tmp_path)

    client = TestClient(app)
    with existing.open("rb") as existing_file, sample_vendor_excel.open("rb") as revised_file:
        response = client.post(
            "/api/comparisons/update",
            files=[
                (
                    "existing_comparison",
                    (
                        existing.name,
                        existing_file,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                ),
                (
                    "revised_vendor",
                    (
                        "Rensar revised.xlsx",
                        revised_file,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                ),
            ],
            data={"vendor_name": "Rensar", "revision": "R2"},
        )

    assert response.status_code == 200
    assert "Rensar" in response.headers["content-disposition"]
    assert "Rev%20R2" in response.headers["content-disposition"]


def test_index_page_contains_preview_review_controls():
    client = TestClient(app)
    response = client.get("/")

    assert response.status_code == 200
    assert 'id="preview-button"' in response.text
    assert 'id="review-panel"' in response.text
