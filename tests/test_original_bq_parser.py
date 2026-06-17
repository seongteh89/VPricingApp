from vpricing_app.parsers.original_bq import parse_original_bq


def test_parse_original_bq_extracts_project_and_lines(sample_original_bq):
    original = parse_original_bq(sample_original_bq)

    assert original.project_name == "#Fire#Sample Project"
    assert list(original.packages) == ["PACKAGE 1 with Meter Run"]
    lines = original.packages["PACKAGE 1 with Meter Run"]
    assert len(lines) == 2
    assert lines[0].section == "A"
    assert lines[0].section_description == "ADDRESSABLE FIRE ALARM SYSTEM"
    assert lines[0].item_no == "1"
    assert lines[0].description == "Sub Alarm Panel"
    assert lines[0].quantity == 5
    assert lines[0].revised_quantity == 5
    assert lines[0].unit == "Nos"


def test_parse_original_bq_accepts_request_for_quotation_sheet(sample_original_rfq):
    original = parse_original_bq(sample_original_rfq)

    assert list(original.packages) == ["Request For Quotation"]
    lines = original.packages["Request For Quotation"]
    assert len(lines) == 2
    assert lines[0].description == "PBB : Automatic Fire Sprinkler System"
    assert lines[0].quantity == 1
    assert lines[0].revised_quantity == 12
