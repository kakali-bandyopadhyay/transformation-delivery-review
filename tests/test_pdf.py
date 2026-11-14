import csv
import re
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / ".b2deps"))
from pypdf import PdfReader

from generator.build_model import KAKALI_TEXT

PDF = ROOT / "out" / "delivery_review.pdf"
TRACE = ROOT / "out" / "page_trace.csv"
MODEL = ROOT / "out" / "_recalc" / "delivery_review_model.xlsx"
NUMBER_RE = re.compile(r"(?<![A-Za-z])(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?")
DATE_RE = re.compile(r"\b(?:\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}|Q[1-4]\s+\d{4})\b")


def pdf_pages():
    return PdfReader(str(PDF)).pages


def norm(text):
    return re.sub(r"\s+", " ", text.replace("-\n", "-")).strip()


def raw_value(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        return format(value, ".15g")
    return "" if value is None else str(value)


def test_pdf_has_seven_pages_in_order_and_fourteen_exhibits():
    pages = pdf_pages()
    assert len(pages) == 7
    text = [page.extract_text() or "" for page in pages]
    assert "Enterprise Transformation Delivery Review" in text[0]
    for physical_page, title in enumerate(
        ("1 | Portfolio", "2 | Governance", "3 | Vendor workforce", "4 | Is it working", "Cadence"), 1
    ):
        assert title in text[physical_page]
    counts = []
    for page_text in text[1:5]:
        page_text = re.sub(r"E4\.4 Decision log D1.*?full log on page 7 \(appendix\)", "", page_text)
        counts.append(len(set(re.findall(r"E[1-4]\.\d", page_text))))
    assert counts == [3, 3, 4, 3]
    assert "Appendix | Decision log" in text[6]
    assert len(set(re.findall(r"E[1-4]\.\d", text[6]))) == 1
    assert not re.search(r"E[1-4]\.\d", text[0])
    assert not re.search(r"E[1-4]\.\d", text[5])
    assert "Vendor spend by year" in text[3]


def test_pdf_simulation_language_and_prohibited_terms():
    pages = pdf_pages()
    text = [page.extract_text() or "" for page in pages]
    assert "Simulated data – not observed results" in text[0]
    for page_number, page_text in enumerate(text, 1):
        assert f"page {page_number} of 7" in page_text
    all_text = "\n".join(text)
    assert not re.search(r"\bcrashed\b", all_text, flags=re.I)
    stripped = re.sub(r"in-scenario\s+actual\s*\(simulated\)", "", all_text, flags=re.I)
    assert not re.search(r"\bactual\b", stripped, flags=re.I)


def test_page_trace_covers_pdf_numbers_and_matches_recalculated_cells():
    pages = pdf_pages()
    with TRACE.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    wb = load_workbook(MODEL, data_only=True)
    assert rows
    assert sum(not row["workbook value"] for row in rows) == 0
    by_page = {}
    for row in rows:
        by_page.setdefault(int(row["page"]), []).append(row)
        sheet, cell = row["tab!cell"].split("!")
        assert raw_value(wb[sheet][cell].value) == row["workbook value"]
    for page_num, page in enumerate(pages, 1):
        page_text = page.extract_text() or ""
        pdf_numbers = set(NUMBER_RE.findall(page_text))
        trace_numbers = {
            token
            for row in by_page.get(page_num, [])
            for token in NUMBER_RE.findall(row["text as printed"])
        }
        assert pdf_numbers <= trace_numbers, f"untraced numbers on PDF page {page_num}: {pdf_numbers - trace_numbers}"


def test_decision_forced_boxes_match_decision_log():
    pages = pdf_pages()
    wb = load_workbook(MODEL, data_only=True)
    decisions = wb["DECISIONS"]
    # DECISIONS rows map to portfolio/governance/vendor-workforce/benefits PDF pages after the cover.
    for row, physical_page in ((6, 2), (8, 3), (7, 4), (9, 5)):
        page_text = norm(pages[physical_page - 1].extract_text() or "")
        assert "Decision forced" in page_text
        for col in (1, 2, 3, 6):
            value = decisions.cell(row, col).value
            expected = value.strftime("%b %Y") if isinstance(value, (datetime, date)) else str(value)
            assert norm(expected) in page_text


def test_appendix_decision_log_is_verbatim_and_has_no_forced_box():
    pages = pdf_pages()
    text = norm(pages[6].extract_text() or "")
    decisions = load_workbook(MODEL, data_only=True)["DECISIONS"]
    for row in range(5, 10):
        for col in range(1, 10):
            value = decisions.cell(row, col).value
            expected = value.strftime("%d %b %Y") if isinstance(value, (datetime, date)) else str(value)
            assert norm(expected) in text
    assert "Decision forced" not in text
    assert "full log on page 7 (appendix)" in norm(pages[4].extract_text() or "")


def test_text_fit_and_layout_audit_includes_chart_font():
    audit = __import__("json").loads((ROOT / "out" / "_preview" / "layout_audit.json").read_text(encoding="utf-8"))
    assert audit["text_fit_blocks"] > 0
    assert audit["overlap_check"] == "passed"
    assert audit["smallest_text_font_pt"] >= 7.5
    assert audit["smallest_chart_font_at_placed_size_pt"] >= 7.5
    assert audit["word_fit_checks"] > 0
    assert audit["chart_text_overlap_checks"] > 0
    assert audit["chart_text_overlap_count"] == 0
    assert audit["chart_images_audited"] == 2
    assert audit["chart_text_min_clearance_pt"] >= 1.0
    roadmap = next(item for item in audit["capacity_chart_labels"] if "roadmap_y_tick_labels" in item)
    assert len(roadmap["roadmap_y_tick_labels"]) == 8
    assert roadmap["roadmap_min_clearance_pt"] >= 1.0
    capacity = next(item for item in audit["capacity_chart_labels"] if "roles" in item)
    assert capacity["capacity_chart_min_clearance_pt"] >= 1.0
    assert capacity["roles"][4:] == [
        "Back-end development", "Front-end development", "Cloud and DevOps", "Architecture",
        "Project and delivery management", "Scrum and agile delivery",
    ]
    assert capacity["quarters"] == [f"Q{q} {year}" for year in (2026, 2027, 2028) for q in range(1, 5)]
    roadmap = next(item for item in audit["capacity_chart_labels"] if "roadmap_legend" in item)
    assert roadmap["roadmap_legend"] == ["2026 funded", "2027 funded", "2028 funded", "Forecast after 31 Oct 2026"]


def test_tables_do_not_split_words_and_requested_role_names_are_in_workbook():
    text = norm(" ".join(page.extract_text() or "" for page in pdf_pages()))
    for split,whole in (("Invite d","Invited"),("Traine d","Trained"),("Habitu al","Habitual"),
                        ("D-numb er","D-number"),("Data-engineeri ng","Data-engineering")):
        assert split not in text
        assert whole in text
    wb = load_workbook(MODEL, data_only=True)
    capacity = wb["CAPACITY"]
    names = {capacity.cell(r,1).value for r in range(5, capacity.max_row + 1)}
    assert {"Back-end development", "Front-end development", "Cloud and DevOps", "Architecture",
            "Project and delivery management", "Scrum and agile delivery"} <= names


def test_e12_finding_and_d5_full_width_foot_box():
    pages = pdf_pages()
    assert "the same six data engineers" in norm(pages[1].extract_text() or "").casefold()
    audit = __import__("json").loads((ROOT / "out" / "_preview" / "layout_audit.json").read_text(encoding="utf-8"))
    page_five_tables = [b for b in audit["layout_blocks"] if b[0] == 5 and b[1] == "table"]
    assert any(abs(block[4] - (841.88976378 - 50)) < 1 for block in page_five_tables)


def test_dates_have_years_and_future_dates_are_marked_forecast():
    for page_num, page in enumerate(pdf_pages(), 1):
        text = page.extract_text() or ""
        for match in re.finditer(r"\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b(?!\s+\d{4})", text):
            assert False, f"date without year on PDF page {page_num}: {match.group(0)}"
        for match in DATE_RE.finditer(text):
            if match.group(0).endswith(("2027", "2028", "2029")):
                nearby = text[max(0, match.start() - 75):match.end() + 75]
                assert "forecast" in nearby.casefold(), f"future date lacks forecast label on page {page_num}: {match.group(0)}"


def test_all_eleven_kakali_page_text_values_appear_verbatim():
    text = norm(" ".join(page.extract_text() or "" for page in pdf_pages()))
    values = list(KAKALI_TEXT["PAGE_TEXT"].values()) + list(KAKALI_TEXT["VENDORS"].values()) + list(KAKALI_TEXT["PLAN"].values())
    assert len(values) == 11
    for value in values:
        assert norm(value) in text
    assert "[Kakali" not in text


def test_embedded_font_preserves_required_glyphs():
    pages = pdf_pages()
    page_text = "\n".join(page.extract_text() or "" for page in pages)
    assert "≤" in page_text
    assert "−14" in page_text
    assert "″" not in page_text


def test_model_deliverable_checks_are_all_true_after_recalc():
    wb = load_workbook(MODEL, data_only=True)
    assumptions = wb["ASSUMPTIONS"]
    assert assumptions["K113"].value is True
    check_values = [assumptions.cell(row, 11).value for row in range(76, 113)]
    assert len(check_values) == 37
    assert all(value is True or value == 1 for value in check_values)
    errors = [
        (sheet.title, cell.coordinate, cell.value)
        for sheet in wb.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value in {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}
    ]
    assert errors == []

