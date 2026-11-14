# Enterprise Transformation Delivery Review

**Problem:** A transformation portfolio has more demand than capacity (205 vs 186 FTE), a vendor above its quality threshold (4.1% vs 1.5%), a regulatory date with no float and a fixed $42.84m envelope. Leadership needs decisions, not status reports.

**What it does:** An eight-tab Excel model generates a seven-page review pack, and each core page ends in a dated decision the pack forces (D1–D5).

**Evidence:** 37 model checks TRUE; all 1,284 figures printed in the PDF traced to their workbook cells.

Simulated data – not observed results. The firm is generic and unnamed.

## Outputs

- `out/delivery_review_model.xlsx` — 8 tabs: ASSUMPTIONS, PORTFOLIO, CAPACITY, DECISIONS, VENDORS, COMMERCIAL, PLAN, BENEFITS.
- `out/delivery_review.pdf` — 7 pages: Cover; Portfolio; Governance; Vendor workforce; Is it working; Cadence; Appendix – decision log.
- `out/page_trace.csv` — every printed figure → tab!cell.

## How the numbers hold

Inputs only from `generator/params.yaml` (each entry has a `basis`) and `generator/kakali_text.yaml` (own-words text, loaded verbatim); live formulas; ALL CHECKS block: 37 TRUE. PDF figures read from recalculated cells. Trace CSV rows: 1284.

Decisions D1–D5 are the decisions the pack forces, presented as such. Status wording: “in-scenario actual (simulated)”. Currency: USD; amounts in the PDF are shown in $m.

## Rebuild

Python 3.14; install dependencies with `pip install -r requirements.txt`. LibreOffice headless is needed for recalculation.

From the repository root:

```powershell
python generator/build_pages.py
python -m pytest -q
```

## Links

- [Case study](https://kakali-bandyopadhyay.github.io/case-studies/transformation-delivery-review.html)
- [Portfolio filter](https://kakali-bandyopadhyay.github.io/?track=transformation)
