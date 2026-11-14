# Delivery review model findings

## Stage A.1 correction

- The own-words input in `generator/kakali_text.yaml` is loaded verbatim into all 46 named cells. The source file was not changed.
- Q2 2026 demand changed from 163 to 164 FTE through the specified +1 data-engineering demand in Bengaluru. Supply remains 170: 99 base vendor FTE, six Vendor A QA transition FTE, and 65 internal FTE. The total surplus is six and equals the Vendor A QA overlap. Vendor D's six data engineers are baseline supply from 1 April 2026.
- The dual-run is recorded as Vendor A outgoing QA, in a separate `transition overlap (contingency)` category. It is not included in the commercial vendor spend. The $183,600 cost is paid from contingency; the vendor spend targets and $42.84m envelope do not change.
- Follow-up baseline allocation: six Vendor A QA FTE are assigned to London in Q1 2026 with matching six-FTE QA demand, for a zero cell gap. The Q1 vendor base total remains 99 FTE after reallocating the vendor mix. From 1 April, those six FTE are the existing transition overlap and remain excluded from base vendor spend.
- The synthetic `vendor_a_region` is London; v1.5 does not specify a region. Data-engineering role demand is synthetically allocated to Bengaluru based on the golden-source location. The synthetic Q1 DE regional shortfall allocation is Bengaluru −14 FTE, London 0, New York 0, Singapore 0. Q1 DE is the only role-quarter exception; its post-M1 position is −6 on 31 March and zero from 1 April.
- Cell coherence is checked for all 480 role × region × quarter cells. All non-exempt cells are within ±2 FTE; the only exempt cells are the four Q1 DE cells (all non-positive) and the Q2 QA cell in London's Vendor A region (+6). The largest absolute cell gap was 67 FTE before A.1 (Other role 3, New York, Q4 2027) and is 14 FTE after A.1 (Data engineering, Bengaluru, Q1 2026).
- The existing Q3 2027 role bridge and all Stage A controls remain intact. Checks C17–C21c are TRUE after LibreOffice recalculation. There is no cost, vendor-period, annual-funding, or envelope effect from the +1 demand correction.

## Stage B1

- Added only the VENDORS, COMMERCIAL, and PLAN model tabs. The approved sequencing change is applied: B1 settles model tabs only; BENEFITS and all six PDF pages remain for B2.
- The quarterly capacity grid shows first-day-of-quarter positions. Commercial calculations use the §10.1 period averages, so Q4 2026 grid vendor supply is 99 FTE on 1 October while the November–December cost average is 110 FTE.
- VENDORS contains the five-vendor four-dimension scorecard, dated monthly quality series, remediation checkpoints, six A-to-D knowledge-transfer milestones, and three continuity risks. Only Vendor B is amber at the 31 October 2026 cut-off. Its 1.6% exit gate is reached in 54 calendar days; the two-month ≤1.5% amber criterion is not met by October. The May–July parallel run adds $0 spend.
- COMMERCIAL allocates vendor-role FTE-months from base supplier rows using quarter-specific mixes and modeled period month weights. This gives Vendor A 18 QA FTE-months in Jan–Mar 2026. The rate card uses the “calibration factor to the scenario blended rate – assumption” for each year; the factors are 1.0086 (2026), 0.9748 (2027), and 0.9972 (2028), producing the $480/day blended rate ($8,800/FTE-month) in each year. Vendor spend calculates to $10.648m / $12.3552m / $6.1248m, rounding to $10.6m / $12.4m / $6.1m, and $29.128m three-year, rounding to $29.1m. Year 1 splits to $8.712m / $1.936m, rounding to $8.7m / $1.9m. Three-year vendor shares are A 1.7%, B 33.3%, C 31.0%, D 25.8%, E 8.1%; exactly three exceed 25%. The $0.1836m overlap is a separate contingency drawdown; $1.8564m remains unallocated.
- PLAN contains three workstreams, 14 dated milestones with formula-driven status labels, the five-stage critical path and formula-derived 11 working days of float, the six specified decision-rights rows, five RAID rows, four escalation levels, and the two structured recovery plans.
- The spec's §6 schedule-recovery wording conflicts with D4 and the decision log. The workbook follows the log and D4: two reports are phased to phase 2; no squads are crashed. This is the logged decision, not an added recommendation.
- `[Kakali]` remains only in these eight mitigation cells: VENDORS H49:H51 and PLAN F36:F40. No other placeholder remains.

## Verification and scope

- LibreOffice headless recalculation: all 31 checks C01–C28 (including C20b, C21b, and C21c) are TRUE; ALL CHECKS is TRUE. Formula-error scan: zero formula errors.
- `pytest`: 25 passed. Tests cover workbook formulas and checks, verbatim own-words loading, named ranges, seeded reproducibility, role-region tolerance, Q1 Vendor A QA baseline and commercial FTE-month allocation, RAG thresholds and boundaries, vendor quality/remediation, commercial tie-outs, plan outputs, placeholder locations, and firm-name/status-language scans.
- No tie-out remains open. The forecast total spend rounds to the required vendor spend totals; the contingency overlap is excluded from vendor spend.
- This Stage B1 work, including the follow-up baseline correction, is committed locally; no push was performed. BENEFITS and all PDF pages were not started, as required for B1.

## Stage B2

- Added the BENEFITS tab with six register rows, the three-unit adoption funnel, and six formula-driven delivery-health measures. BENEFITS is the eighth tab. The register has two `in-scenario actual (simulated)` rows and four `forecast (simulated)` rows. Its run-budget benefits remain outside the $42.84m envelope. E4.2 and E4.3 are presented from BENEFITS alongside E4.1 because the fixed eight-tab structure places all “is it working” measures together; this follows the Stage B2 brief.
- Added the three-forum cadence block and three formula-linked managing-director monthly measures to PLAN. All 11 page-facing own-words values were read verbatim from `generator/kakali_text.yaml`; that file was not edited. No own-words text failed to fit; all 11 appear in the PDF at a minimum of 7.5 pt.
- Generated the six-page `out/delivery_review.pdf` and `out/page_trace.csv` from the recalculated workbook. The PDF has 14 exhibits across pages 1–4 (3/3/4/4), no exhibits on the cover or cadence page, and a 7.5 pt smallest font. The trace has 1,256 figure rows. E4.2 total self-service share is 107/330 = 32.4%; delivery-health RAG is Green, Green, Amber, Green, Green, Amber.
- LibreOffice headless recalculation: ALL CHECKS is TRUE, all 36 checks are TRUE, and the formula-error scan found zero errors. `pytest`: 47 passed. Visual review of all six 100 dpi page renders found no clipping or text-fit issue.
- Stage B1's Vendor A QA baseline correction remains in place: six London QA FTE match Q1 demand with a zero gap; Q1 base vendor supply remains 99. From 1 April 2026 the six are the existing transition-overlap FTE. COMMERCIAL records 18 Vendor A QA FTE-months in Jan–Mar 2026 inside base vendor spend. Vendor spend remains $10.648m / $12.3552m / $6.1248m (rounding to $10.6m / $12.4m / $6.1m), with the Year 1 $8.712m / $1.936m split (rounding to $8.7m / $1.9m) and exactly three vendor shares above 25%.
- This Stage B2 work is committed locally; no push was performed.

## Stage B2 fix

- F1 — Added PORTFOLIO formula totals, a separate formula-driven ask/funded/deferred/stopped line, and the funded-plus-contingency envelope reconciliation check.
- F2 — Wrapped the E2.3 contingency, escalation, and recovery text in a measured table cell so it stays inside its frame.
- F3 — Wrapped the E4.2 recommendation note and measured its text against the available width and height.
- F4 — Decision log moved to a seventh appendix page – 265.7 pt needed vs 64.2 pt available on page 4; Kakali approved 14 Nov 2026; spec §17 forbids cutting the log, not adding a page.
- F5 — Registered and embedded DejaVu Sans for page text and chart labels; verified the required ≤ and −14 glyphs and absence of ″.
- F6 — Rebuilt E1.2 as a quarterly portfolio-line roadmap with forecast shading, the Q2 lines 3/6 collision, the Q3 replan, and cell-sourced requirements.
- F7 — Drew five critical-path boxes with arrows and read the 11 working-day float from PLAN.
- F8 — Rebuilt E3.1 with four role heatmaps and separate vendor-ramp, bench, and utilisation panels; chart text remains at least 7.5 pt when placed.
- F9 — Added KP1, KP2, and SS1 with their verbatim mitigation text to E3.2.
- F10 — Added formula-driven adoption conversion columns and highlighted the Corporate Services trained-to-active rate.
- F11 — Split amber and red escalation into separate E4.3 columns.
- F12 — Removed supporting mitigations from Cadence; the page contains three forums, three managing-director measures, and cadence rationale.
- F13 — Added fit measurement for paragraphs and table cells plus a page-wide bounding-box overlap check; all generated blocks pass.
- F14 — Recorded the smallest PDF text font (7.50 pt) and chart font at placed size (8.32 pt) separately.
- Test environment — Excluded `.venv` from the firm-name/history-employer repository scan after the virtual environment caused a false positive.

## Stage B2 fix round 2

- G1 – Exhibit counts corrected to 3/3/4/3 (an earlier review brief had them as 3/3/3/4); E3.4 is restored to Vendor workforce and E4.4 sits in the Appendix.
- G2 – Kakali confirmed the role taxonomy on 14 Nov 2026; params now names the six roles, feeding CAPACITY and E3.1 with no figure changes, readable quarter/year axes, keys, panel titles, and a checked chart text layout.
- G3 – Table paragraphs no longer split long words; unbreakable words are measured against cell widths and the funnel and decision-log columns were widened.
- G4 – D5 is shown in a full-width forced-decision box at the foot of page 4; the E4.4 appendix pointer remains.
- G5 – E1.2 labels funded bars by year, separates the forecast label into its legend, and says “the same six data engineers” using the PORTFOLIO cells.

## Stage B2 fix round 3

- G6 – E1.2 y-axis label collision (present since the first review round, found at 220 dpi); reclaimed whitespace inside the page 2 chart frame so all eight y-axis labels have at least 1 pt clearance without reducing font size or moving exhibits; the overlap audit now covers both chart images.
- G7 – Not fitting: the E3.4 notes have 0.5 pt clearance above D3; the 8.2–10.8 pt gaps on pages 2–3 would require 7.7–10.3 pt more, while only 1.7 pt remains before the footer, so the spacing is unchanged.
