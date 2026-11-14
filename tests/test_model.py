import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from generator import build_model


def blocked_tokens():
    codes = [
        [66,108,97,99,107,115,116,111,110,101],
        [66,88,84,105],
        [66,88],
        [68,101,108,111,105,116,116,101],
        [80,119,67],
        [70,105,100,101,108,105,116,121],
        [67,111,103,110,105,122,97,110,116],
        [69,81,84],
        [66,108,97,99,107,82,111,99,107],
        [68,46,69,46,32,83,104,97,119],
        [83,119,105,115,115,32,82,101],
        [74,80,77,111,114,103,97,110,67,104,97,115,101],
        [71,111,108,100,109,97,110,32,83,97,99,104,115],
        [67,105,116,105],
    ]
    return ["".join(map(chr, values)).casefold() for values in codes]


def workbook_strings(path):
    wb = load_workbook(path, data_only=False, read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    yield cell.value


def test_eight_b2_model_tabs_in_order():
    wb = load_workbook(ROOT / "out" / "delivery_review_model.xlsx", data_only=False, read_only=True)
    assert wb.sheetnames == ["ASSUMPTIONS", "PORTFOLIO", "CAPACITY", "DECISIONS", "VENDORS", "COMMERCIAL", "PLAN", "BENEFITS"]


def test_libreoffice_recalculated_all_checks_true():
    path = ROOT / "out" / "delivery_review_model.xlsx"
    formulas = load_workbook(path, data_only=False, read_only=True)
    values = load_workbook(path, data_only=True, read_only=True)
    assert values["ASSUMPTIONS"]["K113"].value in (True,1)
    for row in range(76, 113):
        assert formulas["ASSUMPTIONS"].cell(row, 11).value.startswith("=")
        assert values["ASSUMPTIONS"].cell(row, 11).value in (True,1)


def test_derived_outputs_are_formulas():
    wb = load_workbook(ROOT / "out" / "delivery_review_model.xlsx", data_only=False, read_only=True)
    p = wb["PORTFOLIO"]
    assert all(p.cell(row, 13).value.startswith("=") for row in range(5, 18))
    assert all(p.cell(row, 3).value.startswith("=") for row in range(21, 29))
    c = wb["CAPACITY"]
    assert all(c.cell(row, col).value.startswith("=") for row in range(5, 17) for col in range(3, 9))
    assert all(c.cell(row, col).value.startswith("=") for row in range(21, 26) for col in (2, 4, 9, 10))
    a = wb["ASSUMPTIONS"]
    rows = {a.cell(r, 2).value: r for r in range(5, 75)}
    for name in ["Contingency", "Scenario envelope", "Dual-run working days", "Paid, not delivering working days", "A-to-D overlap cost", "Unallocated contingency", "Vendor cost per FTE-month"]:
        assert a.cell(rows[name], 3).value.startswith("=")

def test_no_formula_error_values():
    wb = load_workbook(ROOT / "out" / "delivery_review_model.xlsx", data_only=True, read_only=True)
    errors = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value in {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"}:
                    errors.append((ws.title, cell.coordinate, cell.value))
    assert errors == []


def test_kakali_text_is_loaded_verbatim():
    wb = load_workbook(ROOT / "out" / "delivery_review_model.xlsx", data_only=True, read_only=True)
    for sheet_name, cells in build_model.KAKALI_TEXT.items():
        if sheet_name == "version":
            continue
        if sheet_name == "PAGE_TEXT":
            ws=wb["ASSUMPTIONS"]
            text_by_id={row[0]:row[1] for row in ws.iter_rows(min_row=1,max_col=2,values_only=True) if row[0] in cells}
            for key,expected in cells.items():
                assert text_by_id[key]==expected
        else:
            for coordinate, expected in cells.items():
                assert wb[sheet_name][coordinate].value == expected
    assert len(build_model.KAKALI_TEXT["DECISIONS"]) == 15
    assert len(build_model.KAKALI_TEXT["PORTFOLIO"]) == 31
    assert len(build_model.KAKALI_TEXT["VENDORS"]) == 3
    assert len(build_model.KAKALI_TEXT["PLAN"]) == 5
    assert len(build_model.KAKALI_TEXT["PAGE_TEXT"]) == 3
    assert sum(len(cells) for cells in build_model.KAKALI_TEXT.values() if isinstance(cells,dict))==57
    assert not any("[Kakali" in value for cells in build_model.KAKALI_TEXT.values() if isinstance(cells,dict) for value in cells.values())


def test_no_kakali_placeholders_remain():
    wb = load_workbook(ROOT / "out" / "delivery_review_model.xlsx", data_only=True, read_only=True)
    observed = []
    for ws in wb:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value == "[Kakali]":
                    observed.append((ws.title, cell.coordinate))
    assert observed == []


def test_seed_reproducibility_and_constraints():
    phases1 = build_model.phase_funded_lines()
    phases2 = build_model.phase_funded_lines()
    assert phases1 == phases2
    demand1, supply1, vendor1 = build_model.capacity_inputs()
    demand2, supply2, vendor2 = build_model.capacity_inputs()
    assert demand1 == demand2
    assert supply1 == supply2
    assert vendor1 == vendor2
    assert [round(sum(phases1[item["id"]][j] for item in build_model.FUNDED_ROWS), 1) for j in range(3)] == [14.2, 17.6, 9.0]
    assert phases1["5a"] == [0.9, 0.0, 0.0]
    peak_vendor = sum(row[4] for row in supply1 if row[3] == "Q3 2027" and row[1] != "Internal")
    peak_internal = sum(row[4] for row in supply1 if row[3] == "Q3 2027" and row[1] == "Internal")
    peak_demand = sum(row[3] for row in demand1 if row[2] == "Q3 2027")
    assert (peak_vendor, peak_internal, peak_demand) == (121, 65, 186)
    dual_run_rows = [row for row in supply1 if row[3] == "Q2 2026" and row[5] == "transition overlap (contingency)"]
    dual_run = sum(row[4] for row in dual_run_rows)
    assert dual_run == 6
    assert all(row[0] == "QA" and row[1] == "Vendor A" and row[2] == build_model.PARAMS["vendor_a_region"] for row in dual_run_rows)
    q2d = sum(row[3] for row in demand1 if row[2] == "Q2 2026")
    assert q2d == 164
    for row in demand1:
        if row[2] == "Q2 2026" and row[0] == "Data engineering" and row[1] == "Bengaluru":
            assert row[3] >= 1
    roles = build_model.PARAMS["capacity"]["role_taxonomy"]["names"]
    regions = build_model.PARAMS["capacity"]["regions"]
    demand_cells = {(r, g, q): n for r, g, q, n in demand1}
    supply_cells = {}
    for role, supplier, region, quarter, n, category in supply1:
        supply_cells[(role, region, quarter)] = supply_cells.get((role, region, quarter), 0) + n
    tol = build_model.PARAMS["cell_tolerance_fte"]
    for quarter in build_model.QUARTERS:
        for role in roles:
            for region in regions:
                gap = supply_cells[(role, region, quarter)] - demand_cells[(role, region, quarter)]
                if quarter == "Q1 2026" and role == "Data engineering":
                    assert gap <= 0
                elif quarter == "Q2 2026" and role == "QA" and region == build_model.PARAMS["vendor_a_region"]:
                    assert gap == 6
                else:
                    assert abs(gap) <= tol, (quarter, role, region, gap)
    for quarter in build_model.QUARTERS[1:]:
        de_gap = sum(supply_cells[("Data engineering", region, quarter)] - demand_cells[("Data engineering", region, quarter)] for region in regions)
        assert de_gap == 0, (quarter, de_gap)
    q1_de_gap = sum(supply_cells[("Data engineering", region, "Q1 2026")] - demand_cells[("Data engineering", region, "Q1 2026")] for region in regions)
    assert q1_de_gap == -14
    for quarter in build_model.QUARTERS:
        demand = sum(row[3] for row in demand1 if row[2] == quarter)
        supply = sum(row[4] for row in supply1 if row[3] == quarter)
        assert demand <= supply, (quarter, demand, supply)
        if quarter.startswith("Q") and quarter.endswith("2028"):
            assert demand == supply
    assert max(sum(row[3] for row in demand1 if row[2] == q) for q in build_model.QUARTERS) == 186


def test_requested_portfolio_and_decision_labels():
    wb = load_workbook(ROOT / "out" / "delivery_review_model.xlsx", data_only=True, read_only=True)
    p = wb["PORTFOLIO"]
    assert p["I10"].value == 0.9
    assert "Year 2-3 spend conditional on 31 Dec 2026 re-entry gate (D5)" in p["O8"].value
    d = wb["DECISIONS"]
    assert d["H8"].value == "mixed: in-scenario actual (simulated) to 31 Oct 2026; forecast (simulated) after"
    assert d["H9"].value == "mixed: in-scenario actual (simulated) to 31 Oct 2026; forecast (simulated) after"
    amber = [row for row in d.iter_rows(min_row=15, values_only=True) if row[2] == "Amber sustained target period"]
    assert len(amber) == 1
    assert amber[0][5] == "assumption (criterion, not met at cut-off)"


def test_stage_a1_named_checks_and_cell_tolerance():
    wb = load_workbook(ROOT / "out" / "delivery_review_model.xlsx", data_only=True, read_only=True)
    a = wb["ASSUMPTIONS"]
    assert a["K113"].value in (True,1)
    assert all(a.cell(r,11).value in (True,1) for r in range(92,100))
    cap = wb["CAPACITY"]
    assert [cap.cell(6,c).value for c in (3,4,5,6,7,8,9,10,11,12)] == [164,105,6,65,170,6,6,6,99,0]
    assert [cap.cell(r,12).value for r in range(6,17)] == [0]*11
    assert cap["C31"].value == -14 and cap["C32"].value == -6 and cap["C33"].value == 0
    assert "Vendor A outgoing QA" in cap["A3"].value
    assert build_model.PARAMS["cell_tolerance_fte"] == 2
    names = load_workbook(ROOT / "out" / "delivery_review_model.xlsx", data_only=False, read_only=False).defined_names
    for name in ("Q2_2026_Supply","Q2_2026_Demand","DualRun_FTE","Cell_Check_Result","Cell_Tolerance"):
        assert name in names


@pytest.mark.parametrize("direction,value,green,amber,expected", [
    ("higher",11,11,6,"Green"),("higher",6,11,6,"Amber"),("higher",5,11,6,"Red"),
    ("lower",0.015,0.015,0.02,"Green"),("lower",0.02,0.015,0.02,"Amber"),("lower",0.021,0.015,0.02,"Red"),
    ("lower",0,0,2,"Green"),("lower",2,0,2,"Amber"),("lower",2.1,0,2,"Red"),
    ("higher",0.4,0.4,0.3,"Green"),("higher",0.3,0.4,0.3,"Amber"),("higher",0.299,0.4,0.3,"Red"),
])
def test_delivery_health_boundaries_are_mutually_exclusive(direction,value,green,amber,expected):
    assert build_model.delivery_rag_band(direction,value,green,amber)==expected


def test_benefits_funnel_and_delivery_health():
    path=ROOT / "out" / "delivery_review_model.xlsx"
    values=load_workbook(path,data_only=True,read_only=True)
    formulas=load_workbook(path,data_only=False,read_only=True)
    b=values["BENEFITS"]
    assert [b.cell(r,1).value for r in range(5,11)]==["B1","B2","B3","B4","B5","B6"]
    assert [b.cell(r,12).value for r in range(5,11)]==[
        "in-scenario actual (simulated)","forecast (simulated)","forecast (simulated)",
        "forecast (simulated)","in-scenario actual (simulated)","forecast (simulated)"]
    assert b["J5"].value==4 and b["J9"].value==1.6
    assert formulas["BENEFITS"]["J5"].value.startswith("=")
    assert formulas["BENEFITS"]["J9"].value=="=VENDORS!B30"
    assert [b.cell(r,1).value for r in range(26,32)]==["H1","H2","H3","H4","H5","H6"]
    assert [b.cell(r,7).value for r in range(26,32)]==["Green","Green","Amber","Green","Green","Amber"]
    assert all(formulas["BENEFITS"].cell(r,3).value.startswith("=") for r in range(26,32))
    assert b["B18"].value==330 and b["E18"].value==107
    assert abs(b["H18"].value-107/330)<1e-12
    assert abs(b["G17"].value-18/52)<1e-12
    p=values["PLAN"]
    on_time=[p.cell(r,11).value for r in range(5,19)]
    assert sum(value is True for value in on_time)==6
    assert all(value is None for value in on_time if value is not True)
    a=values["ASSUMPTIONS"]
    assert all(a.cell(r,11).value in (True,1) for r in range(107,112))


def test_plan_cadence_and_monthly_measures_are_formula_linked():
    path=ROOT / "out" / "delivery_review_model.xlsx"
    values=load_workbook(path,data_only=True,read_only=True)
    formulas=load_workbook(path,data_only=False,read_only=True)
    p=values["PLAN"]
    assert [p.cell(r,1).value for r in range(58,61)]==[
        "Weekly delivery review","Vendor performance forum","SteerCo"]
    assert [p.cell(r,2).value for r in range(58,61)]==["Weekly","Monthly","Fortnightly"]
    assert p["B64"].value==11 and p["C64"].value==11
    assert abs(p["B65"].value-107/330)<1e-12 and p["C65"].value==0.4
    assert abs(p["B66"].value-1.8564)<1e-12 and p["C66"].value==2.04
    assert all(formulas["PLAN"][f"{col}{r}"].value.startswith("=") for r in range(64,67) for col in ("B","C"))
    assert p["B68"].value==build_model.KAKALI_TEXT["PAGE_TEXT"]["cadence_rationale"]


def test_portfolio_roadmap_collision_and_replan_inputs():
    wb=load_workbook(ROOT / "out" / "delivery_review_model.xlsx",data_only=True,read_only=True)
    p=wb["PORTFOLIO"]
    assert [p.cell(r,2).value for r in range(34,42)]==[
        "Q2 2026","3","6","Data engineering",6,6,"Q3 2026",4]
    assert all(p.cell(r,3).value==build_model.PARAMS["roadmap_transition"]["basis"] for r in range(34,42))


@pytest.mark.parametrize("dimension,value,expected", [
    ("sla",95,"Green"),("sla",90,"Amber"),("sla",89.9,"Red"),
    ("quality",1.5,"Green"),("quality",2.0,"Amber"),("quality",2.01,"Red"),
    ("delivery",90,"Green"),("delivery",80,"Amber"),("delivery",79.9,"Red"),
    ("cost",5,"Green"),("cost",10,"Amber"),("cost",10.1,"Red"),
])
def test_rag_boundaries_have_one_band(dimension,value,expected):
    threshold=build_model.PARAMS["vendor_scorecard"]["thresholds"][dimension]
    assert build_model.rag_band(dimension,value,threshold)==expected


def test_vendor_and_plan_outputs():
    wb=load_workbook(ROOT / "out" / "delivery_review_model.xlsx",data_only=True,read_only=True)
    v=wb["VENDORS"]
    assert [v.cell(r,11).value for r in range(5,10)]==["Green","Amber","Green","Green","Green"]
    assert v["J20"].value==54
    assert v["J24"].value=="not met at 31 Oct 2026"
    assert [v.cell(r,5).value for r in range(21,31)]==["in-scenario actual (simulated)"]*10
    assert [v.cell(r,6).value for r in range(34,38)]==["met"]*4
    assert [v.cell(r,8).value for r in range(34,38)]==["in-scenario actual (simulated)"]*4
    assert [v.cell(r,4).value for r in range(41,47)]==["in-scenario actual (simulated)"]*6
    assert [v.cell(r,11).value for r in range(49,52)]==["in-scenario actual (simulated)"]*3
    c=wb["COMMERCIAL"]
    assert [round(c.cell(71,col).value,1) for col in (2,3,4,5)]==[10.6,12.4,6.1,29.1]
    assert round(c["B62"].value)==round(c["C62"].value)==round(c["D62"].value)==480
    assert sum(1 for r in range(66,71) if c.cell(r,6).value>0.25)==3
    p=wb["PLAN"]
    assert p["B23"].value==11
    assert [p.cell(r,1).value for r in range(5,19)]==[f"M{i:02d}" for i in range(1,15)]
    assert [p.cell(r,1).value for r in range(27,33)]==["Funding release, tranche 2","Scope change above 10% of a workstream","Vendor award or re-compete","Go-live readiness","Regulatory commitment date","Benefit re-baseline"]


def test_vendor_a_q1_qa_baseline_and_commercial_months():
    wb=load_workbook(ROOT / "out" / "delivery_review_model.xlsx",data_only=True,read_only=True)
    cap=wb["CAPACITY"]
    demand_start=50
    demand_end=demand_start+len(build_model.capacity_inputs()[0])-1
    supply_start=demand_end+4
    supply_end=supply_start+len(build_model.capacity_inputs()[1])-1
    qa_region=build_model.PARAMS["vendor_a_region"]
    demand_rows=cap.iter_rows(min_row=demand_start,max_row=demand_end,min_col=1,max_col=4,values_only=True)
    supply_rows=cap.iter_rows(min_row=supply_start,max_row=supply_end,min_col=1,max_col=6,values_only=True)
    q1_qa_demand=sum(row[3] for row in demand_rows if row[0]=="QA" and row[1]==qa_region and row[2]=="Q1 2026")
    supply_rows=list(supply_rows)
    q1_vendor_a_qa=sum(row[4] for row in supply_rows if row[0]=="QA" and row[1]=="Vendor A" and row[2]==qa_region
                       and row[3]=="Q1 2026" and row[5]=="Baseline")
    q1_total_qa_supply=sum(row[4] for row in supply_rows if row[0]=="QA" and row[2]==qa_region
                           and row[3]=="Q1 2026" and row[5]=="Baseline")
    q1_total_vendor=sum(row[4] for row in supply_rows if row[3]=="Q1 2026" and row[1]!="Internal" and row[5]=="Baseline")
    assert q1_vendor_a_qa==6
    assert q1_qa_demand==6
    assert q1_total_qa_supply==q1_qa_demand
    assert q1_total_vendor==99
    comm=wb["COMMERCIAL"]
    vendor_a_qa=next(r for r in range(5,55) if comm.cell(r,1).value=="Vendor A" and comm.cell(r,2).value=="QA")
    assert comm.cell(vendor_a_qa,5).value==18
    assert comm["A60"].value==build_model.PARAMS["commercial"]["rate_scaling"]["label"]
    assert [round(comm.cell(60,c).value,8) for c in (2,3,4)]==[
        round(comm.cell(61,c).value/comm.cell(59,c).value,8) for c in (2,3,4)
    ]


def test_firm_name_and_history_employer_scan():
    root = ROOT
    terms = blocked_tokens()
    for path in root.rglob("*"):
        if any(part in {".git", ".pytest_cache", "__pycache__", ".lo_profile", ".lo_recalc", ".b2deps", ".venv"} for part in path.parts):
            continue
        if path.name and any(term in path.name.casefold() for term in terms):
            pytest.fail(f"Prohibited organization token in file name: {path.name}")
        if path.is_file() and path.suffix.lower() in {".py", ".yaml", ".md", ".txt", ".gitignore", ".toml"}:
            body = path.read_text(encoding="utf-8-sig", errors="ignore").casefold()
            assert not any(term in body for term in terms), f"Prohibited organization token in {path.name}"
    model = root / "out" / "delivery_review_model.xlsx"
    assert not any(term in value.casefold() for value in workbook_strings(model) for term in terms)
    import subprocess
    try:
        log = subprocess.run(["git", "log", "--all", "--format=%s"], cwd=root, text=True, capture_output=True, check=True).stdout.casefold()
    except subprocess.CalledProcessError:
        log = ""
    assert not any(term in log for term in terms)


def test_workbook_uses_only_the_qualified_actual_label():
    token = "".join(map(chr, [97,99,116,117,97,108]))
    allowed = f"in-scenario {token} (simulated)"
    for value in workbook_strings(ROOT / "out" / "delivery_review_model.xlsx"):
        remaining = value.casefold().replace(allowed.casefold(), "")
        import re
        assert re.search(r"" + token + r"", remaining) is None, value


