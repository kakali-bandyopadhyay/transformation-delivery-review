import json
import math
import random
import re
from datetime import datetime

import yaml
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parents[1]
PARAMS = json.loads((ROOT / "generator" / "params.yaml").read_text(encoding="utf-8-sig"))
KAKALI_TEXT = yaml.safe_load((ROOT / "generator" / "kakali_text.yaml").read_text(encoding="utf-8-sig"))
OUTPUT = ROOT / "out" / "delivery_review_model.xlsx"
YEARS = [2026, 2027, 2028]
QUARTERS = [f"Q{q} {year}" for year in YEARS for q in range(1, 5)]
FUNDED_ROWS = [x for x in PARAMS["portfolio"] if x["decision"].startswith("Fund")]

def split_integer(total, caps, rng):
    caps = [int(max(0, c)) for c in caps]
    order = list(range(len(caps)))
    rng.shuffle(order)
    result = [0] * len(caps)
    left = int(total)
    for n, idx in enumerate(order):
        rem_cap = sum(caps[j] for j in order[n + 1:])
        low = max(0, left - rem_cap)
        high = min(left, caps[idx])
        take = left if n == len(order) - 1 else rng.randint(low, high)
        result[idx] = take
        left -= take
    if left:
        for idx, cap in enumerate(caps):
            take = min(left, cap - result[idx])
            result[idx] += take
            left -= take
            if left == 0:
                break
    if left:
        raise ValueError("Seeded integer allocation did not fit")
    return result

def phase_funded_lines():
    rng = random.Random(PARAMS["seed"])
    row_units = [round(x["after_rescope_m"] * 10) for x in FUNDED_ROWS]
    remaining = [round(x * 10) for x in PARAMS["funding_by_year_m"]]
    phases = {}
    for ix, item in enumerate(FUNDED_ROWS):
        units = row_units[ix]
        if item["id"] == "5a":
            parts = [9, 0, 0]
        elif ix == len(FUNDED_ROWS) - 1:
            parts = remaining[:]
        else:
            order = list(range(3))
            rng.shuffle(order)
            parts = [0, 0, 0]
            left = units
            for n, col in enumerate(order):
                if n == len(order) - 1:
                    take = left
                else:
                    later = sum(remaining[k] for k in order[n + 1:])
                    low = max(0, left - later)
                    high = min(left, remaining[col])
                    take = rng.randint(low, high)
                parts[col] = take
                left -= take
        if sum(parts) != units or any(parts[k] > remaining[k] for k in range(3)):
            raise ValueError("Funding phasing does not reconcile")
        remaining = [remaining[k] - parts[k] for k in range(3)]
        phases[item["id"]] = [x / 10 for x in parts]
    if remaining != [0, 0, 0]:
        raise ValueError(f"Year phasing remainder: {remaining}")
    return phases

def split_regions(n, rng):
    return split_integer(n, [n, n, n, n], rng)

def weighted_integer(total, weights):
    names=list(weights)
    vals=[max(0,float(weights[name])) for name in names]
    denom=sum(vals)
    if denom<=0:
        raise ValueError("Weighted allocation has no positive weight")
    exact=[total*v/denom for v in vals]
    out=[math.floor(x) for x in exact]
    left=int(total)-sum(out)
    order=sorted(range(len(names)),key=lambda i:(exact[i]-out[i],-i),reverse=True)
    for i in order[:left]:
        out[i]+=1
    return dict(zip(names,out))

def rag_band(dimension, value, thresholds):
    if dimension == "sla":
        return "Green" if value >= thresholds["green_min"] else ("Amber" if value >= thresholds["amber_min"] else "Red")
    if dimension == "quality":
        return "Green" if value <= thresholds["green_max"] else ("Amber" if value <= thresholds["amber_max"] else "Red")
    if dimension == "delivery":
        return "Green" if value >= thresholds["green_min"] else ("Amber" if value >= thresholds["amber_min"] else "Red")
    if dimension == "cost":
        return "Green" if abs(value) <= thresholds["green_abs_max"] else ("Amber" if abs(value) <= thresholds["amber_abs_max"] else "Red")
    raise ValueError(f"Unknown RAG dimension: {dimension}")

def delivery_rag_band(direction, value, green_threshold, amber_threshold):
    if direction == "higher":
        return "Green" if value >= green_threshold else ("Amber" if value >= amber_threshold else "Red")
    if direction == "lower":
        return "Green" if value <= green_threshold else ("Amber" if value <= amber_threshold else "Red")
    raise ValueError(f"Unknown delivery-health direction: {direction}")

def capacity_inputs():
    rng = random.Random(PARAMS["seed"] + 1)
    roles = PARAMS["capacity"]["role_taxonomy"]["names"]
    regions = PARAMS["capacity"]["regions"]
    vendors = PARAMS["capacity"]["vendors"]
    vendor_counts = PARAMS["vendor"]["grid_fte_by_quarter"]
    other = split_integer(62, [62] * 6, rng)
    weights = dict(zip(roles, [62, 28, 14, 20] + other))
    peak_demand = dict(weights)
    demand=[]
    supply=[]
    q1_de_gap=PARAMS["q1_2026_de_shortfall_by_region"]
    de_region=PARAMS["de_demand_region"]["region"]
    if set(q1_de_gap)!=set(regions) or sum(q1_de_gap.values())!=-14 or any(v>0 for v in q1_de_gap.values()):
        raise ValueError("Q1 DE regional exception must cover all regions, be non-positive and sum to -14")
    if de_region not in regions:
        raise ValueError("Synthetic DE demand region must be one of the configured regions")
    if PARAMS["q2_2026_demand_increment"]["region"] != de_region:
        raise ValueError("The Q2 DE demand increment must align with the synthetic DE delivery region")
    vendor_a_region=PARAMS["vendor_a_region"]
    if vendor_a_region not in regions:
        raise ValueError("Vendor A synthetic region must be one of the configured regions")
    for qi,quarter in enumerate(QUARTERS):
        base_vendor=int(vendor_counts[qi])
        base_total=base_vendor+65
        transition_fte=int(PARAMS["overlap"]["fte"]) if quarter=="Q2 2026" else 0
        if quarter=="Q1 2026":
            target=base_total
            role_demand=weighted_integer(target,weights)
        elif quarter=="Q2 2026":
            target=int(PARAMS["q2_2026_pre_a1_demand_total"])
            role_demand=weighted_integer(target,weights)
        elif quarter=="Q3 2027":
            role_demand=peak_demand.copy()
            target=sum(role_demand.values())
        elif quarter.endswith("2028"):
            target=base_total
            role_demand=weighted_integer(target,weights)
        else:
            target=base_total-rng.randint(0,7)
            role_demand=weighted_integer(target,weights)
        regional_demand={}
        for role in roles:
            alloc=split_regions(role_demand[role],rng)
            regional_demand[role]={region:qty for region,qty in zip(regions,alloc)}
        if role_demand.get("Data engineering") is not None:
            # Keep the synthetic DE role concentrated in the golden-source
            # delivery region so the six-person Vendor D team fits its cells.
            regional_demand["Data engineering"]={region:(role_demand["Data engineering"] if region==de_region else 0) for region in regions}
        if quarter=="Q1 2026":
            # Reserve six baseline Vendor A QA FTE in London for Jan–Mar.
            # Reallocate QA demand from other regions so total demand is unchanged.
            qa_total=sum(regional_demand["QA"].values())
            if qa_total<6:
                raise ValueError("Q1 QA demand cannot support the six-FTE Vendor A London baseline")
            qa_other_regions=[region for region in regions if region!=vendor_a_region]
            qa_other_demand=split_integer(qa_total-6,[100]*len(qa_other_regions),rng)
            regional_demand["QA"]={region:(6 if region==vendor_a_region else 0) for region in regions}
            for region,qty in zip(qa_other_regions,qa_other_demand):
                regional_demand["QA"][region]=qty
        if quarter=="Q2 2026":
            increment=PARAMS["q2_2026_demand_increment"]
            regional_demand["Data engineering"][increment["region"]]+=int(increment["fte"])
        for role in roles:
            for region in regions:
                demand.append([role,region,quarter,regional_demand[role][region]])
        regional_supply={role:dict(regional_demand[role]) for role in roles}
        if quarter=="Q1 2026":
            for region,gap in q1_de_gap.items():
                regional_supply["Data engineering"][region]+=gap
            compensation=-sum(q1_de_gap.values())
            eligible=[(role,region) for role in roles if role!="Data engineering" for region in regions
                      if not (role=="QA" and region==vendor_a_region)]
            rng.shuffle(eligible)
            for role,region in eligible[:compensation]:
                regional_supply[role][region]+=1
        elif quarter=="Q3 2027":
            region=next(r for r in regions if regional_demand["QA"][r]>0)
            regional_supply["QA"][region]-=1
            regional_supply["Business analysis"][region]+=1
        elif quarter!="Q2 2026":
            slack=base_total-sum(sum(x.values()) for x in regional_demand.values())
            eligible=[(role,region) for role in roles if role!="Data engineering" for region in regions]
            rng.shuffle(eligible)
            for role,region in eligible[:slack]:
                regional_supply[role][region]+=1
        if any(qty<0 for role in roles for qty in regional_supply[role].values()):
            raise ValueError(f"Negative role-region supply in {quarter}")
        cell_totals={(role,region):regional_supply[role][region] for role in roles for region in regions}
        if sum(cell_totals.values())!=base_total:
            raise ValueError(f"Base supply does not tie in {quarter}")
        if quarter=="Q1 2026":
            if sum(regional_supply["Data engineering"][r]-regional_demand["Data engineering"][r] for r in regions)!=-14:
                raise ValueError("Q1 DE role total does not tie to the named -14 exception")
        if quarter=="Q2 2026" and sum(sum(x.values()) for x in regional_demand.values())!=164:
            raise ValueError("Q2 demand after A.1 increment must be 164")
        if quarter=="Q3 2027":
            for role in roles:
                other_roles=set(roles[4:])
                expected={"Data engineering":62,"Integration":28,"QA":13,"Business analysis":21}.get(role,62 if role in other_roles else None)
                if role=="Data engineering":
                    assert sum(regional_supply[role].values())==62
                elif role=="Integration":
                    assert sum(regional_supply[role].values())==28
                elif role=="QA":
                    assert sum(regional_supply[role].values())==13
                elif role=="Business analysis":
                    assert sum(regional_supply[role].values())==21
        profile_key="Q1 2026" if quarter=="Q1 2026" else ("Q2 2026" if quarter=="Q2 2026" else "post_exit")
        vendor_weights=PARAMS["vendor"]["share_profiles"][profile_key]
        vendor_quota=weighted_integer(base_vendor,vendor_weights)
        remaining=dict(cell_totals)
        vendor_cells={vendor:{cell:0 for cell in cell_totals} for vendor in vendors}
        if quarter!="Q1 2026":
            cell=("Data engineering",de_region)
            reserve=int(PARAMS["vendor_d_de_baseline_fte"])
            if remaining[cell]<reserve or vendor_quota["Vendor D"]<reserve:
                raise ValueError(f"Vendor D DE baseline cannot fit in {quarter}")
            vendor_cells["Vendor D"][cell]=reserve
            remaining[cell]-=reserve
            vendor_quota["Vendor D"]-=reserve
        else:
            # These six are Q1 baseline supply; Q2 transition overlap remains a separate row.
            reserved=6
            cell=("QA",vendor_a_region)
            if remaining[cell]<reserved or vendor_quota["Vendor A"]<reserved:
                raise ValueError("Vendor A London QA baseline cannot fit in Q1 2026")
            vendor_cells["Vendor A"][cell]=reserved
            remaining[cell]-=reserved
            vendor_quota["Vendor A"]-=reserved
        for vendor in vendors:
            quota=vendor_quota[vendor]
            caps=[]
            cell_order=list(cell_totals)
            for role,region in cell_order:
                cap=remaining[(role,region)]
                if quarter=="Q1 2026" and vendor=="Vendor D" and role=="Data engineering":
                    cap=0
                caps.append(cap)
            assigned=split_integer(quota,caps,rng)
            for cell,qty in zip(cell_order,assigned):
                vendor_cells[vendor][cell]+=qty
                remaining[cell]-=qty
        if sum(remaining.values())!=65 or any(qty<0 for qty in remaining.values()):
            raise ValueError(f"Internal supply does not reconcile in {quarter}")
        for role,region in cell_totals:
            for vendor in vendors:
                supply.append([role,vendor,region,quarter,vendor_cells[vendor][(role,region)],"Baseline"])
            supply.append([role,"Internal",region,quarter,remaining[(role,region)],"Baseline"])
        if transition_fte:
            supply.append(["QA","Vendor A",vendor_a_region,quarter,transition_fte,"transition overlap (contingency)"])
    return demand,supply,vendor_counts

def title(ws, text, widths, freeze):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws["A1"] = text
    ws["A1"].font = Font(name="Aptos Display", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="15263D")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(widths))
    ws["A2"] = "Simulated data – not observed results"
    ws["A2"].font = Font(italic=True, size=10, color="68778A")
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = freeze

def headers(ws, row, vals):
    for idx, value in enumerate(vals, 1):
        c = ws.cell(row, idx, value)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="31506E")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 34

def add_table(ws, name, ref):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(t)

def build():
    phases = phase_funded_lines()
    demand_data, supply_data, vendor_quarters = capacity_inputs()
    wb = Workbook()
    asm = wb.active
    asm.title = "ASSUMPTIONS"
    port = wb.create_sheet("PORTFOLIO")
    cap = wb.create_sheet("CAPACITY")
    dec = wb.create_sheet("DECISIONS")
    vend = wb.create_sheet("VENDORS")
    comm = wb.create_sheet("COMMERCIAL")
    plan = wb.create_sheet("PLAN")
    benefits = wb.create_sheet("BENEFITS")
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"

    title(asm, "Stage A | Assumptions and checks", [12, 38, 22, 20, 58, 34, 27, 4, 14, 62, 14], "A5")
    headers(asm, 4, ["ID", "Parameter", "Value", "Unit", "Basis", "Status", "Used in"])
    a = {}
    records = []
    def put(key, name, val, unit, basis, status="assumption", used="All tabs", formula=False):
        records.append((key,name,val,unit,basis,status,used,formula))
    put("seed","Generator seed",PARAMS["seed"],"integer","Fixed seed specified in Stage A brief.","assumption","PORTFOLIO; CAPACITY")
    put("as_at","As-at date",datetime.fromisoformat(PARAMS["as_at"]),"date","Spec §4; scenario cut-off.","in-scenario actual (simulated)")
    put("funded_total","Funded work, three years","=PORTFOLIO!C22","$m","Formula: sum of eight funded lines.")
    put("cont_pct","Contingency rate",0.05,"fraction","Confirmed §19: 5% of funded work.")
    put("cont","Contingency","=C7*C8","$m","Formula: funded work × contingency rate.")
    put("envelope","Scenario envelope","=C7+C9","$m","Formula: funded work + contingency.")
    put("envelope_label","Envelope approval label",PARAMS["envelope_label"],"text","Confirmed §19 choice; generic scenario assumes approval.")
    for ix, value in enumerate(PARAMS["funding_by_year_m"],1):
        put(f"fund_y{ix}",f"Funded work, Year {ix}",value,"$m","Spec §10.1; confirmed funding profile.","forecast (simulated)","PORTFOLIO")
    for ix, value in enumerate(PARAMS["other_costs_by_year_m"],1):
        put(f"other_y{ix}",f"Other costs, Year {ix}",value,"$m","Software, cloud, incremental backfill and non-vendor costs; spec §10.1.","forecast (simulated)")
    put("rate_month","Vendor cost per FTE-month",8800,"USD / FTE-month","220 days × $480 blended daily rate ÷ 12.")
    put("days_year","Working days per year",220,"days","Spec §10.1 vendor cost basis.")
    put("rate_blend","Blended vendor day rate",480,"USD / day","Spec §10.1 vendor cost basis.")
    put("rate_special","Specialist day rate",600,"USD / day","Spec §10.3 overlap basis.")
    put("y1_fte_a","Year 1 average vendor FTE, Jan-Oct",99,"FTE","Spec §10.1; 10 months to cut-off.","in-scenario actual (simulated)")
    put("y1_month_a","Year 1 months, Jan-Oct",10,"months","Spec §10.1; January-October.","in-scenario actual (simulated)")
    put("y1_fte_f","Year 1 average vendor FTE, Nov-Dec",110,"FTE","Spec §10.1; November-December run-rate.","forecast (simulated)")
    put("y1_month_f","Year 1 months, Nov-Dec",2,"months","Spec §10.1; November-December.","forecast (simulated)")
    for i,v in enumerate(PARAMS["vendor"]["year2_quarter_average_fte"],1):
        put(f"y2q{i}",f"Year 2 Q{i} average vendor FTE",v,"FTE",f"Spec §10.1; Q{i} 2027.","forecast (simulated)","ASSUMPTIONS; CAPACITY")
    for i,v in enumerate(PARAMS["vendor"]["year3_quarter_average_fte"],1):
        put(f"y3q{i}",f"Year 3 Q{i} average vendor FTE",v,"FTE",f"Spec §10.1; Q{i} 2028.","forecast (simulated)","ASSUMPTIONS; CAPACITY")
    put("v_y1_a","Vendor cost, Year 1 Jan-Oct","=C22*C23*C18/1000000","$m","Formula: FTE × months × FTE-month cost.","in-scenario actual (simulated)")
    put("v_y1_f","Vendor cost, Year 1 Nov-Dec","=C24*C25*C18/1000000","$m","Formula: FTE × months × FTE-month cost.","forecast (simulated)")
    put("v_y1","Vendor cost, Year 1","=C28+C29","$m","Formula: Jan-Oct plus Nov-Dec.","forecast (simulated)")
    put("v_y2_avg","Year 2 average vendor FTE","=AVERAGE(C26:C29)","FTE","Formula: average of four quarterly forecasts.","forecast (simulated)")
    put("v_y2","Vendor cost, Year 2","=C31*12*C18/1000000","$m","Formula: average vendor FTE × 12 months × rate.","forecast (simulated)")
    put("v_y3_avg","Year 3 average vendor FTE","=AVERAGE(C30:C33)","FTE","Formula: average of four quarterly forecasts.","forecast (simulated)")
    put("v_y3","Vendor cost, Year 3","=C33*12*C18/1000000","$m","Formula: average vendor FTE × 12 months × rate.","forecast (simulated)")
    for ix, value in enumerate([10.6,12.4,6.1],1):
        put(f"vendor_y{ix}",f"Vendor spend target, Year {ix}",value,"$m","Spec §10.1 rounded vendor spend; reconciled to FTE-month calculation.","forecast (simulated)")
    put("week","Working week","Five-day working week","text","Spec §10.3; weekends excluded.")
    put("holidays","Public holidays","Ignored","text","Spec §10.3; public holidays ignored in working-day count.")
    overlap = PARAMS["overlap"]
    for key,name in [("dual_start","Dual-run start"),("dual_end","Dual-run end"),("notice_start","Vendor A notice served"),("notice_end","Vendor A notice expires"),("paid_start","Paid, not delivering start"),("paid_end","Paid, not delivering end")]:
        put(key,name,datetime.fromisoformat(overlap[key.replace("dual_start","dual_run_start").replace("dual_end","dual_run_end").replace("notice_start","notice_start").replace("notice_end","notice_end").replace("paid_start","paid_not_delivering_start").replace("paid_end","paid_not_delivering_end")]),"date","Spec §10.3; A-to-D transition dates.","in-scenario actual (simulated)")
    put("dual_days","Dual-run working days","=NETWORKDAYS(C42,C43)","days","Formula: NETWORKDAYS; five-day week, holidays ignored.","in-scenario actual (simulated)")
    put("paid_days","Paid, not delivering working days","=NETWORKDAYS(C46,C47)","days","Formula: NETWORKDAYS; five-day week, holidays ignored.","in-scenario actual (simulated)")
    put("overlap_fte","Overlap FTE",6,"FTE","Spec §10.3; six specialist FTE.")
    put("overlap_cost","A-to-D overlap cost","=(C50+C51)*C52*C21","USD","Formula: (dual-run + paid notice days) × FTE × specialist day rate.","in-scenario actual (simulated)")
    put("unalloc","Unallocated contingency","=C9*1000000-C53","USD","Formula: contingency less overlap cost.","forecast (simulated)")
    capparams = PARAMS["capacity"]
    put("peak_pre","Peak demand before decisions",205,"FTE","Spec §10.2; Q3 2027.","forecast (simulated)","CAPACITY")
    for key,name in [("Client portal","Portal requisitions removed"),("Finance close","Finance-close requisitions removed"),("Cloud migration wave 2","Cloud-wave requisitions removed")]:
        put("req_portal" if key=="Client portal" else ("req_finance" if key=="Finance close" else "req_cloud"),name,capparams["deferred_requisitions"][key],"FTE","Spec §10.2; unfilled requisitions.")
    put("req_total","Deferred requisitions removed","=SUM(C56:C58)","FTE","Formula: three deferred initiative requisition counts.","forecast (simulated)","CAPACITY")
    put("peak_post","Peak demand after decisions","=C55-C59","FTE","Formula: pre-decision demand less deferred requisitions.","forecast (simulated)","CAPACITY")
    put("vendor_peak","Vendor supply at peak",121,"FTE","Spec §10.2; Q3 2027 forecast.","forecast (simulated)","CAPACITY")
    put("internal_peak","Internal supply at peak",65,"FTE","Spec §10.2; charged to BU run budgets, outside envelope.","forecast (simulated)","CAPACITY")
    put("supply_peak","Total supply at peak","=C61+C62","FTE","Formula: vendor + internal.","forecast (simulated)","CAPACITY")
    put("d5_live","D5 live use cases",2,"count","Confirmed D5 split.")
    put("d5_rescope","D5 re-scoped use cases",5,"count","Confirmed D5 split.")
    put("d5_retired","D5 retired use cases",2,"count","Confirmed D5 split.")
    put("d5_gate","D5 re-entry gate",datetime(2026,12,31),"date","Spec §8; end Q4 2026.","forecast (simulated)")
    put("workday_basis","Working-day assumption","Five-day week; public holidays ignored","text","Spec §10.3.")
    put("capacity_note","Quarterly grid basis", "First-day-of-quarter position; Q4 2026 base vendor supply is 99 on 1 Oct while Nov-Dec cost average is 110.","text","Stage B1 quarterly grid convention; §10.1 averages feed COMMERCIAL.","forecast (simulated)","CAPACITY")
    put("cell_tolerance","Role-region cell tolerance",PARAMS["cell_tolerance_fte"],"FTE",PARAMS["input_basis"]["cell_tolerance_fte"],"assumption","CAPACITY")
    put("vendor_a_region","Vendor A QA region (synthetic)",PARAMS["vendor_a_region"],"text",PARAMS["input_basis"]["vendor_a_region"],"assumption","CAPACITY")
    put("de_demand_region","Synthetic DE demand region",PARAMS["de_demand_region"]["region"],"text",PARAMS["de_demand_region"]["basis"],"assumption","CAPACITY")
    put("q1_de_distribution","Q1 2026 DE shortfall by region",json.dumps(PARAMS["q1_2026_de_shortfall_by_region"],ensure_ascii=False),"text",PARAMS["input_basis"]["q1_2026_de_shortfall_by_region"],"assumption","CAPACITY")
    put("q2_de_increment","Q2 2026 incremental DE demand",PARAMS["q2_2026_demand_increment"]["fte"],"FTE",PARAMS["input_basis"]["q2_2026_demand_increment"],"in-scenario actual (simulated)","CAPACITY")
    put("m1_de_fte","M1 DE training supply",8,"FTE","Spec §10.3; M1 training completes 31 Mar 2026.","in-scenario actual (simulated)","CAPACITY")
    put("vendor_d_de_fte","Vendor D DE baseline supply",PARAMS["vendor_d_de_baseline_fte"],"FTE",PARAMS["input_basis"]["vendor_d_de_baseline_fte"],"in-scenario actual (simulated)","CAPACITY")
    for rownum, record in enumerate(records,5):
        key,name,value,unit,basis,status,used,formula = record
        a[key] = rownum
        vals = [key,name,value,unit,basis,status,used]
        for col,val in enumerate(vals,1):
            cell=asm.cell(rownum,col,val)
            cell.alignment=Alignment(vertical="top",wrap_text=col in (2,3,5,6,7))
            if col==3 and formula:
                cell.font=Font(color="1F5EAA")
            elif col==3:
                cell.fill=PatternFill("solid",fgColor="FFF4D6")
            if col==3:
                cell.number_format='$#,##0.00' if unit=="$m" else ('$#,##0' if unit=="USD" else ("0.00%" if unit=="fraction" else ("dd mmm yyyy" if unit=="date" else "0")))
        if unit=="text":
            asm.row_dimensions[rownum].height=30
    def set_assumption_formula(key, formula):
        cell = asm.cell(a[key], 3)
        cell.value = formula
        cell.font = Font(color="1F5EAA")
    set_assumption_formula("rate_month", f"=C{a['days_year']}*C{a['rate_blend']}/12")
    set_assumption_formula("funded_total", "=PORTFOLIO!C22")
    set_assumption_formula("cont", f"=C{a['funded_total']}*C{a['cont_pct']}")
    set_assumption_formula("envelope", f"=C{a['funded_total']}+C{a['cont']}")
    set_assumption_formula("v_y1_a", f"=C{a['y1_fte_a']}*C{a['y1_month_a']}*C{a['rate_month']}/1000000")
    set_assumption_formula("v_y1_f", f"=C{a['y1_fte_f']}*C{a['y1_month_f']}*C{a['rate_month']}/1000000")
    set_assumption_formula("v_y1", f"=C{a['v_y1_a']}+C{a['v_y1_f']}")
    set_assumption_formula("v_y2_avg", f"=AVERAGE(C{a['y2q1']}:C{a['y2q4']})")
    set_assumption_formula("v_y2", f"=C{a['v_y2_avg']}*12*C{a['rate_month']}/1000000")
    set_assumption_formula("v_y3_avg", f"=AVERAGE(C{a['y3q1']}:C{a['y3q4']})")
    set_assumption_formula("v_y3", f"=C{a['v_y3_avg']}*12*C{a['rate_month']}/1000000")
    set_assumption_formula("dual_days", f"=NETWORKDAYS(C{a['dual_start']},C{a['dual_end']})")
    set_assumption_formula("paid_days", f"=NETWORKDAYS(C{a['paid_start']},C{a['paid_end']})")
    set_assumption_formula("overlap_cost", f"=(C{a['dual_days']}+C{a['paid_days']})*C{a['overlap_fte']}*C{a['rate_special']}")
    set_assumption_formula("unalloc", f"=C{a['cont']}*1000000-C{a['overlap_cost']}")
    set_assumption_formula("req_total", f"=SUM(C{a['req_portal']}:C{a['req_cloud']})")
    set_assumption_formula("peak_post", f"=C{a['peak_pre']}-C{a['req_total']}")
    set_assumption_formula("supply_peak", f"=C{a['vendor_peak']}+C{a['internal_peak']}")
    # PORTFOLIO
    title(port,"Stage A | Portfolio phasing",[9,36,22,18,12,14,18,14,12,12,12,24,14,30,54,20,20,20],"A5")
    headers(port,4,["#","Initiative","BU","Region","Value","Feasibility","Before re-scope $m","Moved $m","Y1 $m","Y2 $m","Y3 $m","Requested, outside envelope $m","Total $m","Product owner (role)","Decision","Re-entry trigger","Stop condition","User outcome"])
    deferred_or_stopped={"5","7","10","11","12"}
    for rownum,item in enumerate(PARAMS["portfolio"],5):
        funded=item["decision"].startswith("Fund")
        years=phases.get(item["id"],[0.0,0.0,0.0])
        outside=0.0 if funded else item["after_rescope_m"]
        vals=[item["id"],item["name"],item["bu"],item["region"],item["value"],item["feasibility"],item.get("before_rescope_m"),item.get("moved_m"),years[0],years[1],years[2],outside,f"=SUM(I{rownum}:K{rownum})+L{rownum}",item["owner"],item["decision"],"[Kakali]" if item["id"] in deferred_or_stopped else "", "[Kakali]","[Kakali]"]
        for col,val in enumerate(vals,1):
            cell=port.cell(rownum,col,val)
            cell.alignment=Alignment(vertical="top",wrap_text=True)
            if col not in (13,) and val is not None:
                cell.fill=PatternFill("solid",fgColor="FFF4D6")
            if col in (7,8,9,10,11,12,13):
                cell.number_format='$#,##0.00'
            if col==13:
                cell.font=Font(color="1F5EAA")
    add_table(port,"PortfolioLines","A4:R17")
    port["A20"]="Portfolio summary"
    port["A20"].font=Font(bold=True,size=12,color="15263D")
    summary=[
        ("Total request, $m","=SUM(M5:M17)"),
        ("Funded, $m",'=SUMIF(O5:O17,"Fund*",M5:M17)'),
        ("Deferred, $m",'=SUMIF(O5:O17,"Defer*",L5:L17)'),
        ("Stopped, $m",'=SUMIF(O5:O17,"Stop*",L5:L17)'),
        ("D1 withdrawn, $m","=L9"),
        ("Funded Year 1, $m",'=SUMIF(O5:O17,"Fund*",I5:I17)'),
        ("Funded Year 2, $m",'=SUMIF(O5:O17,"Fund*",J5:J17)'),
        ("Funded Year 3, $m",'=SUMIF(O5:O17,"Fund*",K5:K17)')
    ]
    for rownum,(label,formula) in enumerate(summary,21):
        port.cell(rownum,2,label)
        port.cell(rownum,3,formula).number_format='$#,##0.00'
        port.cell(rownum,3).font=Font(color="1F5EAA")
    port["A18"]="Total"
    for col in range(9,14):
        letter=get_column_letter(col)
        port.cell(18,col,f"=SUM({letter}5:{letter}17)").number_format='$#,##0.00'
        port.cell(18,col).font=Font(bold=True,color="1F5EAA")
        port.cell(18,col).border=Border(top=Side(style="thin",color="31506E"))
    port["B30"]="Reconciliation note"
    port["C30"]="Line 1 $0.9m moved to 5a; D1 withdrawal is the full $4.8m replacement scope."
    port.merge_cells("C30:R30")
    roadmap=PARAMS["roadmap_transition"]
    port["A32"]="E1.2 | Q2 2026 collision and D4 replan"
    port["A32"].font=Font(bold=True,size=12,color="15263D")
    headers(port,33,["Roadmap field","Value","Basis"])
    roadmap_fields=[
        ("Collision quarter","collision_quarter"),
        ("Portfolio line needing engineers 1","line3"),
        ("Portfolio line needing engineers 2","line6"),
        ("Shared role","shared_role"),
        ("Shared data-engineering requirement, FTE","shared_requirement_fte"),
        ("Line 6 requirement before D4, FTE","line6_original_requirement_fte"),
        ("Line 6 re-planned quarter","line6_replanned_quarter"),
        ("Line 6 revised requirement after D4, FTE","line6_revised_requirement_fte")
    ]
    roadmap_transition_cells={}
    for rownum,(label,key) in enumerate(roadmap_fields,34):
        port.cell(rownum,1,label); port.cell(rownum,2,roadmap[key]); port.cell(rownum,3,roadmap["basis"])
        for col in range(1,4): port.cell(rownum,col).alignment=Alignment(wrap_text=True,vertical="top")
        roadmap_transition_cells[key]=f"PORTFOLIO!B{rownum}"
    add_table(port,"RoadmapCollision","A33:C41")

    # CAPACITY
    demand_data,supply_data,vendor_quarters=capacity_inputs()
    title(cap,"Stage A | Quarterly demand and supply",[24,12,18,20,22,18,18,14,18,24,20,16],"A5")
    cap["A3"]="Q2 2026 includes six Vendor A outgoing QA overlap FTE; Vendor D data engineers are baseline from 1 Apr. Q4 grid is 99 on 1 Oct; Nov-Dec cost average is 110."
    cap["A3"].font=Font(italic=True,color="68778A")
    headers(cap,4,["Quarter","Year","Demand FTE","Vendor supply FTE","QA transition overlap FTE","Internal supply FTE","Total supply FTE","Gap FTE","QA role gap","Vendor A QA region gap","Base vendor supply","DE role gap"])
    dhead=49
    dstart=dhead+1
    dend=dstart+len(demand_data)-1
    shead=dend+3
    sstart=shead+1
    send=sstart+len(supply_data)-1
    for rownum,q in enumerate(QUARTERS,5):
        cap.cell(rownum,1,q)
        cap.cell(rownum,2,int(q[-4:]))
        cap.cell(rownum,3,f"=SUMIFS(D{dstart}:D{dend},C{dstart}:C{dend},A{rownum})")
        cap.cell(rownum,4,f'=SUMIFS(CAPACITY!E{sstart}:E{send},CAPACITY!D{sstart}:D{send},A{rownum},CAPACITY!B{sstart}:B{send},"<>Internal")')
        cap.cell(rownum,5,f'=SUMIFS(CAPACITY!E{sstart}:E{send},CAPACITY!D{sstart}:D{send},A{rownum},CAPACITY!F{sstart}:F{send},"transition overlap (contingency)")')
        cap.cell(rownum,6,f'=SUMIFS(CAPACITY!E{sstart}:E{send},CAPACITY!D{sstart}:D{send},A{rownum},CAPACITY!B{sstart}:B{send},"Internal")')
        cap.cell(rownum,7,f"=D{rownum}+F{rownum}")
        cap.cell(rownum,8,f"=G{rownum}-C{rownum}")
        cap.cell(rownum,9,f'=SUMIFS(CAPACITY!E{sstart}:E{send},CAPACITY!D{sstart}:D{send},A{rownum},CAPACITY!A{sstart}:A{send},"QA")-SUMIFS(CAPACITY!D{dstart}:D{dend},CAPACITY!C{dstart}:C{dend},A{rownum},CAPACITY!A{dstart}:A{dend},"QA")')
        cap.cell(rownum,10,f'=SUMIFS(CAPACITY!E{sstart}:E{send},CAPACITY!D{sstart}:D{send},A{rownum},CAPACITY!A{sstart}:A{send},"QA",CAPACITY!C{sstart}:C{send},"{PARAMS["vendor_a_region"]}")-SUMIFS(CAPACITY!D{dstart}:D{dend},CAPACITY!C{dstart}:C{dend},A{rownum},CAPACITY!A{dstart}:A{dend},"QA",CAPACITY!B{dstart}:B{dend},"{PARAMS["vendor_a_region"]}")')
        cap.cell(rownum,11,f"=D{rownum}-E{rownum}")
        cap.cell(rownum,12,f'=SUMIFS(CAPACITY!E{sstart}:E{send},CAPACITY!D{sstart}:D{send},A{rownum},CAPACITY!A{sstart}:A{send},"Data engineering")-SUMIFS(CAPACITY!D{dstart}:D{dend},CAPACITY!C{dstart}:C{dend},A{rownum},CAPACITY!A{dstart}:A{dend},"Data engineering")')
        for col in range(3,13):
            cap.cell(rownum,col).font=Font(color="1F5EAA")
            cap.cell(rownum,col).number_format="#,##0"
    cap["A19"]="Role bridge at Q3 2027"
    cap["A19"].font=Font(bold=True,size=12,color="15263D")
    headers(cap,20,["Role","Demand","Supply before","Gap before","M1","M2","M3","M4","Supply after","Gap after"])
    for rownum,item in enumerate(capparams["role_bridge"],21):
        role=item["role"]
        cap.cell(rownum,1,role)
        if role=="Six other roles":
            other_roles=PARAMS["capacity"]["role_taxonomy"]["names"][4:]
            dformula="="+"+".join(f'SUMIFS(D{dstart}:D{dend},C{dstart}:C{dend},"Q3 2027",A{dstart}:A{dend},"{role_name}")' for role_name in other_roles)
        else:
            dformula=f'=SUMIFS(D{dstart}:D{dend},C{dstart}:C{dend},"Q3 2027",A{dstart}:A{dend},A{rownum})'
        cap.cell(rownum,2,dformula)
        cap.cell(rownum,3,item["supply_before"])
        cap.cell(rownum,4,f"=C{rownum}-B{rownum}")
        for col,key in zip(range(5,9),["M1","M2","M3","M4"]):
            cap.cell(rownum,col,item[key])
        cap.cell(rownum,9,f"=C{rownum}+SUM(E{rownum}:H{rownum})")
        cap.cell(rownum,10,f"=I{rownum}-B{rownum}")
        for col in (2,4,9,10):
            cap.cell(rownum,col).font=Font(color="1F5EAA")
        for col in range(2,11):
            cap.cell(rownum,col).number_format='#,##0;[Red](#,##0);-'
    cap["A26"]="Total"
    for col in range(2,11):
        letter=chr(64+col)
        cap.cell(26,col,f"=SUM({letter}21:{letter}25)")
        cap.cell(26,col).font=Font(color="1F5EAA")
        cap.cell(26,col).number_format='#,##0;[Red](#,##0);-'
    cap["A27"]="Named residual"
    cap["B27"]="QA -1 FTE and business analysis +1 FTE at peak; covered by flex hours."
    cap.merge_cells("B27:J27")
    cap["A29"]="Data engineering gap timeline"
    cap["A29"].font=Font(bold=True,size=12,color="15263D")
    headers(cap,30,["Milestone","Date","Gap FTE","Basis"])
    q1_de_gap_formula=f'=SUMIFS(E{sstart}:E{send},D{sstart}:D{send},"Q1 2026",A{sstart}:A{send},"Data engineering")-SUMIFS(D{dstart}:D{dend},C{dstart}:C{dend},"Q1 2026",A{dstart}:A{dend},"Data engineering")'
    timeline=[
        ["Before M1 (1 Jan 2026)",datetime(2026,1,1),q1_de_gap_formula,"Q1 first-day role total; synthetic regional allocation in params."],
        ["After M1 completes",datetime(2026,3,31),f"=C31+ASSUMPTIONS!C{a['m1_de_fte']}","Eight M1 data engineers land at quarter end."],
        ["After Vendor D starts",datetime(2026,4,1),f"=C32+ASSUMPTIONS!C{a['vendor_d_de_fte']}","Vendor D six data engineers are baseline supply from 1 April."]
    ]
    for rownum,vals in enumerate(timeline,31):
        for col,val in enumerate(vals,1):
            cap.cell(rownum,col,val)
            cap.cell(rownum,col).alignment=Alignment(wrap_text=True,vertical="top")
            if isinstance(val,str) and val.startswith("="):
                cap.cell(rownum,col).font=Font(color="1F5EAA")
        cap.cell(rownum,2).number_format="dd mmm yyyy"
    headers(cap,dhead,["Role","Region","Quarter","Demand FTE"])
    for rownum,row in enumerate(demand_data,dstart):
        for col,val in enumerate(row,1):
            cap.cell(rownum,col,val)
    headers(cap,shead,["Role","Supplier","Region","Quarter","Supply FTE","Supply category"])
    for rownum,row in enumerate(supply_data,sstart):
        for col,val in enumerate(row,1):
            cap.cell(rownum,col,val)
    cap.auto_filter.ref=f"A{dhead}:D{dend}"
    cap.freeze_panes="A5"
    recon_head=send+3
    recon_start=recon_head+1
    roles=PARAMS["capacity"]["role_taxonomy"]["names"]
    regions=PARAMS["capacity"]["regions"]
    headers(cap,recon_head,["Quarter","Role","Region","Demand FTE","Supply FTE","Gap FTE","Named exception","Tolerance result"])
    recon_keys=[]
    recon_row=recon_start
    for quarter in QUARTERS:
        for role in roles:
            for region in regions:
                exception="E1" if quarter=="Q1 2026" and role=="Data engineering" else ("E2" if quarter=="Q2 2026" and role=="QA" and region==PARAMS["vendor_a_region"] else "")
                cap.cell(recon_row,1,quarter); cap.cell(recon_row,2,role); cap.cell(recon_row,3,region)
                cap.cell(recon_row,4,f'=SUMIFS($D${dstart}:$D${dend},$A${dstart}:$A${dend},B{recon_row},$B${dstart}:$B${dend},C{recon_row},$C${dstart}:$C${dend},A{recon_row})')
                cap.cell(recon_row,5,f'=SUMIFS($E${sstart}:$E${send},$A${sstart}:$A${send},B{recon_row},$C${sstart}:$C${send},C{recon_row},$D${sstart}:$D${send},A{recon_row})')
                cap.cell(recon_row,6,f"=E{recon_row}-D{recon_row}")
                cap.cell(recon_row,7,exception)
                cap.cell(recon_row,8,f'=IF(G{recon_row}<>"", "EXEMPT", IF(ABS(F{recon_row})<=ASSUMPTIONS!C{a["cell_tolerance"]},"PASS","FAIL"))')
                recon_keys.append((quarter,role,region,recon_row))
                recon_row+=1
    recon_end=recon_row-1
    add_table(cap,"RoleRegionQuarterCheck",f"A{recon_head}:H{recon_end}")
    for col in range(1,14): cap.column_dimensions[get_column_letter(col)].width=max(cap.column_dimensions[get_column_letter(col)].width or 10,12)
    ramp_head=recon_end+3
    headers(cap,ramp_head,["Vendor"]+QUARTERS)
    ramp_start=ramp_head+1
    for rownum,vendor in enumerate(PARAMS["capacity"]["vendors"],ramp_start):
        cap.cell(rownum,1,vendor)
        for col,quarter in enumerate(QUARTERS,2):
            cap.cell(rownum,col,f'=SUMIFS($E${sstart}:$E${send},$B${sstart}:$B${send},$A{rownum},$D${sstart}:$D${send},"{quarter}")')
            cap.cell(rownum,col).font=Font(color="1F5EAA")
    ramp_end=ramp_start+len(PARAMS["capacity"]["vendors"])-1
    add_table(cap,"VendorRampByVendor",f"A{ramp_head}:M{ramp_end}")
    bench_head=ramp_end+3
    headers(cap,bench_head,["Quarter","Bench surplus FTE"])
    bench_start=bench_head+1
    for rownum,q in enumerate(QUARTERS,bench_start):
        source_row=5+QUARTERS.index(q)
        cap.cell(rownum,1,q); cap.cell(rownum,2,f"=H{source_row}")
        cap.cell(rownum,2).font=Font(color="1F5EAA")
    bench_end=bench_start+len(QUARTERS)-1
    add_table(cap,"BenchByQuarter",f"A{bench_head}:B{bench_end}")
    util_head=bench_end+3
    headers(cap,util_head,["Quarter","Demand utilisation"])
    util_start=util_head+1
    for rownum,q in enumerate(QUARTERS,util_start):
        source_row=5+QUARTERS.index(q)
        cap.cell(rownum,1,q); cap.cell(rownum,2,f'=IFERROR(C{source_row}/G{source_row},0)')
        cap.cell(rownum,2).number_format="0.0%"; cap.cell(rownum,2).font=Font(color="1F5EAA")
    util_end=util_start+len(QUARTERS)-1
    add_table(cap,"UtilisationByQuarter",f"A{util_head}:B{util_end}")

    # DECISIONS: narrative cells remain placeholders; measures are structured separately.
    title(dec,"Stage A | Decision register",[12,18,32,42,32,32,32,36,20],"A5")
    headers(dec,4,["#","Date","Decision","Options","Evidence","Decider (role)","Outcome","Status label","Page that forces it"])
    measures=[]
    for rownum,item in enumerate(PARAMS["decision_data"],5):
        month=item["period"]
        vals=[item["id"],month,"[Kakali – own words]",item["options"],"[Kakali – own words]",item["decider"],"[Kakali – own words]",item["status"],item["page"]]
        for col,val in enumerate(vals,1):
            c=dec.cell(rownum,col,val)
            c.alignment=Alignment(vertical="top",wrap_text=True)
            if col in (3,5,7):
                c.fill=PatternFill("solid",fgColor="FFF4D6")
    dec["A12"]="Structured evidence and outcomes (numeric facts only)"
    dec["A12"].font=Font(bold=True,size=12,color="15263D")
    headers(dec,14,["Decision ID","Type","Measure","Value","Unit","Status"])
    measure_rows=[]
    for item in PARAMS["decision_data"]:
        for kind,items in [("Evidence",item["evidence"]),("Outcome",item["outcomes"])]:
            for measure in items:
                label,value,unit=measure[:3]
                if len(measure) > 3:
                    status=measure[3]
                elif unit=="forecast (simulated)":
                    status="forecast (simulated)"
                elif item["status"].startswith("mixed:"):
                    status="in-scenario actual (simulated) to 31 Oct 2026"
                else:
                    status=item["status"]
                if unit=="forecast (simulated)":
                    value=datetime.fromisoformat(value)
                measure_rows.append([item["id"],kind,label,value,unit,status])
    for rownum,row in enumerate(measure_rows,15):
        for col,val in enumerate(row,1):
            c=dec.cell(rownum,col,val)
            c.alignment=Alignment(vertical="top",wrap_text=True)
            if col==4:
                c.fill=PatternFill("solid",fgColor="FFF4D6")
                if row[4]=="fraction":
                    c.number_format="0.00%"
                elif row[4]=="$m":
                    c.number_format='$#,##0.00'
                elif row[4]=="date":
                    c.number_format="dd mmm yyyy"
                elif row[4]=="USD":
                    c.number_format='$#,##0'
    add_table(dec,"DecisionMeasures",f"A14:F{14+len(measure_rows)}")
    dec.auto_filter.ref="A4:I9"
    dec.freeze_panes="A5"

    # VENDORS: simulated scorecard, remediation, transition, and continuity.
    vdata=PARAMS["vendor_scorecard"]
    title(vend,"Stage B1 | Vendor performance and continuity",[16,25,18,14,18,14,18,14,18,14,16,46,42,20],"A5")
    headers(vend,4,["Vendor","As at / disposition","SLA attainment %","SLA RAG","Defect escape %","Quality RAG","Milestones on time %","Delivery RAG","Spend variance %","Cost RAG","Overall RAG","Basis / notes","Threshold basis","Status label"])
    vendor_rows={}
    for rownum,(vendor,metrics) in enumerate(vdata["vendors"].items(),5):
        vendor_rows[vendor]=rownum
        period="exited 10 Jun 2026" if vendor=="Vendor A" else "as at 31 Oct 2026"
        vals=[vendor,period,metrics["sla"],None,metrics["quality"],None,metrics["delivery"],None,metrics["cost_variance"],None,None,metrics["basis"],"assumption: thresholds in rows 14–17", "in-scenario actual (simulated)"]
        for col,val in enumerate(vals,1):
            if val is not None: vend.cell(rownum,col,val)
        vend.cell(rownum,4,f'=IF(C{rownum}>=95,"Green",IF(C{rownum}>=90,"Amber","Red"))')
        vend.cell(rownum,6,f'=IF(E{rownum}<=1.5,"Green",IF(E{rownum}<=2,"Amber","Red"))')
        vend.cell(rownum,8,f'=IF(G{rownum}>=90,"Green",IF(G{rownum}>=80,"Amber","Red"))')
        vend.cell(rownum,10,f'=IF(ABS(I{rownum})<=5,"Green",IF(ABS(I{rownum})<=10,"Amber","Red"))')
        vend.cell(rownum,11,f'=IF(COUNTIF(D{rownum}:J{rownum},"Red")>0,"Red",IF(COUNTIF(D{rownum}:J{rownum},"Amber")>0,"Amber","Green"))')
        for col in (3,5,7,9): vend.cell(rownum,col).number_format="0.0"
    headers(vend,13,["Dimension","Green rule","Amber rule","Red rule","Basis","Green min / max","Amber min / max"])
    threshold_rows=[]
    rules=[("SLA","≥95%", "≥90% and <95%", "<90%"),("Quality","≤1.5%", ">1.5% and ≤2.0%", ">2.0%"),("Delivery","≥90%", "≥80% and <90%", "<80%"),("Cost variance","|x|≤5%", "5%<|x|≤10%", "|x|>10%")]
    for rownum,(dimension,green,amber,red) in enumerate(rules,14):
        threshold_rows.append(rownum)
        threshold_key={"SLA":"sla","Quality":"quality","Delivery":"delivery","Cost variance":"cost"}[dimension]
        for col,val in enumerate([dimension,green,amber,red,vdata["thresholds"][threshold_key]["basis"]],1): vend.cell(rownum,col,val)
    vend["F14"]=vdata["thresholds"]["sla"]["green_min"]; vend["G14"]=vdata["thresholds"]["sla"]["amber_min"]
    vend["F15"]=vdata["thresholds"]["quality"]["green_max"]; vend["G15"]=vdata["thresholds"]["quality"]["amber_max"]
    vend["F16"]=vdata["thresholds"]["delivery"]["green_min"]; vend["G16"]=vdata["thresholds"]["delivery"]["amber_min"]
    vend["F17"]=vdata["thresholds"]["cost"]["green_abs_max"]; vend["G17"]=vdata["thresholds"]["cost"]["amber_abs_max"]
    for rownum in range(5,10):
        vend.cell(rownum,4,f'=IF(C{rownum}>=$F$14,"Green",IF(C{rownum}>=$G$14,"Amber","Red"))')
        vend.cell(rownum,6,f'=IF(E{rownum}<=$F$15,"Green",IF(E{rownum}<=$G$15,"Amber","Red"))')
        vend.cell(rownum,8,f'=IF(G{rownum}>=$F$16,"Green",IF(G{rownum}>=$G$16,"Amber","Red"))')
        vend.cell(rownum,10,f'=IF(ABS(I{rownum})<=$F$17,"Green",IF(ABS(I{rownum})<=$G$17,"Amber","Red"))')
    headers(vend,20,["Month 2026","Vendor B defect escape %","Trend RAG","Two-month sustained target","Status label"])
    month_names=["Jan 2026","Feb 2026","Mar 2026","Apr 2026","May 2026","Jun 2026","Jul 2026","Aug 2026","Sep 2026","Oct 2026"]
    monthly_start=21
    for ix,(month,value) in enumerate(zip(month_names,vdata["quality_monthly_pct"]),monthly_start):
        vend.cell(ix,1,month); vend.cell(ix,2,value); vend.cell(ix,2).number_format="0.0\"%\""
        vend.cell(ix,3,f'=IF(B{ix}<=1.5,"Green",IF(B{ix}<=2,"Amber","Red"))')
        vend.cell(ix,4,f'=IF(AND(B{ix}<=$F$15,B{ix+1}<=$F$15),"met","not met")' if ix<monthly_start+9 else '=IF(AND(B29<=$F$15,B30<=$F$15),"met","not met at 31 Oct 2026")')
        vend.cell(ix,5,f'=IF(DATE(2026,{ix-monthly_start+1},1)<=ASSUMPTIONS!C{a["as_at"]},"in-scenario actual (simulated)","forecast (simulated)")')
    rem=vdata["remediation"]
    vend["F20"]="Remediation start"; vend["G20"]=datetime.fromisoformat(rem["start"]); vend["G20"].number_format="dd mmm yyyy"
    headers(vend,33,["Checkpoint","Date","Criterion max %","Observed %","Days from start","Result","Basis","Status label"])
    checkpoint_start=34
    for rownum,item in enumerate(rem["checkpoints"],checkpoint_start):
        vals=[item["name"],datetime.fromisoformat(item["date"]),item["criterion_max_pct"],item["value_pct"],f'=B{rownum}-$G$20',f'=IF(D{rownum}<=C{rownum},"met","not met")',item["basis"],f'=IF(B{rownum}<=ASSUMPTIONS!C{a["as_at"]},"in-scenario actual (simulated)","forecast (simulated)")']
        for col,val in enumerate(vals,1): vend.cell(rownum,col,val)
        vend.cell(rownum,2).number_format="dd mmm yyyy"
        for col in (3,4): vend.cell(rownum,col).number_format="0.0\"%\""
    vend["I20"]="Exit gate elapsed calendar days"; vend["J20"]="=B35-$G$20"
    vend["I21"]="Held milestone payments released"; vend["J21"]=2
    vend["I22"]="Parallel run added spend"; vend["J22"]=PARAMS["commercial"]["remediation_parallel_run_added_spend_usd"]
    vend["K22"]=PARAMS["commercial"]["basis_remediation"]
    vend["I23"]="Amber rule"; vend["J23"]="≤1.5% for two consecutive months (assumption)"
    vend["I24"]="Amber at cut-off"; vend["J24"]='=IF(AND(B29<=$F$15,B30<=$F$15),"Green","not met at 31 Oct 2026")'
    vend["I25"]="Amber status basis"; vend["J25"]="assumption (criterion, not met at cut-off)"
    vend["I27"]="Notice served"; vend["J27"]=f'=ASSUMPTIONS!C{a["notice_start"]}'
    vend["I28"]="Notice expires"; vend["J28"]=f'=ASSUMPTIONS!C{a["notice_end"]}'
    vend["I29"]="Dual-run working days"; vend["J29"]=f'=ASSUMPTIONS!C{a["dual_days"]}'
    vend["I30"]="Paid, not delivering working days"; vend["J30"]=f'=ASSUMPTIONS!C{a["paid_days"]}'
    vend["I31"]="Overlap cost, contingency"; vend["J31"]=f'=ASSUMPTIONS!C{a["overlap_cost"]}'
    vend["I32"]="Q1 2026 SLA breaches"; vend["J32"]=2; vend["K32"]="Spec §8 D3"
    vend["I33"]="Total transition working days"; vend["J33"]="=J29+J30"; vend["J33"].number_format="0"
    for r in (27,28): vend.cell(r,10).number_format="ddd dd mmm yyyy"
    vend["J31"].number_format='$#,##0'
    for r in range(27,32): vend.cell(r,11,'="in-scenario actual (simulated)"')
    headers(vend,40,["KT ID","Date","Milestone","Status label","Basis"])
    kt_start=41
    for rownum,item in enumerate(vdata["kt_milestones"],kt_start):
        vals=[item["id"],datetime.fromisoformat(item["date"]),item["milestone"],f'=IF(B{rownum}<=ASSUMPTIONS!C{a["as_at"]},"in-scenario actual (simulated)","forecast (simulated)")',item["basis"]]
        for col,val in enumerate(vals,1): vend.cell(rownum,col,val)
        vend.cell(rownum,2).number_format="ddd dd mmm yyyy"
    headers(vend,48,["Risk ID","Type","Role","Vendor","What is at risk","Trigger","Owner (role)","Mitigation","Date","Basis","Status label"])
    risk_start=49
    for rownum,item in enumerate(vdata["continuity_risks"],risk_start):
        vals=[item["id"],item["type"],item["role"],item["vendor"],item["risk"],item["trigger"],item["owner"],"[Kakali]",datetime.fromisoformat(item["date"]),item["basis"],f'=IF(I{rownum}<=ASSUMPTIONS!C{a["as_at"]},"in-scenario actual (simulated)","forecast (simulated)")']
        for col,val in enumerate(vals,1): vend.cell(rownum,col,val)
        vend.cell(rownum,9).number_format="dd mmm yyyy"
    add_table(vend,"VendorContinuityRisks","A48:K51")
    vend.freeze_panes="A5"

    # COMMERCIAL: formula allocations from base supply; the overlap stays outside vendor spend.
    cparams=PARAMS["commercial"]
    title(comm,"Stage B1 | Commercial exposure and funding",[18,24,17,48,18,18,18,18,18,18,18,18],"A5")
    comm["A3"]="Base vendor FTE only; transition overlap is excluded from vendor spend and shown on contingency."
    comm["A3"].font=Font(italic=True,color="68778A")
    headers(comm,4,["Vendor","Role","Input rate $/day","Basis","2026 FTE-months","2027 FTE-months","2028 FTE-months","2026 effective rate","2027 effective rate","2028 effective rate","Source base FTE share","Notes"])
    role_vendor_rows=[]
    rate_rows={}
    rate_start=5
    for vendor in PARAMS["capacity"]["vendors"]:
        for role in PARAMS["capacity"]["role_taxonomy"]["names"]:
            rownum=rate_start+len(role_vendor_rows)
            role_vendor_rows.append((vendor,role,rownum))
            comm.cell(rownum,1,vendor); comm.cell(rownum,2,role)
            comm.cell(rownum,3,cparams["rate_card"][vendor][role])
            comm.cell(rownum,4,cparams["rate_card_basis"][vendor][role])
            for col,year in zip((5,6,7),(2026,2027,2028)):
                year_quarters=[q for q in QUARTERS if q.endswith(str(year))]
                terms=[]
                for qi,quarter in enumerate(year_quarters):
                    if year==2026:
                        month_weight=(f'ASSUMPTIONS!C{a["y1_fte_a"]}*3' if qi<3 else
                                      f'ASSUMPTIONS!C{a["y1_fte_a"]}+ASSUMPTIONS!C{a["y1_fte_f"]}*2')
                    elif year==2027:
                        month_weight=f'ASSUMPTIONS!C{a[f"y2q{qi+1}"]}*3'
                    else:
                        month_weight=f'ASSUMPTIONS!C{a[f"y3q{qi+1}"]}*3'
                    role_share=(f'SUMIFS(CAPACITY!$E${sstart}:$E${send},CAPACITY!$A${sstart}:$A${send},$B{rownum},'
                                f'CAPACITY!$B${sstart}:$B${send},$A{rownum},CAPACITY!$D${sstart}:$D${send},"{quarter}",'
                                f'CAPACITY!$F${sstart}:$F${send},"Baseline")/SUMIFS(CAPACITY!$E${sstart}:$E${send},'
                                f'CAPACITY!$B${sstart}:$B${send},"<>Internal",CAPACITY!$D${sstart}:$D${send},"{quarter}",'
                                f'CAPACITY!$F${sstart}:$F${send},"Baseline")')
                    terms.append(f'IFERROR(({role_share})*({month_weight}),0)')
                formula=f'=SUM({",".join(terms)})'
                comm.cell(rownum,col,formula)
    rate_end=rate_start+len(role_vendor_rows)-1
    raw_blend_row=57
    headers(comm,raw_blend_row,["Blended rate control","2026","2027","2028","Basis"])
    comm.cell(58,1,"Vendor FTE-months total")
    comm.cell(58,2,f'=ASSUMPTIONS!C{a["y1_fte_a"]}*ASSUMPTIONS!C{a["y1_month_a"]}+ASSUMPTIONS!C{a["y1_fte_f"]}*ASSUMPTIONS!C{a["y1_month_f"]}')
    comm.cell(58,3,f'=SUM(ASSUMPTIONS!C{a["y2q1"]}:C{a["y2q4"]})*3')
    comm.cell(58,4,f'=SUM(ASSUMPTIONS!C{a["y3q1"]}:C{a["y3q4"]})*3')
    comm.cell(59,1,"Raw weighted rate")
    comm.cell(59,2,f'=SUMPRODUCT(C{rate_start}:C{rate_end},E{rate_start}:E{rate_end})/B58')
    comm.cell(59,3,f'=SUMPRODUCT(C{rate_start}:C{rate_end},F{rate_start}:F{rate_end})/C58')
    comm.cell(59,4,f'=SUMPRODUCT(C{rate_start}:C{rate_end},G{rate_start}:G{rate_end})/D58')
    comm.cell(60,1,cparams["rate_scaling"]["label"])
    for col in range(2,5): comm.cell(60,col,f'=COMMERCIAL!{get_column_letter(col)}61/{get_column_letter(col)}59')
    comm.cell(61,1,"Target blended rate $/day")
    for col in range(2,5): comm.cell(61,col,cparams["rate_blend_target_usd_per_day"])
    comm.cell(62,1,"Effective weighted rate $/day")
    for col in range(2,5): comm.cell(62,col,f'={get_column_letter(col)}59*{get_column_letter(col)}60')
    for col in range(2,5): comm.cell(60,col).number_format='0.0000"x"'
    comm["E57"]=cparams["basis"]
    for vendor,role,rownum in role_vendor_rows:
        for col,blend_col,fte_col in ((8,"B",5),(9,"C",6),(10,"D",7)):
            comm.cell(rownum,col,f'=C{rownum}*{blend_col}$60')
        comm.cell(rownum,11,f'=IFERROR(AVERAGE(E{rownum}:G{rownum})/AVERAGE($B$58:$D$58),0)')
    headers(comm,65,["Vendor","2026 spend $m","2027 spend $m","2028 spend $m","Three-year $m","Share of three-year spend","RAG period note"])
    spend_start=66
    spend_rows={}
    for rownum,vendor in enumerate(PARAMS["capacity"]["vendors"],spend_start):
        spend_rows[vendor]=rownum
        comm.cell(rownum,1,vendor)
        for col,fte_col,rate_col in ((2,"E","H"),(3,"F","I"),(4,"G","J")):
            comm.cell(rownum,col,f'=SUMPRODUCT(($A${rate_start}:$A${rate_end}=$A{rownum})*{fte_col}${rate_start}:{fte_col}${rate_end}*{rate_col}${rate_start}:{rate_col}${rate_end})*220/12/1000000')
        comm.cell(rownum,5,f'=SUM(B{rownum}:D{rownum})')
        comm.cell(rownum,6,f'=E{rownum}/SUM($E${spend_start}:$E${spend_start+4})')
        comm.cell(rownum,7,"exited 10 Jun 2026" if vendor=="Vendor A" else "simulated")
        comm.cell(rownum,6).number_format="0.0%"
    totalrow=spend_start+5
    comm.cell(totalrow,1,"Total vendor spend")
    for col in range(2,7): comm.cell(totalrow,col,f'=SUM({get_column_letter(col)}{spend_start}:{get_column_letter(col)}{spend_start+4})')
    comm["A73"]="Year 1 period split"; comm["B73"]=f'=ASSUMPTIONS!C{a["v_y1_a"]}'; comm["C73"]=f'=ASSUMPTIONS!C{a["v_y1_f"]}'
    comm["D73"]="in-scenario actual (simulated)"; comm["E73"]="forecast (simulated)"
    comm["A75"]="Contingency drawdown"; comm["B75"]=f'=ASSUMPTIONS!C{a["overlap_cost"]}/1000000'; comm["C75"]="A-to-D overlap; paid from contingency"
    comm["A76"]="Contingency unallocated"; comm["B76"]=f'=ASSUMPTIONS!C{a["unalloc"]}/1000000'
    comm["A77"]="Total contingency"; comm["B77"]="=B75+B76"
    headers(comm,79,["Funded vs spend","Year 1","Year 2","Year 3","Basis"])
    comm["A80"]="Vendor spend"; comm["B80"]=f'=ASSUMPTIONS!C{a["v_y1"]}'; comm["C80"]=f'=ASSUMPTIONS!C{a["v_y2"]}'; comm["D80"]=f'=ASSUMPTIONS!C{a["v_y3"]}'
    comm["A81"]="Other costs"; comm["B81"]=f'=ASSUMPTIONS!C{a["other_y1"]}'; comm["C81"]=f'=ASSUMPTIONS!C{a["other_y2"]}'; comm["D81"]=f'=ASSUMPTIONS!C{a["other_y3"]}'
    comm["A82"]="Funded work"; comm["B82"]=f'=ASSUMPTIONS!C{a["fund_y1"]}'; comm["C82"]=f'=ASSUMPTIONS!C{a["fund_y2"]}'; comm["D82"]=f'=ASSUMPTIONS!C{a["fund_y3"]}'
    comm["A83"]="Contingency"; comm["B83"]=f'=ASSUMPTIONS!C{a["cont"]}'; comm["C83"]=f'=ASSUMPTIONS!C{a["cont"]}'; comm["D83"]=f'=ASSUMPTIONS!C{a["cont"]}'
    comm.freeze_panes="A5"

    # PLAN: dated milestone schedule, decision rights, RAID and recovery facts.
    plandata=PARAMS["plan"]
    title(plan,"Stage B1 | Integrated delivery plan",[13,12,16,64,22,20,28,23,23,26,14],"A5")
    headers(plan,4,["Milestone","Workstream","Date","Milestone / deliverable","Source","Predecessors","Status label","Owner role","Mitigation / note","Basis","On-time flag"])
    milestone_start=5
    for rownum,item in enumerate(plandata["milestones"],milestone_start):
        date=datetime.fromisoformat(item["date"])
        status="in-scenario actual (simulated)" if date.date()<=datetime.fromisoformat(PARAMS["as_at"]).date() else "forecast (simulated)"
        vals=[item["id"],item["workstream"],date,item["name"],item["source"],item["predecessors"],None,"Workstream lead","",item["basis"],item["on_time"]]
        for col,val in enumerate(vals,1): plan.cell(rownum,col,val)
        plan.cell(rownum,3).number_format="ddd dd mmm yyyy"
        plan.cell(rownum,7,f'=IF(C{rownum}<=ASSUMPTIONS!C{a["as_at"]},"in-scenario actual (simulated)","forecast (simulated)")')
    plan["C7"].number_format="mmm yyyy"
    plan["A21"]="Critical path"; plan["B21"]="M01 → M02 → golden source data-engineering tranche (Apr–Jun 2026) → M07/M08 regulatory change build → M09 regulatory date"
    plan.merge_cells("B21:J21")
    plan["A22"]="Float before D4 (days)"; plan["B22"]=0; plan["C22"]= "Spec §8 D4; stated input"
    plan["A23"]="Float at regulatory date (working days)"; plan["B23"]="=NETWORKDAYS(C12,C13)-1"; plan["C23"]="Formula: M08 completion to M09 date, weekends excluded; public holidays ignored."
    headers(plan,26,["Decision type","Driver","Approver","Contributors","Informed","Basis"])
    for rownum,item in enumerate(plandata["daci"],27):
        for col,key in enumerate(("type","driver","approver","contributors","informed","basis"),1): plan.cell(rownum,col,item[key])
    headers(plan,35,["RAID ID","Type","Description","Owner (role)","Trigger (measure + threshold)","Mitigation","Date","Basis"])
    raid_start=36
    for rownum,item in enumerate(plandata["raid"],raid_start):
        vals=[item["id"],item["type"],item["description"],item["owner"],item["trigger"],"[Kakali]",datetime.fromisoformat(item["date"]),item["basis"]]
        for col,val in enumerate(vals,1): plan.cell(rownum,col,val)
        plan.cell(rownum,7).number_format="dd mmm yyyy"
    headers(plan,43,["Escalation role","Response time","Basis"])
    for rownum,item in enumerate(plandata["escalation"],44):
        plan.cell(rownum,1,item["role"]); plan.cell(rownum,2,item["response"]); plan.cell(rownum,3,item["basis"])
    headers(plan,50,["Recovery plan","Starting position","Action","Result","Status / basis"])
    recovery=[
      ["Schedule recovery","Float 0 days","Two reports phased to phase 2; no squads crashed","Float 11 days; 29 Jan 2027 date protected","Forecast; spec §8 D4 and decision log governs"],
      ["Quality recovery","Defect escape 4.1% vs 1.5% target","Re-compete; vendor-funded parallel run; two payments held","1.6% exit gate met in 54 days; payments released; $0 added spend","Simulated; spec §8 D3"]]
    for rownum,vals in enumerate(recovery,51):
        for col,val in enumerate(vals,1): plan.cell(rownum,col,val)
    plan["A54"]="Quarterly grid convention"; plan["B54"]="Every quarter shows the first-day position; period-average vendor FTE drives COMMERCIAL."
    plan.merge_cells("B54:J54")
    cadence=PARAMS["cadence"]
    plan["A56"]="Cadence block"
    plan["A56"].font=Font(bold=True,size=12,color="15263D")
    headers(plan,57,["Forum","Frequency","Chair","Attends","Decides here and nowhere else","Pack received","Basis"])
    cadence_start=58
    cadence_rows={}
    for rownum,item in enumerate(cadence["forums"],cadence_start):
        cadence_rows[item["forum"]]=rownum
        vals=[item["forum"],item["frequency"],item["chair"],item["attends"],item["decides"],item["pack"],item["basis"]]
        for col,val in enumerate(vals,1):
            plan.cell(rownum,col,val)
            plan.cell(rownum,col).alignment=Alignment(wrap_text=True,vertical="top")
        plan.row_dimensions[rownum].height=62
    add_table(plan,"DeliveryCadence",f"A57:G{cadence_start+len(cadence['forums'])-1}")
    measure_header=63
    plan["A62"]="Managing director monthly measures"
    plan["A62"].font=Font(bold=True,size=12,color="15263D")
    headers(plan,measure_header,["Measure","Value","Benchmark","Basis"])
    measures_start=measure_header+1
    for rownum,item in enumerate(cadence["managing_director_measures"],measures_start):
        if item["id"]=="Contingency":
            value_formula=f'=ASSUMPTIONS!C{a["unalloc"]}/1000000'
            benchmark_formula=f'=ASSUMPTIONS!C{a["cont"]}'
        else:
            value_formula=f'={item["value_ref"]}'
            benchmark_formula=f'={item["benchmark_ref"]}'
        vals=[item["label"],value_formula,benchmark_formula,item["basis"]]
        for col,val in enumerate(vals,1):
            plan.cell(rownum,col,val)
            plan.cell(rownum,col).alignment=Alignment(wrap_text=True,vertical="top")
            if col in (2,3): plan.cell(rownum,col).font=Font(color="1F5EAA")
        if item["id"]=="H1":
            plan.cell(rownum,2).number_format='0" days"'; plan.cell(rownum,3).number_format='0" days"'
        elif item["id"]=="H6":
            plan.cell(rownum,2).number_format="0.0%"; plan.cell(rownum,3).number_format="0.0%"
        else:
            plan.cell(rownum,2).number_format='$0.00"m"'; plan.cell(rownum,3).number_format='$0.00"m"'
    add_table(plan,"ManagingDirectorMeasures",f"A{measure_header}:D{measures_start+len(cadence['managing_director_measures'])-1}")
    plan["A68"]="Cadence rationale"
    plan["A68"].font=Font(bold=True,color="15263D")
    plan["B68"]="=PageText_cadence_rationale"
    plan.merge_cells("B68:J69")
    plan["B68"].alignment=Alignment(wrap_text=True,vertical="top")
    plan.row_dimensions[68].height=28
    plan.row_dimensions[69].height=28
    plan.freeze_panes="A5"

    # BENEFITS: value dates drive simulated actual/forecast labels; sourced values remain formulas.
    bparams=PARAMS["benefits"]
    title(benefits,"Stage B2 | Benefits and delivery health",[8,23,27,31,48,16,16,16,16,16,16,31,26,42,50],"A5")
    benefits["A3"]="Benefits are run-budget effects outside the scenario envelope; no benefit is netted against funding or spend."
    benefits.merge_cells("A3:O3")
    benefits["A3"].font=Font(italic=True,color="68778A")
    headers(benefits,4,["ID","Category","Portfolio line","Benefit","Measure","Baseline","Baseline date","Target","Target date","Value","Value date","Status label","Owner (role)","Source","Basis"])
    benefit_rows={}
    for rownum,item in enumerate(bparams["register"],5):
        benefit_rows[item["id"]]=rownum
        baseline_date=datetime.fromisoformat(item["baseline_date"]) if item["baseline_date"] else None
        target_date=datetime.fromisoformat(item["target_date"])
        value_date=datetime.fromisoformat(item["value_date"])
        vals=[item["id"],item["category"],item["portfolio_line"],item["benefit"],item["measure"],
              item["baseline_value"],baseline_date,item["target_value"],target_date,None,value_date,None,
              item["owner"],item["source"],item["basis"]]
        for col,val in enumerate(vals,1):
            cell=benefits.cell(rownum,col,val)
            cell.alignment=Alignment(vertical="top",wrap_text=True)
            if col in (6,8,11): cell.fill=PatternFill("solid",fgColor="FFF4D6")
        if item["id"]=="B1":
            benefits.cell(rownum,10,'=COUNTIF(VENDORS!$B$5:$B$9,"<>exited 10 Jun 2026")')
        elif item["id"]=="B5":
            benefits.cell(rownum,10,"=VENDORS!B30")
        else:
            benefits.cell(rownum,10,item["value"])
        benefits.cell(rownum,12,f'=IF(K{rownum}<=ASSUMPTIONS!C{a["as_at"]},"in-scenario actual (simulated)","forecast (simulated)")')
        benefits.cell(rownum,10).font=Font(color="1F5EAA" if item.get("value_source") else "000000")
        if item["id"] in ("B3","B6"):
            for col in (6,8,10): benefits.cell(rownum,col).number_format="0.0%"
        elif item["id"]=="B5":
            for col in (6,8,10): benefits.cell(rownum,col).number_format='0.0"%"'
        elif item["id"]=="B2":
            for col in (6,8,10): benefits.cell(rownum,col).number_format='$#,##0.0'
        else:
            for col in (6,8,10): benefits.cell(rownum,col).number_format="0"
        for col in (7,9,11): benefits.cell(rownum,col).number_format="dd mmm yyyy"
        benefits.cell(rownum,12).font=Font(color="1F5EAA")
        benefits.row_dimensions[rownum].height=50
    add_table(benefits,"BenefitsRegister","A4:O10")
    benefits["A12"]="Benefits are run-budget effects; the scenario envelope, funded work and vendor spend remain unchanged."
    benefits.merge_cells("A12:O12")
    benefits["A12"].font=Font(italic=True,color="68778A")

    # Adoption funnel, with all conversion rates calculated in the workbook.
    headers(benefits,14,["Business unit","Invited","Trained","Active","Habitual","Trained / invited","Active / trained","Habitual / invited","Habitual / active"])
    funnel_start=15
    for rownum,item in enumerate(bparams["adoption"]["business_units"],funnel_start):
        for col,key in enumerate(("name","invited","trained","active","habitual"),1):
            benefits.cell(rownum,col,item[key])
            benefits.cell(rownum,col).alignment=Alignment(vertical="top",wrap_text=True)
            if col>1: benefits.cell(rownum,col).fill=PatternFill("solid",fgColor="FFF4D6")
        benefits.cell(rownum,6,f'=IFERROR(C{rownum}/B{rownum},0)')
        benefits.cell(rownum,7,f'=IFERROR(D{rownum}/C{rownum},0)')
        benefits.cell(rownum,8,f'=IFERROR(E{rownum}/B{rownum},0)')
        benefits.cell(rownum,9,f'=IFERROR(E{rownum}/D{rownum},0)')
        for col in range(6,10):
            benefits.cell(rownum,col).number_format="0.0%"
            benefits.cell(rownum,col).font=Font(color="1F5EAA")
    funnel_total=funnel_start+len(bparams["adoption"]["business_units"])
    benefits.cell(funnel_total,1,"Total")
    for col in range(2,6): benefits.cell(funnel_total,col,f'=SUM({get_column_letter(col)}{funnel_start}:{get_column_letter(col)}{funnel_total-1})')
    benefits.cell(funnel_total,6,f'=IFERROR(C{funnel_total}/B{funnel_total},0)')
    benefits.cell(funnel_total,7,f'=IFERROR(D{funnel_total}/C{funnel_total},0)')
    benefits.cell(funnel_total,8,f'=IFERROR(E{funnel_total}/B{funnel_total},0)')
    benefits.cell(funnel_total,9,f'=IFERROR(E{funnel_total}/D{funnel_total},0)')
    for col in range(2,10):
        benefits.cell(funnel_total,col).font=Font(bold=True,color="1F5EAA")
        if col>=6: benefits.cell(funnel_total,col).number_format="0.0%"
    add_table(benefits,"AdoptionFunnel",f"A14:I{funnel_total-1}")
    benefits["A20"]="Recommendation – not a logged decision"
    benefits["A20"].font=Font(bold=True,color="15263D")
    benefits["B20"]="=PageText_e42_note"
    benefits.merge_cells("B20:O21")
    benefits["B20"].alignment=Alignment(wrap_text=True,vertical="top")
    benefits.row_dimensions[20].height=30
    benefits.row_dimensions[21].height=30

    # Delivery health: formula-linked values and formula-derived RAG.
    health_header=25
    health_titles=["ID","Measure","Value","Green threshold","Amber threshold","Red threshold","RAG","Escalation if amber","Escalation if red","Basis","Direction"]
    headers(benefits,health_header,health_titles)
    health_start=health_header+1
    health_rows={}
    for rownum,item in enumerate(bparams["delivery_health"]["measures"],health_start):
        health_rows[item["id"]]=rownum
        vals=[item["id"],item["measure"],None,item["green"],item["amber"],item["red"],None,
              item["amber_escalation"],item["red_escalation"],item["basis"],item["direction"]]
        for col,val in enumerate(vals,1):
            benefits.cell(rownum,col,val)
            benefits.cell(rownum,col).alignment=Alignment(vertical="top",wrap_text=True)
        if item["id"]=="H1":
            value_formula="=PLAN!B23"
        elif item["id"]=="H2":
            cutoff=f'ASSUMPTIONS!C{a["as_at"]}'
            value_formula=f'=COUNTIFS(PLAN!$C$5:$C$18,"<="&{cutoff},PLAN!$K$5:$K$18,TRUE)/COUNTIF(PLAN!$C$5:$C$18,"<="&{cutoff})'
        elif item["id"]=="H3":
            value_formula="=VENDORS!B30"
        elif item["id"]=="H4":
            value_formula=f'=ROUND(COMMERCIAL!B75,2)/ASSUMPTIONS!C{a["cont"]}'
        elif item["id"]=="H5":
            value_formula=f'=INDEX(CAPACITY!$L$5:$L$16,MATCH("Q"&ROUNDUP(MONTH(ASSUMPTIONS!C{a["as_at"]})/3,0)&" "&YEAR(ASSUMPTIONS!C{a["as_at"]}),CAPACITY!$A$5:$A$16,0))'
        else:
            value_formula=f'=BENEFITS!H{funnel_total}'
        benefits.cell(rownum,3,value_formula)
        if item["direction"]=="higher":
            rag_formula=f'=IF(C{rownum}>=D{rownum},"Green",IF(C{rownum}>=E{rownum},"Amber","Red"))'
        else:
            rag_formula=f'=IF(C{rownum}<=D{rownum},"Green",IF(C{rownum}<=E{rownum},"Amber","Red"))'
        benefits.cell(rownum,7,rag_formula)
        benefits.cell(rownum,3).font=Font(color="1F5EAA")
        benefits.cell(rownum,7).font=Font(color="1F5EAA")
        if item["id"] in ("H2","H4","H6"):
            for col in range(3,7): benefits.cell(rownum,col).number_format="0.0%"
        elif item["id"]=="H3":
            for col in range(3,7): benefits.cell(rownum,col).number_format='0.0"%"'
        benefits.row_dimensions[rownum].height=40
    health_end=health_start+len(bparams["delivery_health"]["measures"])-1
    add_table(benefits,"DeliveryHealth",f"A{health_header}:K{health_end}")
    benefits.freeze_panes="A5"
    benefits.sheet_view.showGridLines=False

    # User-authored prose is loaded verbatim; the YAML file is never modified.
    for sheet_name in ("DECISIONS","PORTFOLIO","VENDORS","PLAN"):
        ws=wb[sheet_name]
        for coordinate,value in KAKALI_TEXT[sheet_name].items():
            if not isinstance(value,str):
                raise ValueError(f"Own-words value must be text: {sheet_name}!{coordinate}")
            cell=ws[coordinate]
            cell.value=value
            cell.alignment=Alignment(vertical="top",wrap_text=True)
            width=max(12,ws.column_dimensions[cell.column_letter].width or 12)
            visual_lines=max(1,math.ceil(len(value)/(width*1.15)))
            ws.row_dimensions[cell.row].height=max(ws.row_dimensions[cell.row].height or 15,visual_lines*15)

    # Checks table is bottom of ASSUMPTIONS; result cells are live formulas.
    checks_row=75
    asm.cell(checks_row,9,"CHECK ID")
    asm.cell(checks_row,10,"Check")
    asm.cell(checks_row,11,"Result")
    for col in range(9,12):
        asm.cell(checks_row,col).font=Font(bold=True,color="FFFFFF")
        asm.cell(checks_row,col).fill=PatternFill("solid",fgColor="31506E")
    q3row=5+QUARTERS.index("Q3 2027")
    dstart=50
    dend=dstart+len(demand_data)-1
    shead=dend+3
    sstart=shead+1
    send=sstart+len(supply_data)-1
    other_role_supply_q3="+".join(f'SUMIFS(CAPACITY!E{sstart}:E{send},CAPACITY!D{sstart}:D{send},"Q3 2027",CAPACITY!A{sstart}:A{send},"{role_name}")' for role_name in PARAMS["capacity"]["role_taxonomy"]["names"][4:])
    checks=[
      ("Funded total and annual phasing tie",f'=AND(ABS(PORTFOLIO!C22-40.8)<0.001,ABS(PORTFOLIO!C26-C{a["fund_y1"]})<0.001,ABS(PORTFOLIO!C27-C{a["fund_y2"]})<0.001,ABS(PORTFOLIO!C28-C{a["fund_y3"]})<0.001)'),
      ("Request equals funded + deferred + stopped",'=AND(ABS(PORTFOLIO!C21-63.8)<0.001,ABS(PORTFOLIO!C21-PORTFOLIO!C22-PORTFOLIO!C23-PORTFOLIO!C24)<0.001)'),
      ("Line 1 + 5a equals pre-rescope line",'=AND(ABS(PORTFOLIO!G5-PORTFOLIO!H5-PORTFOLIO!M5)<0.001,ABS(PORTFOLIO!M5+PORTFOLIO!M10-12)<0.001)'),
      ( "Contingency and envelope tie",f'=AND(ABS(C{a["cont"]}-C{a["funded_total"]}*5%)<0.001,ABS(C{a["cont"]}-2.04)<0.001,ABS(C{a["envelope"]}-42.84)<0.001)'),
      ("Vendor plus other equals funded annually",f'=AND(ABS(C{a["vendor_y1"]}+C{a["other_y1"]}-C{a["fund_y1"]})<0.001,ABS(C{a["vendor_y2"]}+C{a["other_y2"]}-C{a["fund_y2"]})<0.001,ABS(C{a["vendor_y3"]}+C{a["other_y3"]}-C{a["fund_y3"]})<0.001)'),
      ("Vendor FTE-month calculation rounds to target",f'=AND(ROUND(C{a["v_y1"]},1)=C{a["vendor_y1"]},ROUND(C{a["v_y2"]},1)=C{a["vendor_y2"]},ROUND(C{a["v_y3"]},1)=C{a["vendor_y3"]})'),
      ("Year 1 vendor split rounds to target",f'=AND(ROUND(C{a["v_y1_a"]},1)=8.7,ROUND(C{a["v_y1_f"]},1)=1.9,ROUND(C{a["v_y1"]},1)=10.6)'),
      ("Overlap, reserve and business days tie",f'=AND(C{a["overlap_cost"]}=183600,C{a["unalloc"]}=1856400,C{a["dual_days"]}=40,C{a["paid_days"]}=11,C{a["dual_days"]}+C{a["paid_days"]}=51)'),
      ("Specified weekdays tie",f'=AND(WEEKDAY(C{a["notice_start"]},2)=1,WEEKDAY(C{a["notice_end"]},2)=3,WEEKDAY(DATE(2027,1,29),2)=5)'),
      ("Role bridge moves and peak supply reconcile",f'=AND(SUM(CAPACITY!E21:E25)=0,SUM(CAPACITY!F21:F25)=0,SUM(CAPACITY!G21:G25)=0,SUM(CAPACITY!H21:H25)=0,SUM(CAPACITY!C21:C25)=186,SUM(CAPACITY!I21:I25)=186,CAPACITY!J21=0,CAPACITY!J22=0,CAPACITY!J23=-1,CAPACITY!J24=1,CAPACITY!J25=0,ABS(CAPACITY!I21-SUMIFS(CAPACITY!E{sstart}:CAPACITY!E{send},CAPACITY!D{sstart}:CAPACITY!D{send},"Q3 2027",CAPACITY!A{sstart}:CAPACITY!A{send},"Data engineering"))=0,ABS(CAPACITY!I22-SUMIFS(CAPACITY!E{sstart}:CAPACITY!E{send},CAPACITY!D{sstart}:CAPACITY!D{send},"Q3 2027",CAPACITY!A{sstart}:CAPACITY!A{send},"Integration"))=0,ABS(CAPACITY!I23-SUMIFS(CAPACITY!E{sstart}:CAPACITY!E{send},CAPACITY!D{sstart}:CAPACITY!D{send},"Q3 2027",CAPACITY!A{sstart}:CAPACITY!A{send},"QA"))=0,ABS(CAPACITY!I24-SUMIFS(CAPACITY!E{sstart}:CAPACITY!E{send},CAPACITY!D{sstart}:CAPACITY!D{send},"Q3 2027",CAPACITY!A{sstart}:CAPACITY!A{send},"Business analysis"))=0,ABS(CAPACITY!I25-({other_role_supply_q3}))=0,SUMIFS(CAPACITY!E{sstart}:CAPACITY!E{send},CAPACITY!D{sstart}:CAPACITY!D{send},"Q3 2027",CAPACITY!A{sstart}:CAPACITY!A{send},"QA",CAPACITY!B{sstart}:CAPACITY!B{send},"Vendor A")=0,SUMIFS(CAPACITY!E{sstart}:CAPACITY!E{send},CAPACITY!D{sstart}:CAPACITY!D{send},"Q3 2027",CAPACITY!A{sstart}:CAPACITY!A{send},"Data engineering",CAPACITY!B{sstart}:CAPACITY!B{send},"Vendor D")>=6)'),
      ("Q3 demand after deferrals ties",f'=AND(C{a["peak_pre"]}-C{a["req_total"]}=C{a["peak_post"]},C{a["peak_post"]}=CAPACITY!C{q3row},C{a["supply_peak"]}=CAPACITY!G{q3row},CAPACITY!D{q3row}=C{a["vendor_peak"]},CAPACITY!F{q3row}=C{a["internal_peak"]})'),
      ("Automated firm-name scan has no hits","=K72=0"),
      ("Standalone word scan has no hits","=K73=0"),
      ("Q3 2027 is maximum quarterly demand",f'=AND(CAPACITY!C{q3row}=MAX(CAPACITY!C5:C16),CAPACITY!C{q3row}=186)'),
      ("CRM configuration line 5a spends $0.9m in Year 1",'=PORTFOLIO!I10=0.9'),
      ("Every quarter has demand covered by supply",'=MIN(CAPACITY!H5:H16)>=0'),
      ("Q2 total supply less demand equals dual run",'=Q2_2026_Supply-Q2_2026_Demand=DualRun_FTE'),
      ("Q2 QA surplus and Vendor A region cell equal dual run",'=AND(Q2_2026_QA_Surplus=DualRun_FTE,VendorA_QA_Region_Surplus=DualRun_FTE)'),
      ("Q2 total surplus equals QA surplus",'=Q2_2026_Supply-Q2_2026_Demand=Q2_2026_QA_Surplus'),
      ("Data engineering role gap is zero from Q2 2026",'=SUMPRODUCT(ABS(DE_Gap_From_Q2))=0'),
      ("DE timeline and Q1 role total tie",'=AND(DE_Before_M1=-14,DE_After_M1=-6,DE_After_VendorD=0,Q1_2026_DE_Role_Gap=-14)'),
      ("Q2 base vendor supply equals Jan-Oct average",'=Q2_2026_Base_Vendor_Supply=99'),
      ("Role-region cells outside tolerance excluding named exceptions",'=COUNTIF(Cell_Check_Result,"FAIL")=0'),
      ("Q1 DE role gap and regional shortfalls tie",'=AND(Q1_2026_DE_Role_Gap=-14,COUNTIFS(Cell_Role_Range,"Data engineering",Cell_Quarter_Range,"Q1 2026",Cell_Gap_Range,">0")=0)'),
      ("Vendor spend by year and three-year total tie",'=AND(ROUND(COMMERCIAL!B71,1)=10.6,ROUND(COMMERCIAL!C71,1)=12.4,ROUND(COMMERCIAL!D71,1)=6.1,ROUND(COMMERCIAL!E71,1)=29.1,ROUND(COMMERCIAL!B62,0)=480,ROUND(COMMERCIAL!C62,0)=480,ROUND(COMMERCIAL!D62,0)=480)'),
      ("Year 1 period split ties to 8.7 and 1.9",'=AND(ROUND(COMMERCIAL!B73,1)=8.7,ROUND(COMMERCIAL!C73,1)=1.9)'),
      ("Three vendors exceed 25 percent concentration",'=COUNTIF(COMMERCIAL!F66:F70,">25%")=3'),
      ("Overlap is contingency-only; remediation adds zero",f'=AND(ABS(COMMERCIAL!B75-0.1836)<0.0001,VENDORS!J22=0,ABS(COMMERCIAL!B71-ASSUMPTIONS!C{a["v_y1"]})<0.001,ABS(COMMERCIAL!C71-ASSUMPTIONS!C{a["v_y2"]})<0.001,ABS(COMMERCIAL!D71-ASSUMPTIONS!C{a["v_y3"]})<0.001)'),
      ("Regulatory-date float is 11 working days",'=PLAN!B23=11'),
      ("Dated records carry years and future items are forecast",'=AND(COUNT(PLAN!C5:C18)=14,COUNT(VENDORS!B34:B37)=4,COUNT(VENDORS!B41:B46)=6,COUNTIF(PLAN!G5:G18,"forecast (simulated)")=8,COUNTIF(VENDORS!E21:E30,"in-scenario actual (simulated)")=10,COUNTIF(VENDORS!H34:H37,"in-scenario actual (simulated)")=4,COUNTIF(VENDORS!D41:D46,"in-scenario actual (simulated)")=6,COUNTIF(VENDORS!K49:K51,"in-scenario actual (simulated)")=3,COUNTIF(VENDORS!K27:K31,"in-scenario actual (simulated)")=5,PLAN!G18="forecast (simulated)",YEAR(PLAN!C5)=2026,YEAR(PLAN!C18)=2027,YEAR(VENDORS!B34)=2026,YEAR(VENDORS!B37)=2026,YEAR(VENDORS!B41)=2026,YEAR(VENDORS!B46)=2026,YEAR(VENDORS!I49)=2026,YEAR(VENDORS!I51)=2026,YEAR(VENDORS!J27)=2026,YEAR(VENDORS!J28)=2026,VENDORS!J20=54,VENDORS!D34=VENDORS!B25,VENDORS!D35=VENDORS!B26,VENDORS!D36=VENDORS!B26,VENDORS!D37=VENDORS!B27,VENDORS!J24="not met at 31 Oct 2026")'),
      ("Firm-name and status-language scans pass",'=AND(K72=0,K73=0)'),
      ("Benefits register has six rows and the status split ties",'=AND(COUNTA(BENEFITS!A5:A10)=6,COUNTIF(BENEFITS!B5:B10,"Cost-out")=2,COUNTIF(BENEFITS!B5:B10,"Revenue enablement")=2,COUNTIF(BENEFITS!B5:B10,"Risk avoidance")=2,COUNTIF(BENEFITS!L5:L10,"in-scenario actual (simulated)")=2,COUNTIF(BENEFITS!L5:L10,"forecast (simulated)")=4)'),
      ("B1 and B5 values link to their source cells",'=AND(BENEFITS!J5=COUNTIF(VENDORS!B5:B9,"<>exited 10 Jun 2026"),BENEFITS!J9=VENDORS!B30)'),
      ("Adoption funnel is monotone and total share is habitual over invited",'=AND(BENEFITS!B15>=BENEFITS!C15,BENEFITS!C15>=BENEFITS!D15,BENEFITS!D15>=BENEFITS!E15,BENEFITS!B16>=BENEFITS!C16,BENEFITS!C16>=BENEFITS!D16,BENEFITS!D16>=BENEFITS!E16,BENEFITS!B17>=BENEFITS!C17,BENEFITS!C17>=BENEFITS!D17,BENEFITS!D17>=BENEFITS!E17,BENEFITS!B18=SUM(BENEFITS!B15:B17),BENEFITS!E18=SUM(BENEFITS!E15:E17),BENEFITS!H18=BENEFITS!E18/BENEFITS!B18)'),
      ("Delivery health has six formula values and expected RAG pattern",'=AND(COUNTA(BENEFITS!A26:A31)=6,COUNTIF(BENEFITS!G26:G31,"Green")=4,COUNTIF(BENEFITS!G26:G31,"Amber")=2,COUNTIF(BENEFITS!G26:G31,"Red")=0)'),
      ("Cadence approver roles and SteerCo frequency align",f'=AND(COUNTIF(PLAN!C27:C32,PLAN!C{cadence_start})>0,COUNTIF(PLAN!C27:C32,PLAN!C{cadence_rows["SteerCo"]})>0,PLAN!B{cadence_rows["SteerCo"]}="Fortnightly",PLAN!B47="next fortnightly SteerCo")'),
      ("Portfolio outside ask and contingency reconcile to the envelope",f'=AND(ABS(PORTFOLIO!C21-PORTFOLIO!C22-PORTFOLIO!C23-PORTFOLIO!C24)<0.001,ABS(PORTFOLIO!C22+C{a["cont"]}-C{a["envelope"]})<0.001)')
    ]
    check_ids=[f"C{i:02d}" for i in range(1,17)]+["C17","C18","C19","C20","C20b","C21b","C21","C21c","C22","C23","C24","C25","C26","C27","C28","C29","C30","C31","C32","C33","C40"]
    for rownum,(label,formula) in enumerate(checks,checks_row+1):
        asm.cell(rownum,9,check_ids[rownum-checks_row-1])
        asm.cell(rownum,10,label)
        asm.cell(rownum,11,formula)
        asm.cell(rownum,11).font=Font(color="1F5EAA")
        asm.cell(rownum,10).alignment=Alignment(wrap_text=True,vertical="top")
    asm.cell(72,10,"Firm-name scan hits (automated)")
    asm.cell(72,11,0)
    asm.cell(73,10,"Standalone word hits (automated)")
    asm.cell(73,11,0)
    allrow=checks_row+1+len(checks)
    asm.cell(allrow,10,"ALL CHECKS")
    asm.cell(allrow,11,f"=AND(K{checks_row+1}:K{allrow-1})")
    asm.cell(allrow,10).font=Font(bold=True,color="FFFFFF")
    asm.cell(allrow,11).font=Font(bold=True,color="FFFFFF")
    asm.cell(allrow,10).fill=PatternFill("solid",fgColor="15263D")
    asm.cell(allrow,11).fill=PatternFill("solid",fgColor="15263D")
    asm.cell(allrow,11).number_format="General"
    definition_start=allrow+3
    headers(asm,definition_start,["Definition","Meaning","Basis"])
    for rownum,(definition,meaning) in enumerate(PARAMS["benefits"]["adoption"]["definitions"].items(),definition_start+1):
        asm.cell(rownum,1,definition)
        asm.cell(rownum,2,meaning)
        asm.cell(rownum,3,PARAMS["benefits"]["adoption"]["basis"])
        for col in range(1,4): asm.cell(rownum,col).alignment=Alignment(vertical="top",wrap_text=True)
        asm.row_dimensions[rownum].height=30
    metadata_start=definition_start+1+len(PARAMS["benefits"]["adoption"]["definitions"])+2
    headers(asm,metadata_start,["PDF metadata","Value","Basis"])
    pdfmeta=PARAMS["pdf_metadata"]
    document_date_row=metadata_start+1
    asm.cell(document_date_row,1,"Document date")
    asm.cell(document_date_row,2,datetime.fromisoformat(pdfmeta["document_date"]))
    asm.cell(document_date_row,2).number_format="dd mmm yyyy"
    asm.cell(document_date_row,3,pdfmeta["basis"])
    page_count_row=metadata_start+2
    asm.cell(page_count_row,1,"Page count")
    asm.cell(page_count_row,2,pdfmeta["page_count"])
    asm.cell(page_count_row,3,pdfmeta["basis"])
    pdf_page_rows={}
    for rownum,page_item in enumerate(pdfmeta["page_index"],metadata_start+3):
        asm.cell(rownum,1,f"PDF page {page_item['number']}")
        asm.cell(rownum,2,page_item["number"])
        asm.cell(rownum,3,page_item["basis"])
        pdf_page_rows[page_item["number"]]=f"B{rownum}"
        for col in range(1,4): asm.cell(rownum,col).alignment=Alignment(vertical="top",wrap_text=True)
    pdf_metadata_cells={"document_date":f"B{document_date_row}","page_count":f"B{page_count_row}","page_numbers":pdf_page_rows}
    page_text_start=metadata_start+3+len(pdfmeta["page_index"])+2
    headers(asm,page_text_start,["ID","Text","Used on"])
    page_text_used_on={
        "cover_purpose":"Cover",
        "e42_note":"Page 4, E4.2",
        "cadence_rationale":"Cadence page"
    }
    page_text_cells={}
    for rownum,(key,value) in enumerate(KAKALI_TEXT["PAGE_TEXT"].items(),page_text_start+1):
        asm.cell(rownum,1,key)
        cell=asm.cell(rownum,2,value)
        cell.alignment=Alignment(vertical="top",wrap_text=True)
        cell.font=Font(size=10)
        asm.cell(rownum,3,page_text_used_on[key])
        page_text_cells[key]=f"B{rownum}"
        wb.defined_names.add(DefinedName(f"PageText_{key}",attr_text=f"'ASSUMPTIONS'!$B${rownum}"))
        width=max(1,asm.column_dimensions["B"].width or 38)
        visual_lines=max(1,math.ceil(len(value)/(width*1.25)))
        asm.row_dimensions[rownum].height=max(30,visual_lines*15)
    def define_name(name, reference):
        wb.defined_names.add(DefinedName(name,attr_text=reference))
    for name,ref in {
        "Q2_2026_Supply":"'CAPACITY'!$G$6",
        "Q2_2026_Demand":"'CAPACITY'!$C$6",
        "DualRun_FTE":"'CAPACITY'!$E$6",
        "Q2_2026_QA_Surplus":"'CAPACITY'!$I$6",
        "VendorA_QA_Region_Surplus":"'CAPACITY'!$J$6",
        "Q2_2026_Base_Vendor_Supply":"'CAPACITY'!$K$6",
        "DE_Gap_From_Q2":f"'CAPACITY'!$L$6:$L$16",
        "Q1_2026_DE_Role_Gap":"'CAPACITY'!$L$5",
        "DE_Before_M1":"'CAPACITY'!$C$31",
        "DE_After_M1":"'CAPACITY'!$C$32",
        "DE_After_VendorD":"'CAPACITY'!$C$33",
        "Cell_Quarter_Range":f"'CAPACITY'!$A${recon_start}:$A${recon_end}",
        "Cell_Role_Range":f"'CAPACITY'!$B${recon_start}:$B${recon_end}",
        "Cell_Gap_Range":f"'CAPACITY'!$F${recon_start}:$F${recon_end}",
        "Cell_Check_Result":f"'CAPACITY'!$H${recon_start}:$H${recon_end}",
        "Cell_Tolerance":f"'ASSUMPTIONS'!$C${a['cell_tolerance']}"
    }.items():
        define_name(name,ref)
    for ws in wb.worksheets:
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.alignment=Alignment(vertical="top",wrap_text=cell.alignment.wrap_text)
    term_codes = [
        [66,108,97,99,107,115,116,111,110,101], [66,88,84,105], [66,88],
        [68,101,108,111,105,116,116,101], [80,119,67], [70,105,100,101,108,105,116,121],
        [67,111,103,110,105,122,97,110,116], [69,81,84], [66,108,97,99,107,82,111,99,107],
        [68,46,69,46,32,83,104,97,119], [83,119,105,115,115,32,82,101],
        [74,80,77,111,114,103,97,110,67,104,97,115,101], [71,111,108,100,109,97,110,32,83,97,99,104,115],
        [67,105,116,105]
    ]
    forbidden = ["".join(map(chr, code)).casefold() for code in term_codes]
    text_values = [cell.value for sheet in wb.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value,str)]
    firm_hits = sum(text.casefold().count(term) for text in text_values for term in forbidden)
    word = "".join(map(chr,[97,99,116,117,97,108]))
    allowed = f"in-scenario {word} (simulated)"
    word_hits = sum(len(re.findall(r"\b"+word+r"\b", text.casefold().replace(allowed.casefold(),""))) for text in text_values)
    asm["K72"] = firm_hits
    asm["K73"] = word_hits
    OUTPUT.parent.mkdir(parents=True,exist_ok=True)
    wb.save(OUTPUT)
    return {"output":str(OUTPUT),"sheets":wb.sheetnames,"phasing":phases,"demand_rows":len(demand_data),"supply_rows":len(supply_data),"checks_row":checks_row,"all_row":allrow,"page_text_cells":page_text_cells,"roadmap_transition_cells":roadmap_transition_cells,"pdf_metadata_cells":pdf_metadata_cells}

if __name__=="__main__":
    print(json.dumps(build(),default=str))


