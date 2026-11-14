import csv
import html
import json
import math
import os
import re
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / ".b2deps"))
sys.path.insert(0, str(ROOT))

import matplotlib
matplotlib.use("Agg")
matplotlib.rcParams["font.family"]="DejaVu Sans"
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Patch
from matplotlib.ticker import MaxNLocator
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle

from generator import build_model

OUT_MODEL = ROOT / "out" / "delivery_review_model.xlsx"
OUT_RECALC = ROOT / "out" / "_recalc" / "delivery_review_model.xlsx"
OUT_PDF = ROOT / "out" / "delivery_review.pdf"
OUT_TRACE = ROOT / "out" / "page_trace.csv"
OUT_PREVIEW = ROOT / "out" / "_preview"
SOFFICE = shutil.which("soffice") or r"C:\Program Files\LibreOffice\program\soffice.exe"
BLUE = colors.HexColor("#193B5A")
PALE = colors.HexColor("#EAF0F5")
GRID = colors.HexColor("#B7C4CE")
INK = colors.HexColor("#172A3A")
MUTED = colors.HexColor("#607487")
GREEN = colors.HexColor("#2D765A")
AMBER = colors.HexColor("#B47716")
RED = colors.HexColor("#B44747")
PAGE_W, PAGE_H = landscape(A4)
MARGIN = 25
FOOTER_Y = 13
MIN_CONTENT_Y = 34
FONT_DIR = Path(matplotlib.get_data_path()) / "fonts" / "ttf"
FONT_REGULAR = "DejaVuSans"
FONT_BOLD = "DejaVuSans-Bold"
pdfmetrics.registerFont(TTFont(FONT_REGULAR, str(FONT_DIR / "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont(FONT_BOLD, str(FONT_DIR / "DejaVuSans-Bold.ttf")))
NUMBER_RE = re.compile(r"(?<![A-Za-z])(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?")
DATE_RE = re.compile(r"\b(?:\d{1,2}\s+[A-Z][a-z]{2}\s+\d{4}|[A-Z][a-z]{3}\s+\d{4}|Q[1-4]\s+\d{4})\b")


def run_recalc_pipeline():
    build_model.build()
    work = ROOT / "out" / "_recalc"
    source = work / "_source"
    ods_dir = work / "_ods"
    xlsx_dir = work / "_xlsx"
    profile = work / "_lo_profile"
    for folder in (source, ods_dir, xlsx_dir, profile):
        folder.mkdir(parents=True, exist_ok=True)
    for folder in (source, ods_dir, xlsx_dir):
        for path in folder.iterdir():
            if path.is_file():
                try:
                    path.unlink()
                except PermissionError:
                    os.chmod(path, 0o666)
                    path.unlink()
            elif path.is_dir():
                shutil.rmtree(path)
    shutil.copy2(OUT_MODEL, source / "model.xlsx")
    profile_uri = profile.resolve().as_uri()
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    # Direct xlsx -> xlsx conversion: LibreOffice computes every formula on load and keeps
    # explicit font colours. The old xlsx -> ods -> xlsx round trip reset all font colours
    # to theme "Text 1" (black), which made header text invisible in Excel.
    second = subprocess.run(
        [SOFFICE, f"-env:UserInstallation={profile_uri}", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(xlsx_dir), str(source / "model.xlsx")],
        capture_output=True, text=True, timeout=120, creationflags=creation_flags,
    )
    recalced = xlsx_dir / "model.xlsx"
    if second.returncode or not recalced.exists():
        raise RuntimeError(f"LibreOffice ods-to-xlsx recalc failed: {second.stdout}\n{second.stderr}")
    shutil.copy2(recalced, OUT_RECALC)
    shutil.copy2(recalced, OUT_MODEL)
    return OUT_RECALC


def raw_value_text(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        return format(value, ".15g")
    if value is None:
        return ""
    return str(value)


class Model:
    def __init__(self, path):
        self.values = load_workbook(path, data_only=True, read_only=False)
        self.formulas = load_workbook(path, data_only=False, read_only=False)
        self.a = self.values["ASSUMPTIONS"]
        self.assumption_rows = {self.a.cell(r, 1).value: r for r in range(1, self.a.max_row + 1) if self.a.cell(r, 1).value}
        self.page_text_rows = {self.a.cell(r, 1).value: r for r in range(1, self.a.max_row + 1) if self.a.cell(r, 1).value in build_model.KAKALI_TEXT["PAGE_TEXT"]}
        self.pdf_page_rows = {}
        for r in range(1, self.a.max_row + 1):
            label = self.a.cell(r, 1).value
            if isinstance(label, str) and label.startswith("PDF page "):
                self.pdf_page_rows[int(label.split()[-1])] = r
        self.meta_rows = {self.a.cell(r, 1).value: r for r in range(1, self.a.max_row + 1) if self.a.cell(r, 1).value in {"Document date", "Page count"}}

    def value(self, sheet, cell):
        return self.values[sheet][cell].value

    def formula(self, sheet, cell):
        return self.formulas[sheet][cell].value

    def assumption(self, key):
        return self.a.cell(self.assumption_rows[key], 3).value

    def assumption_cell(self, key):
        return f"C{self.assumption_rows[key]}"

    def assumption_label(self, label):
        for r in range(1, self.a.max_row + 1):
            if self.a.cell(r, 2).value == label:
                return self.a.cell(r, 3).value, f"ASSUMPTIONS!C{r}"
        raise KeyError(label)

    def page_text(self, key):
        row = self.page_text_rows[key]
        return self.a.cell(row, 2).value, f"ASSUMPTIONS!B{row}"

    def pdf_meta(self, key):
        row = self.meta_rows[key]
        return self.a.cell(row, 2).value, f"ASSUMPTIONS!B{row}"

    def pdf_page(self, page_number):
        row = self.pdf_page_rows[page_number]
        return self.a.cell(row, 2).value, f"ASSUMPTIONS!B{row}"

    def table_rows(self, sheet, table_name, data_only=True):
        ws = self.values[sheet] if data_only else self.formulas[sheet]
        ref = ws.tables[table_name].ref
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        return [
            [ws.cell(r, c).value for c in range(min_col, max_col + 1)]
            for r in range(min_row, max_row + 1)
        ], min_row, min_col


class PdfBuilder:
    def __init__(self, model):
        self.model = model
        self.trace = []
        self.canvas = canvas.Canvas(str(OUT_PDF), pagesize=landscape(A4), pageCompression=1)
        self.page_no = 0
        self.smallest_font = 99.0
        self.smallest_text_font = 99.0
        self.smallest_chart_font = 99.0
        self.page_titles = []
        self.layout_blocks = []
        self.text_fit_checks = 0
        self.word_fit_checks = 0
        self.chart_text_overlap_checks = 0
        self.chart_text_overlap_count = 0
        self.chart_text_min_clearance_pt = float("inf")
        self.chart_images_audited = 0
        self.chart_labels=[]

    def add_block(self, kind, x, bottom, width, height):
        self.layout_blocks.append((self.page_no,kind,float(x),float(bottom),float(width),float(height)))

    def draw_chart(self, image, x, y, width, height, font_size=7.5, dpi=180):
        image_reader=ImageReader(image)
        px_w,px_h=image_reader.getSize()
        scale=min(width/px_w,height/px_h)
        self.smallest_chart_font=min(self.smallest_chart_font,font_size*dpi/72*scale)
        placed_w,placed_h=px_w*scale,px_h*scale
        self.add_block("chart",x,y,width,height)
        self.canvas.drawImage(image_reader,x+(width-placed_w)/2,y+(height-placed_h)/2,
                              width=placed_w,height=placed_h,mask="auto")

    def validate_layout(self):
        for idx,left in enumerate(self.layout_blocks):
            p1,k1,x1,y1,w1,h1=left
            for right in self.layout_blocks[idx+1:]:
                p2,k2,x2,y2,w2,h2=right
                if p1!=p2:continue
                ix=max(0.0,min(x1+w1,x2+w2)-max(x1,x2))
                iy=max(0.0,min(y1+h1,y2+h2)-max(y1,y2))
                if ix>0.75 and iy>0.75:
                    raise ValueError(f"Page {p1}: {k1} {left[2:]} overlaps {k2} {right[2:]} by {ix:.1f} x {iy:.1f} pt")
        if self.text_fit_checks==0:
            raise ValueError("No text blocks were measured for fit")

    def record(self, page, exhibit, text, source, workbook_value):
        if text:
            if workbook_value is None or workbook_value=="":
                raise ValueError(f"Empty workbook source for printed figure {text!r}: {source}")
            self.trace.append([page, exhibit, str(text), source, raw_value_text(workbook_value)])

    def numeric(self, page, exhibit, sheet, cell, fmt=str):
        value = self.model.value(sheet, cell)
        text = fmt(value)
        self.record(page, exhibit, text, f"{sheet}!{cell}", value)
        return text

    def sourced_text(self, page, exhibit, sheet, cell, display=None):
        value = self.model.value(sheet, cell)
        text = str(value or "") if display is None else str(display)
        for match in NUMBER_RE.finditer(text):
            self.record(page, exhibit, match.group(0), f"{sheet}!{cell}", value)
        return text

    def sourced_date(self, page, exhibit, sheet, cell, fmt="%d %b %Y"):
        value = self.model.value(sheet, cell)
        text = value.strftime(fmt) if isinstance(value, (datetime, date)) else str(value)
        self.record(page, exhibit, text, f"{sheet}!{cell}", value)
        return text

    def table_cell(self, value, size=7.5, bold=False, color=INK, leading=None):
        self.smallest_font = min(self.smallest_font, size)
        self.smallest_text_font=min(self.smallest_text_font,size)
        style = ParagraphStyle(
            name=f"p_{size}_{bold}_{color.hexval()}", fontName=FONT_BOLD if bold else FONT_REGULAR,
            fontSize=size, leading=leading or size + 1.2, textColor=color, spaceAfter=0, spaceBefore=0,
            alignment=0, allowWidows=0, allowOrphans=0, splitLongWords=0,
        )
        return Paragraph(html.escape(str(value)).replace("\n", "<br/>"), style)

    def draw_table(self, data, x, top, widths, size=7.5, header_size=7.5, min_y=MIN_CONTENT_Y, header=True, paddings=2.0, row_background=True, border=True, repeat_rows=0, leading=None, highlights=None, spans=None):
        prepared=[]
        for ridx,row in enumerate(data):
            prepared.append([self.table_cell(value, header_size if header and ridx==0 else size, bold=header and ridx==0, color=colors.white if header and ridx==0 else INK, leading=leading) for value in row])
        table=Table(prepared,colWidths=widths,repeatRows=repeat_rows)
        commands=[
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),paddings),
            ("RIGHTPADDING",(0,0),(-1,-1),paddings),
            ("TOPPADDING",(0,0),(-1,-1),paddings),
            ("BOTTOMPADDING",(0,0),(-1,-1),paddings),
        ]
        if border:
            commands.append(("GRID",(0,0),(-1,-1),0.25,GRID))
        if header:
            commands.extend([("BACKGROUND",(0,0),(-1,0),BLUE), ("TEXTCOLOR",(0,0),(-1,0),colors.white)])
            if row_background:
                for r in range(1,len(data)):
                    if r%2==0: commands.append(("BACKGROUND",(0,r),(-1,r),PALE))
        for (row,col),fill in (highlights or {}).items():
            commands.append(("BACKGROUND",(col,row),(col,row),fill))
        for row,col_start,col_end in (spans or []):commands.append(("SPAN",(col_start,row),(col_end,row)))
        table.setStyle(TableStyle(commands))
        _,height=table.wrap(sum(widths),PAGE_H)
        self.text_fit_checks += sum(len(row) for row in prepared)
        for ri,row in enumerate(prepared):
            available=max(0,table._rowHeights[ri]-2*paddings)
            for ci,cell in enumerate(row):
                if not isinstance(cell,Paragraph):
                    continue
                span=next((end-start+1 for rr,start,end in (spans or []) if rr==ri and start==ci),1)
                cell_width=sum(widths[ci:ci+span])-2*paddings
                self.word_fit_checks += 1
                for word in re.findall(r"[^\s]+", re.sub(r"<br\s*/?>", " ", cell.getPlainText())):
                    if stringWidth(word,cell.style.fontName,cell.style.fontSize)>cell_width+0.5:
                        raise ValueError(f"Page {self.page_no}: unbreakable table word {word!r} exceeds column width {cell_width:.1f} pt")
                _,cell_height=cell.wrap(cell_width,PAGE_H)
                if cell_height>available+0.5:
                    raise ValueError(f"Page {self.page_no}: table text does not fit row {ri}")
        if top-height < min_y:
            raise ValueError(f"Page {self.page_no}: table would extend to y={top-height:.1f}, below {min_y:.1f}; table height {height:.1f}")
        table.drawOn(self.canvas,x,top-height)
        self.add_block("table",x,top-height,sum(widths),height)
        return top-height

    def section(self, x, y, exhibit, title_text):
        self.smallest_font=min(self.smallest_font,9.0)
        c=self.canvas
        c.setFillColor(BLUE)
        c.setFont(FONT_BOLD,9)
        c.drawString(x,y,f"{exhibit}  {title_text}")
        c.setStrokeColor(GRID)
        c.line(x,y-3,PAGE_W-MARGIN,y-3)
        self.add_block("heading",x,y-1,stringWidth(f"{exhibit}  {title_text}",FONT_BOLD,9),10)
        return y-12

    def draw_para(self, text, x, top, width, size=9.0, leading=None, min_y=MIN_CONTENT_Y, bold=False, color=INK):
        self.smallest_font=min(self.smallest_font,size)
        self.smallest_text_font=min(self.smallest_text_font,size)
        style=ParagraphStyle(name=f"body_{size}_{bold}",fontName=FONT_BOLD if bold else FONT_REGULAR,fontSize=size,
                             leading=leading or size*1.25,textColor=color,spaceAfter=0,spaceBefore=0)
        p=Paragraph(html.escape(str(text)).replace("\n","<br/>"),style)
        _,height=p.wrap(width,PAGE_H)
        self.text_fit_checks += 1
        if top-height<min_y:
            raise ValueError(f"Page {self.page_no}: paragraph would extend to y={top-height:.1f}, below {min_y:.1f}")
        p.drawOn(self.canvas,x,top-height)
        self.add_block("paragraph",x,top-height,width,height)
        return top-height

    def begin_page(self, title_text, subtitle=None):
        self.page_no += 1
        self.page_titles.append(title_text)
        label_match=re.match(r"^(\d+)\s+\|",title_text)
        if label_match:
            page_index,page_source=self.model.pdf_page(self.page_no)
            self.record(self.page_no,"Page title",str(page_index),page_source,page_index)
        c=self.canvas
        c.setFillColor(BLUE)
        c.rect(0,PAGE_H-51,PAGE_W,51,stroke=0,fill=1)
        c.setFillColor(colors.white)
        c.setFont(FONT_BOLD,17)
        c.drawString(MARGIN,PAGE_H-31,title_text)
        if subtitle:
            c.setFont(FONT_REGULAR,9)
            c.drawRightString(PAGE_W-MARGIN,PAGE_H-31,subtitle)
        return PAGE_H-65

    def footer(self, page_number):
        c=self.canvas
        as_at_value,as_at_source=self.model.assumption_label("As-at date")
        doc_date,doc_source=self.model.pdf_meta("Document date")
        page_index,page_source=self.model.pdf_page(page_number)
        page_count,count_source=self.model.pdf_meta("Page count")
        as_at=as_at_value.strftime("%d %b %Y")
        doc=doc_date.strftime("%d %b %Y")
        self.record(page_number,"Footer",as_at,as_at_source,as_at_value)
        self.record(page_number,"Footer",doc,doc_source,doc_date)
        self.record(page_number,"Footer",str(page_index),page_source,page_index)
        self.record(page_number,"Footer",str(page_count),count_source,page_count)
        footer=f"Simulated data – not observed results · as at {as_at} · dated {doc} · page {page_index} of {page_count}"
        c.setStrokeColor(GRID)
        c.line(MARGIN,FOOTER_Y+9,PAGE_W-MARGIN,FOOTER_Y+9)
        c.setFillColor(MUTED)
        c.setFont(FONT_REGULAR,7.5)
        c.drawString(MARGIN,FOOTER_Y,footer)
        self.smallest_font=min(self.smallest_font,7.5)
        self.smallest_text_font=min(self.smallest_text_font,7.5)

    def exhibit(self, x, y, code, label):
        self.canvas.setFont(FONT_BOLD,8.5)
        self.canvas.setFillColor(BLUE)
        self.canvas.drawString(x,y,f"{code}  {label}")
        self.add_block("heading",x,y-1,stringWidth(f"{code}  {label}",FONT_BOLD,8.5),10)
        return y-11

    def forced_box(self, page, decision_row, min_y=MIN_CONTENT_Y, top=80, x=MARGIN, widths=None):
        d=self.model.values["DECISIONS"]
        data=[["Decision forced", "D-number", "Date", "Decision", "Decider"]]
        items=[]
        for col,fmt in ((1,str),(2,str),(3,str),(6,str)):
            val=d.cell(decision_row,col).value
            items.append((col,val))
        did=self.sourced_text(page,"Decision forced","DECISIONS",f"A{decision_row}")
        date_text=self.sourced_text(page,"Decision forced","DECISIONS",f"B{decision_row}")
        decision=self.sourced_text(page,"Decision forced","DECISIONS",f"C{decision_row}")
        decider=self.sourced_text(page,"Decision forced","DECISIONS",f"F{decision_row}")
        data.append(["Decision forced",did,date_text,decision,decider])
        return self.draw_table(data,x,top,widths or [76,55,62,489,110],size=7.5,header_size=7.5,min_y=min_y,paddings=1.3)

    def save_trace(self):
        OUT_TRACE.parent.mkdir(parents=True,exist_ok=True)
        with OUT_TRACE.open("w",newline="",encoding="utf-8") as f:
            writer=csv.writer(f)
            writer.writerow(["page","exhibit","text as printed","tab!cell","workbook value"])
            writer.writerows(self.trace)


def loc_data_source(path):
    model=Model(path)
    return model


def make_roadmap_chart(model, tracer):
    portfolio=model.values["PORTFOLIO"]
    capacity=model.values["CAPACITY"]
    quarters=[capacity.cell(r,1).value for r in range(5,17)]
    for col in range(5,17):tracer.sourced_text(2,"E1.2","CAPACITY",f"A{col}")
    funded=[]
    for rr in range(5,18):
        if not str(portfolio.cell(rr,15).value or "").startswith("Fund"):continue
        line=portfolio.cell(rr,1).value
        label=f"{line}  {portfolio.cell(rr,2).value}"
        tracer.sourced_text(2,"E1.2","PORTFOLIO",f"A{rr}")
        tracer.sourced_text(2,"E1.2","PORTFOLIO",f"B{rr}")
        funded.append((rr,label))
    fig,ax=plt.subplots(figsize=(10.4,1.30),dpi=180,facecolor="white")
    ax.set_facecolor("#F8FAFC")
    forecast_x=3+1/3
    ax.axvspan(forecast_x,11.5,color="#F5E9B6",alpha=.68,zorder=0)
    line3=str(portfolio["B35"].value);line6=str(portfolio["B36"].value)
    line3_id=str(portfolio["B35"].value);line6_id=str(portfolio["B36"].value)
    # The transition cells are the workbook source for the collision/re-plan annotations.
    for addr in ("B34","B35","B36","B37","B38","B39","B40","B41"):
        tracer.sourced_text(2,"E1.2","PORTFOLIO",addr)
    original=str(portfolio["B39"].value);moved=str(portfolio["B40"].value);revised=str(portfolio["B41"].value)
    shared_role=str(portfolio["B37"].value);shared_fte=portfolio["B38"].value
    for idx,(rr,label) in enumerate(funded):
        y=len(funded)-1-idx
        for year_offset,col in enumerate((9,10,11)):
            if float(portfolio.cell(rr,col).value or 0)<=0:continue
            start=year_offset*4
            stop=start+4
            if str(portfolio.cell(rr,1).value)==line6_id and year_offset==0:
                start=2
            ax.broken_barh([(start,stop-start)],(y-.30,.60),facecolors=("#2B6A8A" if year_offset==0 else "#3B8066" if year_offset==1 else "#AD7420"),edgecolors="white",linewidth=.7)
    y_lookup={str(portfolio.cell(rr,1).value):len(funded)-1-idx for idx,(rr,_) in enumerate(funded)}
    collision_x=1
    for line_id in (line3_id,line6_id):
        if line_id in y_lookup:
            ax.scatter(collision_x,y_lookup[line_id],s=48,facecolors="none",edgecolors="#B44747",linewidth=1.3,zorder=4)
    ax.axvline(collision_x,color="#B44747",linewidth=.8,linestyle="--",zorder=2)
    tracer.sourced_text(2,"E1.2","ASSUMPTIONS","C6")
    ax.set_yticks(range(len(funded)))
    ax.set_yticklabels([label for _,label in reversed(funded)],fontsize=8.5)
    ax.set_xticks(range(12))
    ax.set_xticklabels([q.replace(" ","\n") for q in quarters],fontsize=8.5)
    ax.set_xlim(-.5,11.5);ax.set_ylim(-.7,len(funded)-.1)
    ax.grid(axis="x",color="#C9D4DD",linewidth=.55,zorder=1)
    ax.tick_params(axis="both",length=0,pad=2)
    ax.set_axisbelow(True)
    for spine in ax.spines.values():spine.set_visible(False)
    years=[str(capacity.cell(r,1).value).split()[1] for r in (5,9,13)]
    handles=[Patch(facecolor=color,edgecolor="white",label=f"{year} funded") for color,year in zip(("#2B6A8A","#3B8066","#AD7420"),years)]
    handles.append(Patch(facecolor="#F5E9B6",edgecolor="none",label=f"Forecast after {model.value('ASSUMPTIONS','C6').strftime('%d %b %Y')}"))
    ax.legend(handles=handles,loc="lower center",bbox_to_anchor=(.5,1.01),ncol=4,frameon=False,fontsize=8.5,handlelength=1.2,columnspacing=1.4)
    fig.subplots_adjust(left=.28,right=.99,top=.95,bottom=.05)
    buf=io_bytes();fig.savefig(buf,format="png",dpi=180,bbox_inches="tight",pad_inches=.05);buf.seek(0)
    px_w,px_h=ImageReader(buf).getSize()
    scale=min((PAGE_W-2*MARGIN)/px_w,116/px_h)
    n_checks,n_overlaps,min_clearance=audit_chart_text(fig,scale)
    tracer.chart_text_overlap_checks+=n_checks
    tracer.chart_text_overlap_count+=n_overlaps
    tracer.chart_text_min_clearance_pt=min(tracer.chart_text_min_clearance_pt,min_clearance)
    tracer.chart_images_audited+=1
    tracer.chart_labels.append({"roadmap_legend":[handle.get_label() for handle in handles],
                                "roadmap_y_tick_labels":[str(t.get_text()) for t in ax.get_yticklabels() if t.get_visible()],
                                "roadmap_min_clearance_pt":min_clearance})
    plt.close(fig);buf.seek(0)
    return buf


def io_bytes():
    from io import BytesIO
    return BytesIO()


def audit_chart_text(fig,scale=1.0):
    fig.canvas.draw()
    renderer=fig.canvas.get_renderer()
    entries=[]
    for axis_index,ax in enumerate(fig.axes):
        candidates=[ax.title,ax.xaxis.label,ax.yaxis.label,*ax.get_xticklabels(),*ax.get_yticklabels(),*ax.texts] if ax.axison else []
        if ax.get_legend():candidates.extend(ax.get_legend().get_texts())
        for artist in candidates:
            if artist.get_visible() and str(artist.get_text()).strip():
                bbox=artist.get_window_extent(renderer)
                if bbox.width>0 and bbox.height>0:entries.append((str(artist.get_text()),bbox,axis_index,ax))
    for legend in fig.legends:
        for artist in legend.get_texts():
            if artist.get_visible() and str(artist.get_text()).strip():entries.append((str(artist.get_text()),artist.get_window_extent(renderer),-1,None))
    close_pairs=[]
    min_clearance=float("inf")
    for i,(left_text,left,left_axis,left_owner) in enumerate(entries):
        for right_text,right,right_axis,right_owner in entries[i+1:]:
            gap_x=max(left.x0-right.x1,right.x0-left.x1,0)
            gap_y=max(left.y0-right.y1,right.y0-left.y1,0)
            clearance=math.hypot(gap_x,gap_y)*scale
            min_clearance=min(min_clearance,clearance)
            if clearance<1.0:
                close_pairs.append((left_text,right_text,left_axis,right_axis,clearance,left.bounds,right.bounds))
    if close_pairs:
        raise ValueError(f"Chart text boxes have less than 1 pt clearance: {close_pairs[:4]}")
    return len(entries)*(len(entries)-1)//2,len(close_pairs),min_clearance


def make_capacity_chart(model, tracer):
    cap=model.values["CAPACITY"]
    role_table,role_head,role_min=model.table_rows("CAPACITY","RoleRegionQuarterCheck")
    role_col={name:i for i,name in enumerate(role_table[0])}
    regions=sorted({row[role_col["Region"]] for row in role_table[1:]})
    roles=[]
    for row in role_table[1:]:
        if row[role_col["Role"]] not in roles:roles.append(row[role_col["Role"]])
    quarters=[cap.cell(r,1).value for r in range(5,17)]
    for r in range(5,17):tracer.sourced_text(4,"E3.1","CAPACITY",f"A{r}")
    role_idx=role_col["Role"];region_idx=role_col["Region"];quarter_idx=role_col["Quarter"];gap_idx=role_col["Gap FTE"]
    gaps={reg:{(role,q):0 for role in roles for q in quarters} for reg in regions}
    for r,row in enumerate(role_table[1:],role_head+1):
        role,region,quarter=row[role_idx],row[region_idx],row[quarter_idx]
        gaps[region][(role,quarter)]=row[gap_idx]
        gap_cell=f"{chr(64+role_min+gap_idx)}{r}"
        tracer.numeric(4,"E3.1","CAPACITY",gap_cell,lambda x:format(float(x),".0f"))
        tracer.sourced_text(4,"E3.1","CAPACITY",f"B{r}")
    fig=plt.figure(figsize=(12.0,2.8),dpi=180,facecolor="white",layout="constrained")
    grid=fig.add_gridspec(3,5,width_ratios=[1.2,1,1,1,1],height_ratios=[3.0,.12,.45],hspace=0,wspace=.05)
    norm=plt.Normalize(-3,3);palette=plt.get_cmap("RdYlGn")
    heat_axes=[]
    role_ax=fig.add_subplot(grid[0,0])
    role_ax.set_ylim(2*len(roles)-1,-.5);role_ax.set_yticks([]);role_ax.set_xticks([])
    for role_index,role in enumerate(roles):
        role_ax.text(.80,2*role_index,str(role),transform=role_ax.get_yaxis_transform(),ha="right",va="center",fontsize=9.0)
    for spine in role_ax.spines.values():spine.set_visible(False)
    for gi,region in enumerate(regions):
        ax=fig.add_subplot(grid[0,gi+1],sharey=heat_axes[0] if heat_axes else None)
        heat_axes.append(ax)
        values=np.array([[gaps[region][(role,q)] for q in quarters] for role in roles],dtype=float)
        data=np.full((2*len(roles)-1,len(quarters)),np.nan);data[::2]=values
        heat_palette=palette.copy();heat_palette.set_bad("white")
        ax.imshow(data,aspect="auto",cmap=heat_palette,norm=norm,interpolation="nearest")
        ax.set_title(region,fontsize=10.0,pad=-3)
        ax.set_ylim(2*len(roles)-1,-.5)
        ax.set_yticks(range(0,2*len(roles),2));ax.tick_params(axis="y",labelleft=False,left=False)
        ax.set_xticks(range(12));ax.set_xticklabels([f"{q.split()[0][1:]}\n{q.split()[1][-2:]}" if i%4==0 else q.split()[0][1:] for i,q in enumerate(quarters)],fontsize=8.8)
        ax.tick_params(length=0,pad=1)
        for spine in ax.spines.values():spine.set_visible(False)
    key_ax=fig.add_subplot(grid[1,1:]);key_ax.axis("off")
    gap_handles=[Patch(facecolor=palette(norm(v)),edgecolor="none",label=label) for v,label in ((-2,"Shortfall"),(0,"Within tolerance"),(2,"Surplus"))]
    ramp,ramp_head,ramp_min=model.table_rows("CAPACITY","VendorRampByVendor")
    vendor_colors=plt.rcParams["axes.prop_cycle"].by_key()["color"]
    key_ax.legend(handles=gap_handles,loc="center",ncol=3,frameon=False,fontsize=9.5,handlelength=1.2,columnspacing=1.5)
    panel_axes=[]
    ax=fig.add_subplot(grid[2,1:3]);panel_axes.append(ax)
    for r,row in enumerate(ramp[1:],ramp_head+1):
        ax.plot(range(12),row[1:],marker="o",markersize=2.4,linewidth=1.1,color=vendor_colors[(r-ramp_head-1)%len(vendor_colors)])
        tracer.sourced_text(4,"E3.1","CAPACITY",f"{chr(64+ramp_min)}{r}")
        for i in range(12):tracer.numeric(4,"E3.1","CAPACITY",f"{chr(65+ramp_min+i)}{r}",lambda x:format(float(x),".0f"))
    ax.set_title("Ramp · vendor supply",fontsize=10.0,pad=3);ax.set_ylabel("FTE",fontsize=9.5)
    ax.set_ylim(bottom=0);ax.yaxis.set_major_locator(MaxNLocator(nbins=3,integer=True))
    bench,bench_head,bench_min=model.table_rows("CAPACITY","BenchByQuarter")
    bench_ax=fig.add_subplot(grid[2,3]);panel_axes.append(bench_ax)
    bench_ax.bar(range(12),[row[1] for row in bench[1:]],color="#5C8FA7",width=.72)
    bench_ax.set_title("Bench · surplus",fontsize=10.0,pad=3);bench_ax.set_ylabel("FTE",fontsize=9.5)
    bench_ax.set_ylim(bottom=0);bench_ax.yaxis.set_major_locator(MaxNLocator(nbins=3,integer=True))
    for r,row in enumerate(bench[1:],bench_head+1):tracer.numeric(4,"E3.1","CAPACITY",f"{chr(64+bench_min+1)}{r}",lambda x:format(float(x),".0f"))
    utilisation,util_head,util_min=model.table_rows("CAPACITY","UtilisationByQuarter")
    ux=fig.add_subplot(grid[2,4]);panel_axes.append(ux)
    ux.bar(range(12),[row[1] for row in utilisation[1:]],color="#9A7B4F",width=.72)
    ux.set_title("Utilisation · demand / supply",fontsize=10.0,pad=3);ux.set_ylabel("Ratio",fontsize=9.5)
    ux.set_yticks([0,.5,1.0,1.5]);ux.set_ylim(0,1.6);ux.axhline(1.0,color="#B44747",linewidth=.9,linestyle="--",zorder=3)
    for r,row in enumerate(utilisation[1:],util_head+1):tracer.numeric(4,"E3.1","CAPACITY",f"{chr(64+util_min+1)}{r}",lambda x:f"{float(x):.0%}")
    for ax in heat_axes+panel_axes:
        ax.set_xticks(range(12));ax.set_xticklabels([f"{q.split()[0][1:]}\n{q.split()[1][-2:]}" if i%4==0 else q.split()[0][1:] for i,q in enumerate(quarters)],fontsize=9.0)
        ax.tick_params(axis="x",length=0,pad=1)
        ax.grid(axis="y",alpha=.22)
        for spine in ax.spines.values():spine.set_visible(False)
    role_ax.set_position([.01,.38,.25,.54])
    for gi,ax in enumerate(heat_axes):ax.set_position([.27+gi*.18,.38,.18,.54])
    key_ax.set_position([.01,.95,.25,.04])
    for ax,xpos,width in ((panel_axes[0],.27,.36),(panel_axes[1],.65,.16),(panel_axes[2],.83,.16)):
        ax.set_position([xpos,.07,width,.14])
    for ax in panel_axes:ax.yaxis.set_major_locator(MaxNLocator(nbins=2,integer=True))
    ux.set_yticks([0,1.0])
    tracer.chart_labels.append({"roles":roles,"quarters":quarters,"regions":regions})
    buf=io_bytes();fig.savefig(buf,format="png",dpi=180,bbox_inches="tight",pad_inches=.05);buf.seek(0)
    px_w,px_h=ImageReader(buf).getSize()
    scale=min((PAGE_W-2*MARGIN)/px_w,182/px_h)
    n_checks,n_overlaps,min_clearance=audit_chart_text(fig,scale)
    tracer.chart_text_overlap_checks+=n_checks
    tracer.chart_text_overlap_count+=n_overlaps
    tracer.chart_text_min_clearance_pt=min(tracer.chart_text_min_clearance_pt,min_clearance)
    tracer.chart_images_audited+=1
    tracer.chart_labels[-1]["capacity_chart_min_clearance_pt"]=min_clearance
    plt.close(fig);buf.seek(0)
    return buf

def number_sources(builder, page, exhibit, text, source, value):
    """Record every printed numeric token against its workbook origin."""
    for match in NUMBER_RE.finditer(str(text)):
        builder.record(page, exhibit, match.group(0), source, value)


def signed_number(value):
    return f"−{abs(value):g}" if value < 0 else f"{value:g}"


def model_cell(model, sheet, cell):
    return model.value(sheet, cell)


def cell_text(model, builder, page, exhibit, sheet, cell, fmt=None):
    value=model.value(sheet,cell)
    if fmt:
        text=fmt(value)
    elif isinstance(value,(date,datetime)):
        text=value.strftime("%d %b %Y")
    else:
        text="" if value is None else str(value)
    number_sources(builder,page,exhibit,text,f"{sheet}!{cell}",value)
    return text


def draw_table(builder, page, exhibit, x, top, rows, widths, size=7.5, sources=None,
               header=True, padding=1.5, min_y=MIN_CONTENT_Y, leading=None, highlights=None, spans=None):
    # Use the common table renderer; trace numeric tokens from the corresponding source cells.
    for ri,row in enumerate(rows):
        for ci,text in enumerate(row):
            if sources and (ri,ci) in sources:
                sheet,cell=sources[(ri,ci)]
                number_sources(builder,page,exhibit,text,f"{sheet}!{cell}",model_cell(builder.model,sheet,cell))
    return builder.draw_table(rows,x,top,widths,size=size,header_size=size,min_y=min_y,
                              paddings=padding,repeat_rows=1 if header else 0,header=header,leading=leading,highlights=highlights,spans=spans)


def draw_spend_exhibit(model,builder,page,top):
    commercial=model.values["COMMERCIAL"]
    assumptions=model.values["ASSUMPTIONS"]
    head=builder.exhibit(MARGIN,top,"E3.4","Vendor spend by year · $m; funded; share")
    spends=[["Vendor","2026 $m","2027 $m","2028 $m","3-year share"]]
    sources={}
    for rr in range(66,72):
        spends.append([commercial.cell(rr,1).value]+[f"${commercial.cell(rr,col).value:.1f}" for col in (2,3,4)]+[f"{commercial.cell(rr,6).value:.1%}"])
        for ci,col in enumerate((1,2,3,4,6)):sources[(len(spends)-1,ci)]=( "COMMERCIAL",f"{chr(64+col)}{rr}")
    for ci,col in enumerate((2,3,4),1):sources[(0,ci)]=( "COMMERCIAL",f"{chr(64+col)}65")
    scaling_label=build_model.PARAMS["commercial"]["rate_scaling"]["label"]
    factor_years=[commercial.cell(57,col).value for col in (2,3,4)]
    factor_values=[commercial.cell(60,col).value for col in (2,3,4)]
    factor_text=scaling_label+": "+"; ".join(f"{factor_years[i]} {factor_values[i]:.4f}" for i in range(3))+"."
    fy_text=(f"{commercial['B57'].value} vendor spend: ${commercial['B73'].value:.1f}m in-scenario actual (simulated) + "
             f"${commercial['C73'].value:.1f}m forecast (simulated); funded profile: "
             f"${assumptions['C12'].value:.1f}m / ${assumptions['C13'].value:.1f}m / ${assumptions['C14'].value:.1f}m. "
             "Three vendor shares exceed the concentration threshold.")
    spends.append([factor_text,"","","",""])
    spans=[(len(spends)-1,0,4)]
    spends.append([fy_text,"","","",""])
    spans.append((len(spends)-1,0,4))
    bottom=draw_table(builder,page,"E3.4",MARGIN,head,spends,[140,130,130,130,262],size=7.5,sources=sources,padding=.1,min_y=55,leading=7.5,spans=spans)
    for col in (2,3,4):
        builder.record(page,"E3.4",str(commercial.cell(57,col).value),f"COMMERCIAL!{chr(64+col)}57",commercial.cell(57,col).value)
        builder.record(page,"E3.4",f"{commercial.cell(60,col).value:.4f}",f"COMMERCIAL!{chr(64+col)}60",commercial.cell(60,col).value)
    for shown,source,value in ((str(commercial["B57"].value),"COMMERCIAL!B57",commercial["B57"].value),
                               (f"{commercial['B73'].value:.1f}","COMMERCIAL!B73",commercial["B73"].value),
                               (f"{commercial['C73'].value:.1f}","COMMERCIAL!C73",commercial["C73"].value),
                               (f"{assumptions['C12'].value:.1f}","ASSUMPTIONS!C12",assumptions["C12"].value),
                               (f"{assumptions['C13'].value:.1f}","ASSUMPTIONS!C13",assumptions["C13"].value),
                               (f"{assumptions['C14'].value:.1f}","ASSUMPTIONS!C14",assumptions["C14"].value)):
        number_sources(builder,page,"E3.4",shown,source,value)
    return bottom


def make_pdf(model):
    OUT_PREVIEW.mkdir(parents=True,exist_ok=True)
    builder=PdfBuilder(model)
    c=builder.canvas
    assumptions=model.values["ASSUMPTIONS"]
    portfolio=model.values["PORTFOLIO"]
    plan=model.values["PLAN"]
    capacity=model.values["CAPACITY"]
    benefits=model.values["BENEFITS"]
    vendors=model.values["VENDORS"]
    commercial=model.values["COMMERCIAL"]
    decisions=model.values["DECISIONS"]

    # Cover. Kakali's purpose is read only from its named workbook text cell.
    builder.page_no=1
    builder.page_titles.append("Cover")
    c.setFillColor(BLUE); c.rect(0,0,PAGE_W,PAGE_H,stroke=0,fill=1)
    c.setFillColor(colors.white); c.setFont(FONT_BOLD,24)
    c.drawString(46,PAGE_H-112,"Enterprise Transformation Delivery Review")
    purpose,purpose_source=model.page_text("cover_purpose")
    builder.draw_para(purpose,46,PAGE_H-153,PAGE_W-92,size=12,leading=16,min_y=PAGE_H-235,color=colors.white)
    c.setFont(FONT_BOLD,12); c.drawString(46,PAGE_H-278,"Simulated data – not observed results")
    asat,asat_source=model.assumption_label("As-at date")
    doc,doc_source=model.pdf_meta("Document date")
    asat_text=asat.strftime("%d %b %Y"); doc_text=doc.strftime("%d %b %Y")
    c.setFont(FONT_REGULAR,10); c.drawString(46,PAGE_H-303,f"As at {asat_text}  ·  dated {doc_text}")
    number_sources(builder,1,"Cover",purpose,purpose_source,purpose)
    number_sources(builder,1,"Cover",asat_text,asat_source,asat)
    number_sources(builder,1,"Cover",doc_text,doc_source,doc)
    c.setFont(FONT_BOLD,10); c.drawString(46,PAGE_H-355,"Contents")
    contents=["Cover","Portfolio","Governance","Vendor workforce","Is it working","Cadence","Appendix – decision log"]
    for ix,item in enumerate(contents): c.setFont(FONT_REGULAR,10); c.drawString(62,PAGE_H-378-ix*19,item)
    builder.footer(1); c.showPage()

    # Page 1 – portfolio register, roadmap/collision, and deferrals.
    top=builder.begin_page("1 | Portfolio","Can the portfolio fit capacity and funding?")
    y=builder.exhibit(MARGIN,top,"E1.1","Portfolio register · $m")
    rows=[ ["#","Initiative","BU","Region","V/F","Y1","Y2","Y3","Outside","Owner","Decision"] ]
    sources={}
    for rr in range(5,18):
        vals=[portfolio.cell(rr,1).value,portfolio.cell(rr,2).value,portfolio.cell(rr,3).value,portfolio.cell(rr,4).value,
              f"{portfolio.cell(rr,5).value}/{portfolio.cell(rr,6).value}"]
        vals += [f"{portfolio.cell(rr,col).value:.1f}" for col in (9,10,11)]
        decision_text=str(portfolio.cell(rr,15).value or "")
        decision_short=decision_text.split(";")[0].split(" - ")[0]
        vals += [f"{portfolio.cell(rr,12).value:.1f}",portfolio.cell(rr,14).value,decision_short]
        rows.append(vals)
        for ci,col in enumerate((1,2,3,4,5,9,10,11,12,14,15)):
            sources[(len(rows)-1,ci)]=( "PORTFOLIO",f"{chr(64+col)}{rr}")
    total=["Total","","","","", *[f"{portfolio.cell(18,col).value:.1f}" for col in (9,10,11,12)],""]
    rows.append(total)
    for ci,col in enumerate((9,10,11,12)):
        sources[(len(rows)-1,5+ci)]=( "PORTFOLIO",f"{chr(64+col)}18")
    y=draw_table(builder,2,"E1.1",MARGIN,y,rows,[22,165,65,50,34,44,44,44,50,95,179],size=7.5,sources=sources,padding=.3,leading=7.9)
    summary_cells=("C21","C22","C23","C24")
    summary_labels=("Total ask $m","Funded $m","Deferred $m","Stopped $m")
    summary_values=[f"${portfolio[cell].value:.1f}" for cell in summary_cells]
    summary_sources={(1,col):("PORTFOLIO",cell) for col,cell in enumerate(summary_cells)}
    y=draw_table(builder,2,"E1.1",MARGIN,y-2,[list(summary_labels),summary_values],[185,185,185,185],size=7.5,sources=summary_sources,padding=.3,leading=8.2)
    y-=10
    y=builder.exhibit(MARGIN,y,"E1.2","Roadmap and shared data-engineering capacity")
    first_quarter=capacity["A5"].value;last_quarter=capacity["A16"].value;asat_date=model.value("ASSUMPTIONS","C6")
    line3=portfolio["B35"].value;line6=portfolio["B36"].value;shared_role=portfolio["B37"].value;shared_fte=portfolio["B38"].value
    original_need=portfolio["B39"].value;replan_quarter=portfolio["B40"].value;revised_need=portfolio["B41"].value
    fte_words={1:"one",2:"two",3:"three",4:"four",5:"five",6:"six",7:"seven",8:"eight",9:"nine",10:"ten"}
    fte_word=fte_words.get(int(shared_fte),str(shared_fte))
    role_word=str(shared_role).lower().replace(" engineering", "")
    collision_text=(f"Q2: lines {line3} and {line6} need the same {fte_word} {role_word} engineers; "
                    f"after D4, line {line6} moves to {replan_quarter}, requirement {original_need:g} to {revised_need:g}. "
                    f"Forecast after {asat_date.strftime('%d %b %Y')}.")
    y=builder.draw_para(collision_text,MARGIN,y,PAGE_W-2*MARGIN,size=9,leading=10,min_y=MIN_CONTENT_Y)
    for shown,source,value in ((first_quarter,"CAPACITY!A5",first_quarter),(last_quarter,"CAPACITY!A16",last_quarter),
                               (asat_date.strftime("%d %b %Y"),"ASSUMPTIONS!C6",asat_date),(line3,"PORTFOLIO!B35",line3),
                               (line6,"PORTFOLIO!B36",line6),
                               (shared_role,"PORTFOLIO!B37",shared_role),(replan_quarter,"PORTFOLIO!B40",replan_quarter),
                               (original_need,"PORTFOLIO!B39",original_need),(revised_need,"PORTFOLIO!B41",revised_need)):
        number_sources(builder,2,"E1.2",shown,source,value)
    builder.record(2,"E1.2",fte_word,"PORTFOLIO!B38",shared_fte)
    chart=make_roadmap_chart(model,builder)
    builder.draw_chart(chart,MARGIN,y-116,PAGE_W-2*MARGIN,116,font_size=8.5)
    y-=127
    y=builder.exhibit(MARGIN,y,"E1.3","Deferrals and stops · D1 CRM stop rationale · reason, re-entry trigger, date")
    builder.record(2,"E1.3",decisions.cell(5,1).value,"DECISIONS!A5",decisions.cell(5,1).value)
    defer=[]
    for rr in range(5,18):
        reason=portfolio.cell(rr,17).value; trigger=portfolio.cell(rr,16).value
        disposition=str(portfolio.cell(rr,15).value or "")
        if ("Stop" in disposition or "Defer" in disposition) and (reason or trigger):
            trigger_text=str(trigger or "")
            if any(y in trigger_text for y in ("2027","2028","2029")) and "forecast" not in trigger_text.casefold():trigger_text += " (forecast)"
            defer.append([f"{portfolio.cell(rr,1).value} · {portfolio.cell(rr,2).value}",
                          f"Product reason: {reason or ''}  Re-entry trigger / date: {trigger_text}"])
    if defer:
        selected=[r for r in range(5,18) if ("Stop" in str(portfolio.cell(r,15).value or "") or "Defer" in str(portfolio.cell(r,15).value or "")) and (portfolio.cell(r,16).value or portfolio.cell(r,17).value)]
        sources={(i+1,0):( "PORTFOLIO",f"A{rr}") for i,rr in enumerate(selected)}
        for rr in selected:
            number_sources(builder,2,"E1.3",portfolio.cell(rr,17).value,f"PORTFOLIO!Q{rr}",portfolio.cell(rr,17).value)
            number_sources(builder,2,"E1.3",portfolio.cell(rr,16).value,f"PORTFOLIO!P{rr}",portfolio.cell(rr,16).value)
        y=draw_table(builder,2,"E1.3",MARGIN,y,[["Line · initiative","Product reason · re-entry trigger / date"]]+defer,[185,600],size=7.5,sources=sources,padding=.2,leading=7.9,min_y=MIN_CONTENT_Y+44)
    builder.forced_box(2,6,min_y=34,top=70)
    builder.footer(2); c.showPage()

    # Page 2 – integrated governance.
    top=builder.begin_page("2 | Governance","§6 delivery control")
    y=builder.exhibit(MARGIN,top,"E2.1","Three workstreams · 14 milestones · critical path and float")
    ms=[ ["ID","WS","Date","Milestone / deliverable","Status"] ]
    src={}
    for rr in range(5,19):
        ms.append([cell_text(model,builder,3,"E2.1","PLAN",f"A{rr}"),cell_text(model,builder,3,"E2.1","PLAN",f"B{rr}"),cell_text(model,builder,3,"E2.1","PLAN",f"C{rr}"),cell_text(model,builder,3,"E2.1","PLAN",f"D{rr}"),cell_text(model,builder,3,"E2.1","PLAN",f"G{rr}")])
        for ci,col in enumerate((1,2,3,4,7)):src[(len(ms)-1,ci)]=( "PLAN",f"{chr(64+col)}{rr}")
    number_sources(builder,3,"E2.1","14","PLAN!A18",plan.cell(18,1).value)
    y=draw_table(builder,3,"E2.1",MARGIN,y,ms,[31,32,76,555,98],size=7.5,sources=src,padding=1.0)
    y-=5
    path_value=str(model.value("PLAN","B21"))
    nodes=[part.strip() for part in path_value.split("→")]
    if len(nodes)!=5:raise ValueError(f"Expected five critical-path nodes from PLAN!B21, got {len(nodes)}")
    path_row=[]
    path_widths=[]
    path_sources={}
    for idx,node in enumerate(nodes):
        path_row.append(node);path_widths.append((110,112,172,168,112)[idx]);path_sources[(0,len(path_row)-1)]=( "PLAN","B21")
        if idx<len(nodes)-1:path_row.append("→");path_widths.append((15,15,15,15)[idx])
    for node in nodes:number_sources(builder,3,"E2.1",node,"PLAN!B21",path_value)
    y=draw_table(builder,3,"E2.1",MARGIN,y,[path_row],path_widths,size=7.5,header=False,sources=path_sources,padding=3.0,leading=8.5)
    float_value=model.value("PLAN","B23")
    float_text=f"Float: {float_value:g} working days"
    y=builder.draw_para(float_text,MARGIN,y-3,PAGE_W-50,size=9,min_y=MIN_CONTENT_Y)
    number_sources(builder,3,"E2.1",f"{float_value:g} working days","PLAN!B23",float_value)
    # governance bottom split
    y-=14
    leftx=MARGIN; colw=388; rightx=PAGE_W-MARGIN-colw
    yl=builder.exhibit(leftx,y,"E2.2","DACI · six decisions")
    drows=[["Decision type","Driver","Approver"]]
    dsrc={}
    for rr in range(27,33):
        drows.append([plan.cell(rr,1).value,plan.cell(rr,2).value,plan.cell(rr,3).value])
        for ci,col in enumerate((1,2,3)):dsrc[(len(drows)-1,ci)]=( "PLAN",f"{chr(64+col)}{rr}")
    yl=draw_table(builder,3,"E2.2",leftx,yl,drows,[177,105,106],size=7.5,sources=dsrc,padding=1.1)
    yr=builder.exhibit(rightx,y,"E2.3","Top-five RAID · contingency and escalation")
    raid=[["ID / risk","Mitigation · verbatim"]]
    rsrc={}
    for rr in range(36,41):
        raid.append([f"{plan.cell(rr,1).value}: {plan.cell(rr,3).value}",plan.cell(rr,6).value])
        rsrc[(len(raid)-1,0)]=( "PLAN",f"A{rr}")
        rsrc[(len(raid)-1,1)]=( "PLAN",f"F{rr}")
    yr=draw_table(builder,3,"E2.3",rightx,yr,raid,[100,288],size=7.5,sources=rsrc,padding=.4)
    governance_note=(f"Contingency: ${commercial['B75'].value:.2f}m committed; ${commercial['B76'].value:.2f}m unallocated. "
                     f"Escalation ladder: {plan['A44'].value} {plan['B44'].value}; {plan['A45'].value} {plan['B45'].value}; "
                     f"{plan['A46'].value} {plan['B46'].value}; SteerCo {plan['B47'].value}. "
                     f"Recovery: schedule—two reports phased; quality—re-compete and vendor-funded parallel run, two payments held; "
                     f"{vendors['B30'].value:.1f}% exit gate met in {vendors['J20'].value:g} days; ${vendors['J22'].value:g} added.")
    draw_table(builder,3,"E2.3",rightx,yr-1,[[governance_note]],[colw],size=7.5,header=False,padding=.5,leading=8.3)
    for cell in ("B75","B76"):
        number_sources(builder,3,"E2.3",f"{commercial[cell].value:.2f}",f"COMMERCIAL!{cell}",commercial[cell].value)
    for row in (44,45,46):number_sources(builder,3,"E2.3",plan[f"B{row}"].value,f"PLAN!B{row}",plan[f"B{row}"].value)
    for shown,source,value in ((f"{vendors['B30'].value:.1f}%","VENDORS!B30",vendors["B30"].value),
                               (f"{vendors['J20'].value:g}","VENDORS!J20",vendors["J20"].value),
                               (f"{vendors['J22'].value:g}","VENDORS!J22",vendors["J22"].value)):
        number_sources(builder,3,"E2.3",shown,source,value)
    builder.forced_box(3,8,min_y=26,top=58)
    builder.footer(3); c.showPage()

    # Page 3 – workforce, transition, vendor performance and commercial.
    top=builder.begin_page("3 | Vendor workforce","§7 capacity and commercial")
    y=builder.exhibit(MARGIN,top,"E3.1","Demand vs supply by role, region and quarter · vendor ramp, bench and utilisation")
    chart=make_capacity_chart(model,builder)
    builder.draw_chart(chart,MARGIN,y-192,PAGE_W-2*MARGIN,182,font_size=9.0)
    y-=192
    bridge=model.values["CAPACITY"]
    balance_text=(f"Balanced at {bridge['C11'].value:g} FTE at {bridge['A11'].value} forecast; "
                  f"data engineering {signed_number(bridge['D21'].value)} and QA +{bridge['D23'].value:g} before mitigation (bridge).")
    y=builder.draw_para(balance_text,MARGIN,y,PAGE_W-50,size=9,min_y=MIN_CONTENT_Y)
    for shown,source,value in ((f"{bridge['C11'].value:g}","CAPACITY!C11",bridge["C11"].value),
                               (bridge["A11"].value,"CAPACITY!A11",bridge["A11"].value),
                               (signed_number(bridge["D21"].value),"CAPACITY!D21",bridge["D21"].value),
                               (f"+{bridge['D23'].value:g}","CAPACITY!D23",bridge["D23"].value)):
        number_sources(builder,4,"E3.1",shown,source,value)
    y-=10
    leftw=389; gap=8; rightw=PAGE_W-2*MARGIN-leftw-gap; rightx=MARGIN+leftw+gap
    yl=builder.exhibit(MARGIN,y,"E3.2","A-to-D transition and continuity")
    date_value=lambda cell:model.value("ASSUMPTIONS",cell).strftime("%d %b %Y")
    trows=[["Transition item / risk","Date, value / mitigation"]]
    pairs=[("Knowledge transfer",f"{vendors['A41'].value}–{vendors['A46'].value}"),
           ("Dual run",f"{date_value('C46')}–{date_value('C47')}"),
           ("Notice",f"{date_value('C48')}–{date_value('C49')}"),
           ("Working days",f"{vendors['J29'].value:g} + {vendors['J30'].value:g} = {vendors['J33'].value:g}"),
           ("Contingency cost",f"${model.value('ASSUMPTIONS','C55'):,.0f}"),
           ("D3 remediation parallel run",f"Separate; ${vendors['J22'].value:g} added")]
    trows += [list(p) for p in pairs]
    t_sources={}
    for riskrow in range(49,52):
        trows.append([f"{vendors.cell(riskrow,1).value} · {vendors.cell(riskrow,2).value}",vendors.cell(riskrow,8).value])
        t_sources[(len(trows)-1,0)]=( "VENDORS",f"A{riskrow}")
        t_sources[(len(trows)-1,1)]=( "VENDORS",f"H{riskrow}")
        number_sources(builder,4,"E3.2",vendors.cell(riskrow,8).value,f"VENDORS!H{riskrow}",vendors.cell(riskrow,8).value)
    yl=draw_table(builder,4,"E3.2",MARGIN,yl,trows,[85,304],size=7.5,sources=t_sources,padding=.3,leading=7.9)
    for token,source in (("1","VENDORS!A41"),("6","VENDORS!A46"),("01 Apr 2026","ASSUMPTIONS!C46"),("26 May 2026","ASSUMPTIONS!C47"),("11 May 2026","ASSUMPTIONS!C48"),("10 Jun 2026","ASSUMPTIONS!C49"),("40","VENDORS!J29"),("11","VENDORS!J30"),("51","VENDORS!J33"),("183,600","ASSUMPTIONS!C55"),("0","VENDORS!J22"),("1","VENDORS!A49"),("2","VENDORS!A50"),("1","VENDORS!A51")):
        sn,cell=source.split("!");number_sources(builder,4,"E3.2",token,source,model.value(sn,cell))
    builder.record(4,"E3.2","3","DECISIONS!A7",decisions.cell(7,1).value)
    yr=builder.exhibit(rightx,y,"E3.3","Scorecard and Vendor B checkpoints")
    sc=[["Vendor","SLA","Quality","Delivery","Cost","Overall"]]
    scsrc={}
    for rr in range(5,10):
        sc.append([vendors.cell(rr,1).value, f"{vendors.cell(rr,3).value}% {vendors.cell(rr,4).value}",f"{vendors.cell(rr,5).value}% {vendors.cell(rr,6).value}",f"{vendors.cell(rr,7).value}% {vendors.cell(rr,8).value}",f"{vendors.cell(rr,9).value}% {vendors.cell(rr,10).value}",vendors.cell(rr,11).value])
        for ci,col in enumerate((1,3,5,7,9,11)):scsrc[(len(sc)-1,ci)]=( "VENDORS",f"{chr(64+col)}{rr}")
    yr=draw_table(builder,4,"E3.3",rightx,yr,sc,[55,66,66,74,64,57],size=7.5,sources=scsrc,padding=1.0)
    yr-=10
    yr=builder.exhibit(rightx,yr,"","Vendor B remediation checkpoints · defect escape")
    cp=[["Checkpoint","Date","Max","Observed","Day","Result"]]
    cpsrc={}
    for rr in range(34,38):
        cp.append([vendors.cell(rr,1).value,cell_text(model,builder,4,"E3.3","VENDORS",f"B{rr}"),f"{vendors.cell(rr,3).value:.1f}%",f"{vendors.cell(rr,4).value:.1f}%",str(vendors.cell(rr,5).value),vendors.cell(rr,6).value])
        for ci,col in enumerate((1,2,3,4,5,6)):cpsrc[(len(cp)-1,ci)]=( "VENDORS",f"{chr(64+col)}{rr}")
    yr=draw_table(builder,4,"E3.3",rightx,yr,cp,[65,72,46,58,37,42],size=7.5,sources=cpsrc,padding=.7)
    yr-=3
    right_bottom=draw_table(builder,4,"E3.3",rightx,yr,[ ["Amber rule: defect escape above 1.5% for two consecutive months."] ],[rightw],size=7.5,header=False,padding=.3,sources={(0,0):("VENDORS","J23")})
    e34_top=min(yl,right_bottom)-11
    draw_spend_exhibit(model,builder,4,e34_top)
    builder.forced_box(4,7,min_y=23,top=55)
    builder.footer(4); c.showPage()

    # Page 4 – benefits, funnel, health and decisions.
    top=builder.begin_page("4 | Is it working","§8 outcomes and gates")
    envelope_value=model.assumption("envelope")
    envelope_cell=model.assumption_cell("envelope")
    y=builder.exhibit(MARGIN,top,"E4.1",f"Benefits register · run-budget effects outside the ${envelope_value:.2f}m envelope")
    number_sources(builder,5,"E4.1",f"{envelope_value:.2f}",f"ASSUMPTIONS!{envelope_cell}",envelope_value)
    br=[["ID","Category","Line","Measure","Baseline / date","Target / date","Value / date","Status","Owner","Source"]]
    bs={}
    e4_dates=[]
    for rr in range(5,11):
        measure=str(benefits.cell(rr,5).value)
        base_val=benefits.cell(rr,6).value; target_val=benefits.cell(rr,8).value; current=benefits.cell(rr,10).value
        if rr==5: base_text=f"{base_val:g}"; target_text=f"{target_val:g}"
        elif rr==6: base_text=f"${base_val:.0f}m"; target_text=f"${target_val:.1f}m/year"
        elif rr in (7,9): base_text=f"{base_val:.0%}" if rr==7 else f"{base_val:.1f}%"; target_text=f"{target_val:.0%}" if rr==7 else f"≤{target_val:.1f}%"
        elif rr==8: base_text=f"{base_val:g} hours/week"; target_text=f"{target_val:g} hours/week"
        else: base_text="n/a"; target_text="100%"
        if rr==5: value_text=str(int(current))
        elif rr==6: value_text=f"${current:.1f}m/year"
        elif rr in (7,9): value_text=f"{current:.0%}" if rr==7 else f"{current:.1f}%"
        elif rr==8: value_text=str(int(current))+" hours/week"
        else:value_text=f"{current:.0%}"
        base_text += " ("+(benefits.cell(rr,7).value.strftime("%d %b %Y") if isinstance(benefits.cell(rr,7).value,(datetime,date)) else "n/a")+")"
        target_date=benefits.cell(rr,9).value
        target_date_text=target_date.strftime("%d %b %Y") if isinstance(target_date,(datetime,date)) else "n/a"
        if isinstance(target_date,(datetime,date)) and target_date.date()>date(2026,10,31):target_date_text += " forecast"
        target_text += " ("+target_date_text+")"
        value_text += " ("+benefits.cell(rr,11).value.strftime("%d %b %Y")+")"
        br.append([benefits.cell(rr,1).value,benefits.cell(rr,2).value,benefits.cell(rr,3).value,measure,base_text,target_text,value_text,benefits.cell(rr,12).value,benefits.cell(rr,13).value,benefits.cell(rr,14).value])
        for ci,col in enumerate((1,2,3,5,12,13,14)):
            outci=(ci if ci<4 else ci+3)
            bs[(len(br)-1,outci)]=( "BENEFITS",f"{chr(64+col)}{rr}")
        e4_dates.append((len(br)-1,rr,base_text.split(" (",1)[0],target_text.split(" (",1)[0],value_text.split(" (",1)[0]))
    y=draw_table(builder,5,"E4.1",MARGIN,y,br,[20,65,78,112,77,80,77,97,75,86],size=7.5,sources=bs,padding=.8)
    for rowidx,rr,btxt,ttxt,vtxt in e4_dates:
        for shown,cell,value in ((btxt,f"F{rr}",benefits.cell(rr,6).value),(ttxt,f"H{rr}",benefits.cell(rr,8).value),(vtxt,f"J{rr}",benefits.cell(rr,10).value)):
            number_sources(builder,5,"E4.1",shown,f"BENEFITS!{cell}",value)
        for shown,cell in ((br[rowidx][4].split("(",1)[-1].rstrip(")"),f"G{rr}"),(br[rowidx][5].split("(",1)[-1].rstrip(")"),f"I{rr}"),(br[rowidx][6].split("(",1)[-1].rstrip(")"),f"K{rr}")):
            if shown and shown!="n/a":number_sources(builder,5,"E4.1",shown,f"BENEFITS!{cell}",benefits[cell].value)
    y-=10
    leftw=388; rightx=PAGE_W-MARGIN-leftw
    as_at_text=model.value("ASSUMPTIONS","C6").strftime("%d %b %Y")
    yl=builder.exhibit(MARGIN,y,"E4.2",f"Self-service funnel by business unit · as at {as_at_text}")
    number_sources(builder,5,"E4.2",as_at_text,"ASSUMPTIONS!C6",model.value("ASSUMPTIONS","C6"))
    funnel_cols=(1,2,3,4,5,6,7,9,8)
    fr=[["Business unit","Invited","Trained","Active","Habitual","Trained / invited","Active / trained","Habitual / active","Habitual / invited"]]
    frsrc={}
    for rr in range(15,18):
        values=([str(benefits.cell(rr,1).value)]+[str(benefits.cell(rr,col).value) for col in (2,3,4,5)]+
                [f"{benefits.cell(rr,col).value:.1%}" for col in (6,7,9,8)])
        fr.append(values)
        for ci,col in enumerate(funnel_cols):frsrc[(len(fr)-1,ci)]=( "BENEFITS",f"{chr(64+col)}{rr}")
    totalrow=( ["Total"]+[str(benefits.cell(18,col).value) for col in (2,3,4,5)]+
               [f"{benefits.cell(18,col).value:.1%}" for col in (6,7,9,8)])
    fr.append(totalrow)
    for ci,col in enumerate(funnel_cols[1:],1):frsrc[(len(fr)-1,ci)]=( "BENEFITS",f"{chr(64+col)}18")
    funnel_widths=[60,40,40,40,40,42,42,42,42]
    corporate_active_trained=benefits["G17"].value
    highlight_color=colors.HexColor("#F5D5A6")
    yl=draw_table(builder,5,"E4.2",MARGIN,yl,fr,funnel_widths,size=7.5,sources=frsrc,padding=.8,
                  highlights={(3,6):highlight_color})
    builder.record(5,"E4.2","31 Oct 2026","ASSUMPTIONS!C6",model.value("ASSUMPTIONS","C6"))
    client_rate=f"{benefits['G15'].value:.1%}";investment_rate=f"{benefits['G16'].value:.1%}"
    corporate_rate=f"{corporate_active_trained:.1%}"
    drop_text=f"Corporate Services trained to active: {corporate_rate}; compared with {client_rate} (Client Operations) and {investment_rate} (Investment Operations)."
    builder.draw_para(drop_text,MARGIN,yl-3,leftw,size=7.5,leading=8.5,min_y=MIN_CONTENT_Y)
    for token,cell in ((corporate_rate,"G17"),(client_rate,"G15"),(investment_rate,"G16")):
        number_sources(builder,5,"E4.2",token,f"BENEFITS!{cell}",benefits[cell].value)
    note,note_source=model.page_text("e42_note")
    c.setFont(FONT_BOLD,7.5); c.setFillColor(BLUE); c.drawString(MARGIN,yl-35,"Recommendation – not a logged decision")
    builder.smallest_font=min(builder.smallest_font,7.5);builder.smallest_text_font=min(builder.smallest_text_font,7.5)
    builder.add_block("heading",MARGIN,yl-36,stringWidth("Recommendation – not a logged decision",FONT_BOLD,7.5),9)
    note_bottom=builder.draw_para(note,MARGIN,yl-42,leftw,size=7.5,leading=9,min_y=MIN_CONTENT_Y)
    number_sources(builder,5,"E4.2",note,note_source,note)
    yr=builder.exhibit(rightx,y,"E4.3","Delivery health · value, thresholds and escalation")
    hr=[["Measure","Value","Green","Amber","Red","Escalation if amber","Escalation if red"]]
    hsrc={}
    for ri,row in enumerate(range(26,32),1):
        item=benefits.cell(row,1).value
        val=benefits.cell(row,3).value
        if item in ("H2","H4","H6"):value_text=f"{val:.1%} · {benefits.cell(row,7).value}"
        elif item=="H3":value_text=f"{val:.1f}% · {benefits.cell(row,7).value}"
        elif item=="H1":value_text=f"{val:g} days · {benefits.cell(row,7).value}"
        else:value_text=f"{val:g} FTE · {benefits.cell(row,7).value}"
        if item in ("H2","H4","H6"):
            green=f"≥ {benefits.cell(row,4).value:.0%}";amber=f"≥ {benefits.cell(row,5).value:.0%}";red=f"< {benefits.cell(row,6).value:.0%}"
        elif item=="H3":green=f"≤ {benefits.cell(row,4).value:.1f}%";amber=f"> {benefits.cell(row,5).value:.1f}%";red=f"> {benefits.cell(row,6).value:.1f}%"
        elif item=="H1":green=f"≥ {benefits.cell(row,4).value:g}";amber=f"≥ {benefits.cell(row,5).value:g}";red=f"≤ {benefits.cell(row,6).value:g}"
        else:green=f"≤ {benefits.cell(row,4).value:g}";amber=f"> {benefits.cell(row,4).value:g}";red=f"> {benefits.cell(row,5).value:g}"
        hr.append([benefits.cell(row,2).value,value_text,green,amber,red,
                   benefits.cell(row,8).value,benefits.cell(row,9).value])
        for ci,col in enumerate((2,3,4,5,6,8,9)):hsrc[(ri,ci)]=( "BENEFITS",f"{chr(64+col)}{row}")
        number_sources(builder,5,"E4.3",value_text,f"BENEFITS!C{row}",benefits.cell(row,3).value)
        for ci,col,text in ((2,4,green),(3,5,amber),(4,6,red)):number_sources(builder,5,"E4.3",text,f"BENEFITS!{chr(64+col)}{row}",benefits.cell(row,col).value)
    yr=draw_table(builder,5,"E4.3",rightx,yr,hr,[68,46,37,37,37,80,82],size=7.5,sources=hsrc,padding=.7,leading=8.2)
    referral="E4.4 Decision log D1–D5 – full log on page 7 (appendix)"
    builder.draw_para(referral,rightx,151,rightw,size=8.5,leading=10,min_y=82)
    appendix_page,appendix_page_source=model.pdf_page(7)
    builder.record(5,"E4.4", "7",appendix_page_source,appendix_page)
    builder.forced_box(5,9,min_y=34,top=78)

    builder.footer(5); c.showPage()

    # Page 5 – cadence.
    top=builder.begin_page("Cadence","Accountability and monthly measures")
    y=builder.exhibit(MARGIN,top,"","Three forums · roles, authority and information")
    cr=[["Forum","Frequency","Chair","Attends","Decides here and nowhere else","Pack received"]]
    csrc={}
    for rr in range(58,61):
        cr.append([plan.cell(rr,col).value for col in range(1,7)])
        for ci,col in enumerate(range(1,7)):csrc[(len(cr)-1,ci)]=( "PLAN",f"{chr(64+col)}{rr}")
    y=draw_table(builder,6,"Cadence",MARGIN,y,cr,[100,62,100,180,225,100],size=8.0,sources=csrc,padding=2.0)
    y-=28
    y=builder.exhibit(MARGIN,y,"","Managing director monthly measures · formula references")
    measures=[["Measure","Monthly readout"]]
    measure_text=[f"{benefits['A26'].value} float to {benefits['I10'].value.strftime('%d %b %Y')} (forecast)",
                  f"{benefits['A31'].value} self-service share vs {benefits['H7'].value:.0%} gate",
                  f"Contingency unallocated of ${plan['C66'].value:.2f}m total"]
    measure_values=[f"{plan.cell(64,2).value:.0f} working days",f"{plan.cell(65,2).value:.1%} / {plan.cell(65,3).value:.0%}",f"${plan.cell(66,2).value:.2f}m / ${plan.cell(66,3).value:.2f}m"]
    number_sources(builder,6,"Cadence",measure_text[0].split(" to ",1)[1].split(" (",1)[0],"BENEFITS!I10",benefits["I10"].value)
    number_sources(builder,6,"Cadence",f"{benefits['H7'].value:.0%}","BENEFITS!H7",benefits["H7"].value)
    number_sources(builder,6,"Cadence",f"{plan['C66'].value:.2f}","PLAN!C66",plan["C66"].value)
    for i,rr in enumerate(range(64,67)):
        measures.append([measure_text[i],measure_values[i]])
        if rr==64:number_sources(builder,6,"Cadence",measure_values[i],f"PLAN!B{rr}",plan.cell(rr,2).value)
        elif rr==65:
            builder.record(6,"Cadence",f"{plan.cell(rr,2).value:.1%}",f"PLAN!B{rr}",plan.cell(rr,2).value)
            builder.record(6,"Cadence",f"{plan.cell(rr,3).value:.0%}",f"PLAN!C{rr}",plan.cell(rr,3).value)
        else:
            builder.record(6,"Cadence",f"${plan.cell(rr,2).value:.2f}",f"PLAN!B{rr}",plan.cell(rr,2).value)
            builder.record(6,"Cadence",f"${plan.cell(rr,3).value:.2f}",f"PLAN!C{rr}",plan.cell(rr,3).value)
    y=draw_table(builder,6,"Cadence",MARGIN,y,measures,[300,467],size=8.5,padding=2.3)
    builder.record(6,"Cadence",benefits["I10"].value.strftime("%d %b %Y"),"BENEFITS!I10",benefits.cell(10,9).value)
    rationale,rationale_source=model.page_text("cadence_rationale")
    builder.draw_para(rationale,MARGIN,y-24,PAGE_W-2*MARGIN,size=9,leading=12,min_y=MIN_CONTENT_Y)
    number_sources(builder,6,"Cadence",rationale,rationale_source,rationale)
    builder.footer(6); c.showPage()

    # Page 7 – full decision log appendix.
    top=builder.begin_page("Appendix | Decision log","E4.4 · all fields verbatim")
    log_top=builder.exhibit(MARGIN,top,"E4.4","Decision log D1–D5 · all fields verbatim")
    dr=[["D-number","Date","Decision","Options","Evidence","Decider","Outcome","Status","Page"]]
    ds={}
    for rr in range(5,10):
        cells=[decisions.cell(rr,col).value for col in range(1,10)]
        date_text=cell_text(model,builder,7,"E4.4","DECISIONS",f"B{rr}")
        line=[cells[0],date_text,*cells[2:9]]
        dr.append(line)
        for col in range(1,10):ds[(len(dr)-1,col-1)]=( "DECISIONS",f"{chr(64+col)}{rr}")
    log_widths=[60,60,115,100,145,70,125,80,PAGE_W-2*MARGIN-755]
    draw_table(builder,7,"E4.4",MARGIN,log_top,dr,log_widths,size=7.5,sources=ds,padding=.7,min_y=82,leading=8.3)
    builder.footer(7); c.showPage()

    builder.validate_layout()
    builder.canvas.save()
    builder.save_trace()
    if not shutil.which("pdftoppm"):
        raise RuntimeError("pdftoppm is required to regenerate the 100 dpi page previews")
    preview_root=OUT_PREVIEW.resolve()
    if not preview_root.is_relative_to(ROOT.resolve()):
        raise RuntimeError(f"Preview output path is outside the repository: {preview_root}")
    if OUT_PREVIEW.exists():shutil.rmtree(OUT_PREVIEW)
    OUT_PREVIEW.mkdir(parents=True,exist_ok=True)
    prefix=OUT_PREVIEW/"page"
    render=subprocess.run([shutil.which("pdftoppm"),"-png","-r","100",str(OUT_PDF),str(prefix)],capture_output=True,text=True,timeout=120)
    if render.returncode:raise RuntimeError(f"PDF preview render failed: {render.stderr}")
    builder.smallest_font=min(builder.smallest_text_font,builder.smallest_chart_font)
    audit={"text_fit_blocks":builder.text_fit_checks,"word_fit_checks":builder.word_fit_checks,"layout_blocks":builder.layout_blocks,
           "smallest_text_font_pt":builder.smallest_text_font,"smallest_chart_font_at_placed_size_pt":builder.smallest_chart_font,
           "smallest_pdf_font_pt":builder.smallest_font,"overlap_check":"passed",
           "chart_text_overlap_checks":builder.chart_text_overlap_checks,"chart_text_overlap_count":builder.chart_text_overlap_count,
           "chart_text_min_clearance_pt":builder.chart_text_min_clearance_pt,"chart_images_audited":builder.chart_images_audited,
           "capacity_chart_labels":builder.chart_labels}
    (OUT_PREVIEW/"layout_audit.json").write_text(json.dumps(audit,indent=2),encoding="utf-8")
    return audit


def main():
    recalced=run_recalc_pipeline()
    model=Model(recalced)
    workbook= model.values
    if workbook["ASSUMPTIONS"]["K113"].value is not True:
        raise RuntimeError(f"ALL CHECKS not TRUE after LibreOffice recalc: {workbook['ASSUMPTIONS']['K113'].value}")
    return make_pdf(model)


if __name__ == "__main__":
    result=main()
    print(f"Smallest text-layer font: {result['smallest_text_font_pt']:.2f} pt")
    print(f"Smallest chart font at placed size: {result['smallest_chart_font_at_placed_size_pt']:.2f} pt")
    print(f"Smallest PDF font: {result['smallest_pdf_font_pt']:.2f} pt")
