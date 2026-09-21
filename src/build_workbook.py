"""
Build output/Retail_Peer_Benchmarking.xlsx from the extracted SEC data.

Every number on every analytical tab is a formula. The only hard-coded values
are on Raw Data (blue font, one SEC citation each) and the Scorecard year
input. Design rules:

  * Raw Data blocks carry named ranges (Walmart_Revenue, Kroger_TotalDebt...)
    so downstream formulas read like sentences.
  * Ratio Calculations pulls every input with XLOOKUP against the Years header,
    wrapped in IFERROR so a missing year yields "n/a" instead of an error.
  * Total Debt is itself a formula over the extracted components, applying the
    ASC 842 policy in config/debt_policy.yaml.
  * Scorecard re-ranks for whatever year is typed into the ScorecardYear cell.
  * Trend charts are native Excel combo charts: a shaded column series marks
    FY2020-FY2023 behind the ratio lines; NA() in the helper block makes Excel
    draw gaps where data does not exist.
  * Executive Summary bullets are string formulas, so the numbers quoted in the
    narrative update with the model.

Run:  .venv\\Scripts\\python -m src.build_workbook
Then: .venv\\Scripts\\python -m src.recalc_excel output/Retail_Peer_Benchmarking.xlsx
"""

from __future__ import annotations

import math
import re
from datetime import date

import pandas as pd
import yaml
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines, NumericAxis, TextAxis
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from src import edgar

OUT = edgar.ROOT / "output" / "Retail_Peer_Benchmarking.xlsx"
TODAY = date(2026, 9, 18)

# ------------------------------------------------------------------ palette
NAVY, TEAL, LIGHT, PALE, GREY, MID = "1F3864", "2E75B6", "D9E2F3", "EEF3FA", "F2F2F2", "808080"
INPUT_BLUE = "0000FF"
BAND = "E7E6E6"
SERIES_COLOR = {"WMT": "1F77B4", "COST": "D62728", "KR": "2CA02C", "SAMS": "9467BD", "WFM": "FF7F0E"}

F_TITLE = Font(name="Calibri", size=20, bold=True, color=NAVY)
F_SUB = Font(name="Calibri", size=11, italic=True, color=MID)
F_H = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
F_B = Font(name="Calibri", size=11, bold=True, color=NAVY)
F_N = Font(name="Calibri", size=10)
F_IN = Font(name="Calibri", size=10, color=INPUT_BLUE)
F_NA = Font(name="Calibri", size=10, italic=True, color=MID)
F_NOTE = Font(name="Calibri", size=9, color=MID)
FILL_H = PatternFill("solid", fgColor=NAVY)
FILL_T = PatternFill("solid", fgColor=TEAL)
FILL_L = PatternFill("solid", fgColor=LIGHT)
FILL_P = PatternFill("solid", fgColor=PALE)
FILL_G = PatternFill("solid", fgColor=GREY)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOTTOM = Border(bottom=Side(style="medium", color=NAVY))
CENTER = Alignment(horizontal="center", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
WRAP = Alignment(wrap_text=True, vertical="top")

FMT_M = '#,##0;(#,##0);"–"'
FMT_PCT = '0.0%;-0.0%;0.0%'
FMT_X = '0.00"x"'
FMT_PP = '+0.0%;-0.0%;0.0%'
FMT_DX = '+0.00"x";-0.00"x";0.00"x"'
FMT_FY = '"FY"0'

YEARS = list(range(2015, 2026))         # Raw Data columns (FY2015 = growth base year)
RYEARS = list(range(2016, 2026))        # ratio / chart years
COMPANIES = [  # ticker, display, range prefix
    ("WMT", "Walmart", "Walmart"),
    ("COST", "Costco", "Costco"),
    ("KR", "Kroger", "Kroger"),
    ("SAMS", "Sam's Club", "SamsClub"),
    ("WFM", "Whole Foods", "WholeFoods"),
]
DISPLAY = {t: d for t, d, _ in COMPANIES}
PREFIX = {t: p for t, _, p in COMPANIES}

# Raw Data row specification: (label, metric key, range suffix, kind)
INPUT_ROWS = [
    ("Revenue (net sales)", "revenue", "Revenue"),
    ("Cost of sales", "cost_of_sales", "COGS"),
    ("Net income", "net_income", "NetIncome"),
    ("Current assets", "current_assets", "CurrentAssets"),
    ("Current liabilities", "current_liabilities", "CurrentLiabilities"),
    ("Total equity (incl. noncontrolling interest)", "total_equity", "TotalEquity"),
]
DEBT_ROWS = [
    ("Short-term borrowings", "short_term_borrowings"),
    ("Long-term debt – due within one year", "ltd_current"),
    ("Long-term debt – non-current", "ltd_noncurrent"),
    ("Finance / capital lease liability – current", "finance_lease_current"),
    ("Finance / capital lease liability – non-current", "finance_lease_noncurrent"),
    ("Finance lease liability – total (years not split on the balance sheet)", "finance_lease_total"),
    ("Long-term debt incl. finance leases – current (combined line)", "ltd_and_leases_current"),
    ("Long-term debt incl. finance leases – non-current (combined line)", "ltd_and_leases_noncurrent"),
]
MEMO_ROWS = [
    ("Operating income", "operating_income", "OperatingIncome"),
    ("Operating lease liability (ASC 842; excluded from Total debt)", "operating_lease", "OperatingLease"),
    ("Total assets", "total_assets", "TotalAssets"),
    ("Inventory", "inventory", "Inventory"),
    ("LIFO reserve (where disclosed)", "lifo_reserve", "LIFOReserve"),
]

n_formulas = 0


def put(ws, ref, value, font=F_N, fmt=None, fill=None, align=None, border=None):
    global n_formulas
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if border:
        c.border = border
    if isinstance(value, str) and value.startswith("="):
        n_formulas += 1
    return c


def header_band(ws, row, text, first_col=1, last_col=14, fill=FILL_H, font=F_H, height=20):
    ws.merge_cells(start_row=row, start_column=first_col, end_row=row, end_column=last_col)
    put(ws, f"{L(first_col)}{row}", text, font=font, fill=fill, align=Alignment(vertical="center", indent=1))
    for col in range(first_col, last_col + 1):
        ws.cell(row=row, column=col).fill = fill
    ws.row_dimensions[row].height = height


EYEBROW = Font(name="Aptos", size=8, bold=True, color="9DB4D8")
F_BAND_TITLE = Font(name="Aptos Display", size=22, bold=True, color="FFFFFF")
F_BAND_RIGHT = Font(name="Aptos", size=8, color="9DB4D8")
F_BAND_SUB = Font(name="Aptos", size=10, italic=True, color="D9E2F3")
RULE = Border(bottom=Side(style="thick", color=TEAL))
REPORT_TAG = "RETAIL PEER BENCHMARKING   ·   FY2016–FY2025   ·   SEC 10-K DATA"


def title_band(ws, number, title, subtitle, last_col, text_col=1, subtitle_in_band=False):
    """House-style title band (rows 1-3): navy bar with eyebrow + white title, teal rule,
    report tag on the right; subtitle in gray italics on row 3 (or inside the band)."""
    for r in (1, 2):
        for c in range(1, last_col + 1):
            ws.cell(row=r, column=c).fill = FILL_H
    for c in range(1, last_col + 1):
        ws.cell(row=2, column=c).border = RULE
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 32
    tc = L(text_col)
    put(ws, f"{tc}1", f"{number:02d}   —   {title.upper()}", font=EYEBROW, fill=FILL_H, align=Alignment(vertical="bottom", indent=1))
    put(ws, f"{tc}2", title, font=F_BAND_TITLE, fill=FILL_H, align=Alignment(vertical="center", indent=1), border=RULE)
    rc = L(last_col)
    if rc == tc:   # two-column tabs: eyebrow and report tag share one cell
        put(ws, f"{tc}1", f"{number:02d}   —   {title.upper()}          ·          {REPORT_TAG}", font=EYEBROW, fill=FILL_H, align=Alignment(vertical="bottom", indent=1))
    else:
        put(ws, f"{rc}1", REPORT_TAG, font=F_BAND_RIGHT, fill=FILL_H, align=Alignment(horizontal="right", vertical="bottom", indent=1))
    if subtitle_in_band:
        put(ws, f"{rc}2", subtitle, font=F_BAND_SUB, fill=FILL_H, align=Alignment(horizontal="right", vertical="center", indent=1), border=RULE)
    else:
        ws.row_dimensions[3].height = 18
        put(ws, f"{tc}3", subtitle, font=F_SUB, align=Alignment(vertical="center", indent=1))


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def sheet_setup(ws, tab_color, gridlines=True, landscape=True):
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.showGridLines = gridlines
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = "Retail Peer Financial Benchmarking – Ryan Barry – &D"
    ws.oddFooter.right.text = "Page &P of &N"


# ------------------------------------------------------------------ data
def load() -> dict:
    facts = pd.read_csv(edgar.ROOT / "data" / "facts_long.csv")
    segs = pd.read_csv(edgar.ROOT / "data" / "segments_long.csv")
    drivers = pd.read_csv(edgar.ROOT / "data" / "driver_facts.csv")
    cfg = edgar.load_config()
    policy = yaml.safe_load(open(edgar.ROOT / "config" / "debt_policy.yaml", encoding="utf-8"))
    data: dict[str, dict] = {}
    for t, _, _ in COMPANIES:
        src = segs if t == "SAMS" else facts
        f = src[src.ticker == t]
        val = f.pivot(index="fiscal_year", columns="metric", values="value")
        tag = f.pivot(index="fiscal_year", columns="metric", values="tag")
        status = f.pivot(index="fiscal_year", columns="metric", values="status")
        note = f.pivot(index="fiscal_year", columns="metric", values="note") if "note" in f else None
        ends = f.groupby("fiscal_year").period_end.max()
        # operating lease liability: total tag if present else current + non-current
        if "operating_lease_total" in val:
            ol = val["operating_lease_total"].fillna(val[["operating_lease_current", "operating_lease_noncurrent"]].sum(axis=1, min_count=1))
            # a literal 0 here is a pre-ASC 842 comparative (no such liability existed), not a balance
            val["operating_lease"] = ol.where(ol != 0)
            tag["operating_lease"] = tag["operating_lease_total"].astype(object).where(tag["operating_lease_total"].notna(), tag.get("operating_lease_noncurrent"))
        if t == "COST":
            # Net-sales basis (excludes membership fees) so Costco matches Walmart / Sam's Club.
            ns = segs[(segs.ticker == "COST") & (segs.metric == "net_sales")].set_index("fiscal_year")
            val["total_revenue"], tag["total_revenue"], status["total_revenue"] = val["revenue"], tag["revenue"], status["revenue"]
            val["revenue"] = ns.value.reindex(val.index)
            tag["revenue"] = ns.tag.reindex(val.index)
            status["revenue"] = ns.status.reindex(val.index)
        data[t] = {"val": val, "tag": tag, "status": status, "note": note, "ends": ends}
    amz = segs[(segs.ticker == "WFM") & (segs.metric == "revenue") & (segs.fiscal_year >= 2018)].set_index("fiscal_year")
    data["WFM"]["successor"] = amz
    return {"data": data, "drivers": drivers, "cfg": cfg, "policy": policy, "facts": facts, "segs": segs}


def has(d, key, y):
    v = d["val"].get(key)
    return v is not None and y in v.index and pd.notna(v.loc[y])


# ------------------------------------------------------------------ Cover
def build_cover(wb):
    ws = wb.active
    ws.title = "Cover"
    sheet_setup(ws, NAVY, gridlines=False, landscape=False)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 70
    ws.column_dimensions["E"].width = 4
    for r in range(1, 5):
        for c in range(1, 6):
            ws.cell(row=r, column=c).fill = FILL_H
    # monogram: white letters in a teal-ruled square; sits on the navy band
    ws.merge_cells("B2:B3")
    put(ws, "B2", "RB", font=Font(name="Aptos Display", size=16, bold=True, color="FFFFFF"), fill=FILL_H, align=CENTER,
        border=Border(left=Side(style="medium", color=TEAL), right=Side(style="medium", color=TEAL),
                      top=Side(style="medium", color=TEAL), bottom=Side(style="medium", color=TEAL)))
    ws["B3"].border = Border(left=Side(style="medium", color=TEAL), right=Side(style="medium", color=TEAL), bottom=Side(style="medium", color=TEAL))
    put(ws, "C2", "RYAN BARRY", font=Font(name="Aptos", size=9, bold=True, color="FFFFFF"), align=Alignment(vertical="center", indent=1))
    put(ws, "C3", "Accounting & financial analysis · SEC 10-K filings", font=Font(name="Aptos", size=8, color="9DB4D8"), align=Alignment(vertical="top", indent=1))
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 22
    put(ws, "D2", "01   —   COVER          ·          " + REPORT_TAG, font=Font(name="Aptos", size=8, bold=True, color="9DB4D8"), align=Alignment(horizontal="right", vertical="center"))
    put(ws, "C7", "Retail Peer Financial Benchmarking", font=Font(name="Aptos Display", size=28, bold=True, color=NAVY))
    put(ws, "C8", "Walmart, Costco, Kroger & Sam's Club", font=Font(name="Aptos Display", size=22, bold=True, color=TEAL))
    put(ws, "C9", "with Whole Foods Market on a best-effort basis", font=Font(name="Aptos", size=13, italic=True, color=MID))
    ws.merge_cells("C11:D11")
    put(ws, "C11", "Ten-year financial comparison (FY2016–FY2025) built from SEC 10-K filings: capital structure, "
                   "resilience through 2020–2023, and margin pressure — every ratio a live formula traceable to its XBRL fact.",
        font=Font(name="Calibri", size=11, color="404040"), align=WRAP)
    ws.row_dimensions[11].height = 48
    rows = [
        ("Prepared by", "Ryan Barry"),
        ("Date", TODAY.strftime("%B %d, %Y")),
        ("Data source", "SEC EDGAR – 10-K XBRL filings (companyfacts API and instance documents)"),
        ("Coverage", "Walmart Inc. · Costco Wholesale Corp. · The Kroger Co. – full 10-K coverage FY2016–FY2025"),
        ("", "Sam's Club – Walmart reportable segment (ASC 280): sales, operating income, assets; cost of sales FY2022+"),
        ("", "Whole Foods Market – standalone 10-Ks FY2016–FY2017; Amazon 'Physical stores' net sales thereafter"),
        ("Units", "USD millions unless stated; fiscal years labeled by SEC frame convention (see Methodology)"),
        ("Status", "All formulas live; recalculated in Excel with zero formula errors"),
    ]
    r = 13
    for k, v in rows:
        put(ws, f"C{r}", k, font=F_B, align=Alignment(vertical="top"))
        put(ws, f"D{r}", v, font=F_N, align=Alignment(wrap_text=True, vertical="top"))
        ws.row_dimensions[r].height = 15 * max(1, math.ceil(len(v) / 95))
        r += 1
    put(ws, f"C{r + 1}", "Contents", font=F_B, border=BOTTOM)
    put(ws, f"D{r + 1}", "", border=BOTTOM)
    tabs = [
        ("Executive Summary", "Findings against the three key questions, trend shifts, data caveats"),
        ("Raw Data", "As-reported inputs by company and year, named ranges, SEC citations, debt build"),
        ("Ratio Calculations", "Gross margin, operating margin, net margin, current ratio, debt-to-equity, revenue growth"),
        ("Scorecard", "Latest-year ranking per metric with conditional formatting and takeaways"),
        ("Trend", "Native line charts FY2016–FY2025 with the FY2020–FY2023 shock window shaded"),
        ("Methodology & Tools", "Sources, techniques, accounting-policy comparability, assumptions"),
        ("Data Lineage", "Every input value with XBRL tag, accession number, filing date and restatement flag"),
    ]
    r += 2
    for t, desc in tabs:
        c = put(ws, f"C{r}", t, font=Font(name="Calibri", size=11, color=TEAL, underline="single"))
        c.hyperlink = f"#'{t}'!A1"
        put(ws, f"D{r}", desc, font=F_N)
        r += 1
    r += 1
    put(ws, f"C{r}", "How to read this workbook", font=F_B, border=BOTTOM)
    put(ws, f"D{r}", "", border=BOTTOM)
    r += 1
    guide = [
        ("Start here", "Executive Summary – the three questions, each answered in one bold line, followed by the evidence."),
        ("Then", "Scorecard – who is strongest and weakest on each metric this year. Change the year in cell B3 to re-rank any year since FY2016."),
        ("Then", "Trend – six charts, FY2016–FY2025, with the 2020–2023 shock window shaded. This is where the resilience and margin-pressure answers are visible."),
        ("If you want the numbers", "Ratio Calculations (every ratio, every year, all formulas) and Raw Data (the as-reported inputs and the debt build)."),
        ("If you want to verify one", "Data Lineage – find the value, click its EDGAR link, open the 10-K. Every input has one."),
        ("End here", "Methodology & Tools – how it was built, the accounting-policy choices (ASC 606, 842, 280, LIFO) and what they do to comparability."),
    ]
    for k, v in guide:
        put(ws, f"C{r}", k, font=Font(name="Calibri", size=10, bold=True, color=TEAL), align=Alignment(vertical="top"))
        put(ws, f"D{r}", v, font=F_N, align=Alignment(wrap_text=True, vertical="top"))
        ws.row_dimensions[r].height = 15 * max(1, math.ceil(len(v) / 95))
        r += 1
    ws.page_setup.fitToHeight = 1
    put(ws, f"C{r + 1}", "Personal portfolio project built from public filings only. Not an H-E-B work product; no non-public information used.",
        font=F_NOTE)
    ws.merge_cells(f"C{r + 1}:D{r + 1}")
    return ws


# ------------------------------------------------------------------ Raw Data
def build_raw(wb, D) -> dict:
    ws = wb.create_sheet("Raw Data")
    sheet_setup(ws, TEAL)
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 40
    for i in range(len(YEARS)):
        ws.column_dimensions[L(3 + i)].width = 11.5
    ws.column_dimensions[L(3 + len(YEARS))].width = 70
    last_col = 3 + len(YEARS)
    ycol = {y: L(3 + i) for i, y in enumerate(YEARS)}

    title_band(ws, 3, "Raw Data", "As reported in SEC 10-K filings. USD millions. Blue = hard-coded input from EDGAR XBRL (one citation per value on the "
                                   "Data Lineage tab); black = formula; 'n/a' = not reported – flagged, never estimated. Fiscal years labeled by SEC frame "
                                   "convention (FY2025 = Walmart fiscal 2026, Kroger 2025, Costco 2025).", last_col)
    put(ws, "A4", "Line item", font=F_H, fill=FILL_H, align=Alignment(vertical="center", indent=1))
    put(ws, "B4", "XBRL tag(s) / basis", font=F_H, fill=FILL_H, align=CENTER)
    for y in YEARS:
        put(ws, f"{ycol[y]}4", y, font=F_H, fill=FILL_H, fmt=FMT_FY, align=CENTER)
    put(ws, f"{L(last_col)}4", "Source / notes", font=F_H, fill=FILL_H, align=Alignment(vertical="center", indent=1))
    ws.row_dimensions[4].height = 22
    name(wb, "Years", f"'Raw Data'!$C$4:${ycol[2025]}$4")
    ws.freeze_panes = "C5"

    rowmap: dict[str, dict[str, int]] = {}
    r = 6
    for t, disp, pre in COMPANIES:
        d = D["data"][t]
        rowmap[t] = {}
        cfg = D["cfg"]["companies"][t]
        sub = {"SAMS": " — Walmart reportable segment (ASC 280); no balance sheet or net income at segment level",
               "WFM": " — standalone 10-Ks through FY2017; acquired by Amazon.com 2017-08-28"}.get(t, f" — {cfg['name']}, CIK {cfg['cik']}")
        header_band(ws, r, f"{disp}{sub}", 1, last_col)
        r += 1
        put(ws, f"A{r}", "Fiscal year end", font=F_NOTE, align=Alignment(indent=1))
        put(ws, f"B{r}", cfg.get("fye_rule", "per Walmart"), font=F_NOTE)
        for y in YEARS:
            if y in d["ends"].index:
                put(ws, f"{ycol[y]}{r}", pd.to_datetime(d["ends"].loc[y]).date(), font=F_NOTE, fmt="yyyy-mm-dd", align=CENTER)
        r += 1

        def write_row(label, key, suffix=None, font=F_IN, indent=1, fill=None):
            nonlocal r
            put(ws, f"A{r}", label, font=F_N, align=Alignment(indent=indent), fill=fill)
            tags = sorted({str(x) for x in d["tag"].get(key, pd.Series(dtype=object)).dropna()}) if key in d["tag"] else []
            put(ws, f"B{r}", "; ".join(tags) if tags else "—", font=F_NOTE, fill=fill)
            notes = []
            for y in YEARS:
                cell = f"{ycol[y]}{r}"
                if has(d, key, y):
                    put(ws, cell, float(d["val"][key].loc[y]) / 1e6, font=font, fmt=FMT_M, fill=fill)
                    st = d["status"][key].loc[y] if key in d["status"] else "ok"
                    if st in ("instance", "derived"):
                        ws[cell].fill = PatternFill("solid", fgColor="FFF2CC")
                        notes.append(f"FY{y}: {st} – {d['note'][key].loc[y]}" if d["note"] is not None else f"FY{y}: {st}")
                else:
                    put(ws, cell, "n/a", font=F_NA, align=RIGHT, fill=fill)
            if suffix:
                rowmap[t][suffix] = r
                name(wb, f"{pre}_{suffix}", f"'Raw Data'!$C${r}:${ycol[2025]}${r}")
            rowmap[t][key] = r
            return notes

        base_note = {
            "SAMS": "Walmart 10-K segment note (ASC 280), StatementBusinessSegmentsAxis = SamsClubUSMember; parsed from the XBRL instance. ",
            "WFM": "Whole Foods 10-K XBRL (CIK 865436). No filings after FY2017. ",
        }.get(t, "10-K XBRL via SEC companyfacts API; latest-filed (restated) value. ")
        for label, key, suffix in INPUT_ROWS:
            notes = write_row(label, key, suffix)
            extra = ""
            if t == "SAMS" and key in ("net_income", "current_assets", "current_liabilities", "total_equity"):
                extra = "Not reported at segment level under ASC 280."
            if t == "SAMS" and key == "cost_of_sales":
                extra = "Segment cost of revenue first disclosed in the FY2024 10-K under ASU 2023-07 (three comparative years)."
            if t == "COST" and key == "revenue":
                extra = "NET SALES (merchandise), excluding membership fees - ProductOrServiceAxis = ProductMember from FY2018 (parsed from the XBRL instance), SalesRevenueNet before ASC 606. Total revenue is shown as a memo row. "
            if t == "WFM" and key == "cost_of_sales":
                extra = "Whole Foods' line is 'Cost of goods sold and occupancy costs' – includes rent; gross margin not directly comparable. "
            if key == "total_equity" and t == "WMT":
                extra = "Includes noncontrolling interest (Walmart FY2025: $6.3B). "
            if key == "net_income":
                extra += "Attributable to the parent (NetIncomeLoss)."
            put(ws, f"{L(last_col)}{r}", (base_note + extra + " ".join(notes)).strip(), font=F_NOTE, align=Alignment(wrap_text=False))
            r += 1
        # Total debt (formula)
        put(ws, f"A{r}", "Total debt (formula – see debt build below)", font=Font(name="Calibri", size=10, bold=True), align=Alignment(indent=1), fill=FILL_P)
        put(ws, f"B{r}", "Policy: config/debt_policy.yaml (ASC 842)", font=F_NOTE, fill=FILL_P)
        rowmap[t]["TotalDebt"] = r
        name(wb, f"{pre}_TotalDebt", f"'Raw Data'!$C${r}:${ycol[2025]}${r}")
        debt_row = r
        r += 1
        # debt build rows (only those with data for this company)
        present = [(lab, k) for lab, k in DEBT_ROWS if k in d["val"] and d["val"][k].notna().any()]
        for lab, k in present:
            notes = write_row("    " + lab, k, None, font=F_IN, indent=2, fill=FILL_G)
            put(ws, f"{L(last_col)}{r}", ("Debt component. " + " ".join(notes)).strip(), font=F_NOTE)
            r += 1
        pol = D["policy"]["companies"][t]
        for y in YEARS:
            cell = f"{ycol[y]}{debt_row}"
            if not present or not has(d, "revenue", y) or t == "SAMS":
                put(ws, cell, "n/a", font=F_NA, align=RIGHT, fill=FILL_P)
                continue
            c = ycol[y]
            ref = {k: f"{c}{rowmap[t][k]}" for _, k in present}
            base = [ref[k] for k in ("short_term_borrowings", "ltd_current", "ltd_noncurrent") if k in ref]
            base_expr = f"SUM({','.join(base)})" if base else "0"
            split = "finance_lease_current" in ref and "finance_lease_noncurrent" in ref
            tot = "finance_lease_total" in ref
            if split and tot:
                fl = f"IF(COUNT({ref['finance_lease_current']},{ref['finance_lease_noncurrent']})>0,SUM({ref['finance_lease_current']},{ref['finance_lease_noncurrent']}),IF(ISNUMBER({ref['finance_lease_total']}),{ref['finance_lease_total']},0))"
            elif split:
                fl = f"SUM({ref['finance_lease_current']},{ref['finance_lease_noncurrent']})"
            elif tot:
                fl = f"IF(ISNUMBER({ref['finance_lease_total']}),{ref['finance_lease_total']},0)"
            else:
                fl = "0"
            core = f"{base_expr}+{fl}"
            if "ltd_and_leases_current" in ref:
                formula = f"=IF(ISNUMBER({ref['ltd_and_leases_current']}),SUM({ref['ltd_and_leases_current']},{ref['ltd_and_leases_noncurrent']}),{core})"
            else:
                formula = f"={core}"
            put(ws, cell, formula, font=Font(name="Calibri", size=10, bold=True), fmt=FMT_M, fill=FILL_P)
        put(ws, f"{L(last_col)}{debt_row}", pol.get("note", "") + (" " + pol["rule"] if "rule" in pol else ""), font=F_NOTE)
        # memo rows
        for lab, k, suffix in MEMO_ROWS:
            if k in d["val"] and d["val"][k].notna().any():
                notes = write_row(lab + " (memo)", k, suffix, font=F_IN, indent=1)
                put(ws, f"{L(last_col)}{r}", ("Memo item – not used in the five core ratios except where stated. " + " ".join(notes)).strip(), font=F_NOTE)
                r += 1
        if t == "COST":
            notes = write_row("Total revenue incl. membership fees (memo)", "total_revenue", "TotalRevenue", font=F_IN, indent=1)
            put(ws, f"{L(last_col)}{r}", "Costco's income-statement total ('Total revenue'). Net sales above exclude membership fees so the revenue basis matches Walmart and Sam's Club.", font=F_NOTE)
            r += 1
            put(ws, f"A{r}", "Membership fee revenue = total revenue - net sales (memo, formula)", font=F_N, align=Alignment(indent=1))
            put(ws, f"B{r}", "Formula", font=F_NOTE)
            for y in YEARS:
                put(ws, f"{ycol[y]}{r}", f'=IFERROR({ycol[y]}{rowmap[t]["total_revenue"]}-{ycol[y]}{rowmap[t]["revenue"]},"n/a")', font=Font(name="Calibri", size=10, bold=True), fmt=FMT_M)
            put(ws, f"{L(last_col)}{r}", "Ties to the MembershipMember disaggregated-revenue fact (FY2025: $5,323M, see Supporting facts). Carries no cost of sales; falls straight to operating income.", font=F_NOTE)
            name(wb, "Costco_MembershipFees", f"'Raw Data'!$C${r}:${ycol[2025]}${r}")
            r += 1
        if t == "WFM":
            put(ws, f"A{r}", "Successor: Amazon.com 'Physical stores' net sales (memo)", font=F_N, align=Alignment(indent=1))
            put(ws, f"B{r}", "RevenueFromContractWithCustomerExcludingAssessedTax [ProductOrServiceAxis = PhysicalStoresMember]", font=F_NOTE)
            for y in YEARS:
                cell = f"{ycol[y]}{r}"
                if y in d["successor"].index:
                    put(ws, cell, float(d["successor"].loc[y, "value"]) / 1e6, font=F_IN, fmt=FMT_M)
                else:
                    put(ws, cell, "n/a", font=F_NA, align=RIGHT)
            put(ws, f"{L(last_col)}{r}", "Amazon 10-K, Note – Segment Information / disaggregated revenue (ASC 606). Mostly Whole Foods Market; "
                                          "also Amazon Fresh and Amazon Go. Revenue only – no cost or balance-sheet data exists for Whole Foods after FY2017.", font=F_NOTE)
            name(wb, "WholeFoods_SuccessorRevenue", f"'Raw Data'!$C${r}:${ycol[2025]}${r}")
            r += 1
        r += 1

    # supporting facts cited in the Executive Summary
    header_band(ws, r, "Supporting facts cited in the Executive Summary (USD millions unless stated)", 1, last_col, fill=FILL_T)
    r += 1
    for c, h in zip("ABCDEF", ["Fact", "XBRL tag", "Fiscal year", "Value", "Accession no.", "Filed"]):
        put(ws, f"{c}{r}", h, font=F_B, fill=FILL_L, border=BOX, align=CENTER if c != "A" else Alignment(indent=1))
    r += 1
    driver_names = {}
    for _, f in D["drivers"].iterrows():
        nm = f"{PREFIX[f.ticker]}_{''.join(w.capitalize() for w in str(f.label).replace('(', ' ').replace(')', ' ').replace('-', ' ').split()[:4])}_FY{f.fiscal_year}"
        nm = nm.replace("'", "")
        put(ws, f"A{r}", f"{DISPLAY[f.ticker]} – {f.label}", font=F_N, align=Alignment(indent=1), border=BOX)
        put(ws, f"B{r}", f.tag, font=F_NOTE, border=BOX)
        put(ws, f"C{r}", int(f.fiscal_year), font=F_N, fmt=FMT_FY, border=BOX, align=CENTER)
        put(ws, f"D{r}", float(f.value) / 1e6, font=F_IN, fmt=FMT_M, border=BOX)
        put(ws, f"E{r}", f.accession, font=F_NOTE, border=BOX)
        put(ws, f"F{r}", f.filed, font=F_NOTE, border=BOX, align=CENTER)
        name(wb, nm, f"'Raw Data'!$D${r}")
        driver_names[(f.ticker, f.label, int(f.fiscal_year))] = nm
        r += 1
    r += 1
    put(ws, f"A{r}", "Legend: blue = SEC input · black bold = formula · shaded yellow = value taken from the XBRL instance document or derived from two "
                     "tagged facts (see note) · gray band = debt-build components · 'n/a' = not reported.", font=F_NOTE)
    return {"ws": ws, "rowmap": rowmap, "ycol": ycol, "drivers": driver_names}


# ------------------------------------------------------------------ Ratio Calculations
RATIOS = [
    # key, title, definition, format, delta format, higher_is_better, template(prefix)->formula body
    ("GM", "Gross profit margin", "(Revenue − Cost of sales) / Revenue", FMT_PCT, FMT_PP, True,
     lambda p, y: f"({X(y, p, 'Revenue')}-{X(y, p, 'COGS')})/{X(y, p, 'Revenue')}"),
    ("OM", "Operating margin (supplementary)", "Operating income / Revenue – separates gross-margin pressure from SG&A", FMT_PCT, FMT_PP, True,
     lambda p, y: f"{X(y, p, 'OperatingIncome')}/{X(y, p, 'Revenue')}"),
    ("NM", "Net profit margin", "Net income / Revenue", FMT_PCT, FMT_PP, True,
     lambda p, y: f"{X(y, p, 'NetIncome')}/{X(y, p, 'Revenue')}"),
    ("CR", "Current ratio", "Current assets / Current liabilities", FMT_X, FMT_DX, True,
     lambda p, y: f"{X(y, p, 'CurrentAssets')}/{X(y, p, 'CurrentLiabilities')}"),
    ("DE", "Debt-to-equity", "Total debt / Total equity (finance leases in, operating leases out – ASC 842 policy)", FMT_X, FMT_DX, False,
     lambda p, y: f"{X(y, p, 'TotalDebt')}/{X(y, p, 'TotalEquity')}"),
    ("DEL", "Debt-to-equity incl. operating leases (supplementary)", "(Total debt + Operating lease liability) / Total equity – lease-adjusted view from FY2019", FMT_X, FMT_DX, False,
     lambda p, y: f"({X(y, p, 'TotalDebt')}+{X(y, p, 'OperatingLease')})/{X(y, p, 'TotalEquity')}"),
    ("GR", "Revenue growth YoY", "Revenue / prior-year revenue − 1 (FY2015 base year on Raw Data)", FMT_PCT, FMT_PP, True,
     lambda p, y: f"{X(y, p, 'Revenue')}/_xlfn.XLOOKUP({y}-1,Years,{p}_Revenue)-1"),
]


def X(y, p, suffix):
    return f"_xlfn.XLOOKUP({y},Years,{p}_{suffix})"


def build_ratios(wb) -> dict:
    ws = wb.create_sheet("Ratio Calculations")
    sheet_setup(ws, TEAL)
    ws.column_dimensions["A"].width = 30
    ncol = len(RYEARS)
    for i in range(ncol):
        ws.column_dimensions[L(2 + i)].width = 10.5
    ws.column_dimensions[L(2 + ncol)].width = 13
    ws.column_dimensions[L(3 + ncol)].width = 13
    ws.column_dimensions[L(4 + ncol)].width = 60
    ycol = {y: L(2 + i) for i, y in enumerate(RYEARS)}
    dcol1, dcol2, ncol_notes = L(2 + ncol), L(3 + ncol), L(4 + ncol)

    title_band(ws, 4, "Ratio Calculations", "Live formulas over named ranges: every cell is =IFERROR( … XLOOKUP(year, Years, Company_Metric) … , \"n/a\"). "
                                             "Nothing is typed in. Δ columns are percentage-point (margins, growth) or turns (ratios) changes.", 4 + ncol - 1)
    put(ws, "A4", "Company", font=F_H, fill=FILL_H, align=Alignment(vertical="center", indent=1))
    for y in RYEARS:
        put(ws, f"{ycol[y]}4", y, font=F_H, fill=FILL_H, fmt=FMT_FY, align=CENTER)
    put(ws, f"{dcol1}4", "Δ FY16→FY25", font=F_H, fill=FILL_H, align=CENTER)
    put(ws, f"{dcol2}4", "Δ FY19→FY23", font=F_H, fill=FILL_H, align=CENTER)
    put(ws, f"{ncol_notes}4", "Formula (FY2025 cell shown)", font=F_H, fill=FILL_H, align=Alignment(vertical="center", indent=1))
    ws.row_dimensions[4].height = 22
    name(wb, "RatioYears", f"'Ratio Calculations'!$B$4:${ycol[2025]}$4")
    ws.freeze_panes = "B5"

    rows: dict[str, dict[str, int]] = {}
    r = 6
    for key, title, definition, fmt, dfmt, hib, body in RATIOS:
        header_band(ws, r, f"{title}   —   {definition}", 1, 4 + ncol - 1, fill=FILL_L, font=F_B, height=18)
        r += 1
        rows[key] = {}
        members = COMPANIES + ([("WFMS", "Whole Foods (Amazon 'Physical stores', successor)", "WholeFoods")] if key == "GR" else [])
        for t, disp, pre in members:
            put(ws, f"A{r}", disp, font=F_N, align=Alignment(indent=1))
            for y in RYEARS:
                yref = f"{ycol[y]}$4"
                if t == "WFMS":
                    f = f"=IFERROR({X(yref, pre, 'SuccessorRevenue')}/_xlfn.XLOOKUP({yref}-1,Years,WholeFoods_SuccessorRevenue)-1,\"n/a\")"
                else:
                    f = f"=IFERROR({body(pre, yref)},\"n/a\")"
                put(ws, f"{ycol[y]}{r}", f, font=F_N, fmt=fmt, align=RIGHT)
            put(ws, f"{dcol1}{r}", f"=IFERROR({ycol[2025]}{r}-{ycol[2016]}{r},\"n/a\")", font=F_N, fmt=dfmt, align=RIGHT, fill=FILL_P)
            put(ws, f"{dcol2}{r}", f"=IFERROR({ycol[2023]}{r}-{ycol[2019]}{r},\"n/a\")", font=F_N, fmt=dfmt, align=RIGHT, fill=FILL_P)
            put(ws, f"{ncol_notes}{r}", ws[f"{ycol[2025]}{r}"].value.replace("_xlfn.", ""), font=F_NOTE)
            rows[key][t] = r
            if t != "WFMS":
                name(wb, f"{key}_{pre}", f"'Ratio Calculations'!$B${r}:${ycol[2025]}${r}")
            r += 1
        r += 1
    # n/a cells in gray italic, whole block
    ws.conditional_formatting.add(f"B6:{dcol2}{r}", FormulaRule(formula=["ISTEXT(B6)"], font=Font(italic=True, color=MID)))
    put(ws, f"A{r}", "Sam's Club: no balance sheet or net income at segment level, so current ratio, debt-to-equity and net margin are n/a by construction. "
                     "Whole Foods: standalone data ends FY2017; the successor growth row uses Amazon's 'Physical stores' line (mostly, not only, Whole Foods).", font=F_NOTE)
    return {"ws": ws, "rows": rows, "ycol": ycol}


# ------------------------------------------------------------------ Scorecard
def build_scorecard(wb, R) -> dict:
    ws = wb.create_sheet("Scorecard")
    sheet_setup(ws, NAVY, gridlines=False)
    widths = {"A": 34, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 13, "I": 13, "J": 90}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    title_band(ws, 5, "Scorecard", "Direction-aware ranks and color scales (green = strongest) · change the year to re-rank", 10, subtitle_in_band=True)
    ws.row_dimensions[3].height = 22
    put(ws, "A3", "Scorecard year (input)", font=F_B)
    put(ws, "B3", 2025, font=Font(name="Calibri", size=12, bold=True, color=INPUT_BLUE), fmt=FMT_FY, fill=PatternFill("solid", fgColor="FFF2CC"), border=BOX, align=CENTER)
    name(wb, "ScorecardYear", "'Scorecard'!$B$3")
    dv = DataValidation(type="list", formula1='"2016,2017,2018,2019,2020,2021,2022,2023,2024,2025"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B3")
    put(ws, "C3", "← any year FY2016–FY2025", font=F_NOTE)

    hdr = ["Metric", "Better is"] + [d for _, d, _ in COMPANIES] + ["Leader", "Laggard", "Takeaway"]
    for i, h in enumerate(hdr):
        put(ws, f"{L(1 + i)}5", h, font=F_H, fill=FILL_H, align=CENTER if i else Alignment(vertical="center", indent=1))
    ws.row_dimensions[5].height = 22
    ws.freeze_panes = "A6"

    takeaways = {
        "GM": "Level reflects business model, not execution: warehouse clubs (Costco, Sam's Club) run 10–12% on net sales, supercenters and supermarkets 22–25%. Compare each company's trend on the Trend tab, not levels across models.",
        "OM": "Operating margin isolates cost discipline below gross profit. Costco's has risen steadily; Kroger's FY2025 figure carries $2.7B of restructuring and impairment charges.",
        "NM": "The cleanest cross-model comparison. Costco converts the thinnest gross margin into the highest net margin through membership fees and SG&A control.",
        "CR": "Grocers and clubs run near or below 1.0x by design – inventory turns faster than suppliers are paid. A ratio above 1.0x is rare in this set and is Costco's alone.",
        "DE": "Kroger's debt-to-equity is a multiple of its peers' after the FY2024 $10.5B debt issue and $6.9B of FY2024–25 buybacks; Costco has de-levered every year since FY2021.",
        "GR": "Costco has led growth in nearly every year of the window; Kroger's revenue has been flat to down since FY2022.",
    }
    metric_rows = {}
    r = 6
    for key, title, definition, fmt, dfmt, hib, _ in RATIOS:
        if key in ("DEL",):
            continue
        put(ws, f"A{r}", title.replace(" (supplementary)", ""), font=Font(name="Calibri", size=11, bold=True), align=Alignment(vertical="center", indent=1), border=BOX)
        put(ws, f"B{r}", "Higher" if hib else "Lower", font=F_N, align=CENTER, border=BOX)
        for i, (t, d, pre) in enumerate(COMPANIES):
            col = L(3 + i)
            put(ws, f"{col}{r}", f"=IFERROR(_xlfn.XLOOKUP(ScorecardYear,RatioYears,{key}_{pre}),\"n/a\")", font=F_N, fmt=fmt, align=CENTER, border=BOX)
        rng = f"C{r}:G{r}"
        best, worst = ("MAX", "MIN") if hib else ("MIN", "MAX")
        put(ws, f"H{r}", f"=IFERROR(INDEX($C$5:$G$5,MATCH({best}({rng}),{rng},0)),\"n/a\")", font=Font(name="Calibri", size=10, bold=True, color="375623"), align=CENTER, border=BOX)
        put(ws, f"I{r}", f"=IFERROR(INDEX($C$5:$G$5,MATCH({worst}({rng}),{rng},0)),\"n/a\")", font=Font(name="Calibri", size=10, bold=True, color="9C0006"), align=CENTER, border=BOX)
        def T(expr):
            return f'TEXT({expr},"0.0%")' if fmt == FMT_PCT else f'TEXT({expr},"0.00")&"x"'
        lead, trail = T(f"{best}({rng})"), T(f"{worst}({rng})")
        put(ws, f"J{r}", f'="FY"&ScorecardYear&": "&H{r}&" leads at "&{lead}&", "&I{r}&" trails at "&{trail}&". {takeaways[key]}"',
            font=F_N, align=Alignment(wrap_text=True, vertical="center"), border=BOX)
        ws.row_dimensions[r].height = 46
        # rank row
        put(ws, f"A{r + 1}", "   rank (of companies with data)", font=F_NOTE, align=Alignment(indent=2), border=BOX)
        put(ws, f"B{r + 1}", "", border=BOX)
        for i in range(5):
            col = L(3 + i)
            put(ws, f"{col}{r + 1}", f"=IFERROR(\"#\"&_xlfn.RANK.EQ({col}{r},{rng},{0 if hib else 1})&\" of \"&COUNT({rng}),\"–\")", font=F_NOTE, align=CENTER, border=BOX)
        for c in "HIJ":
            put(ws, f"{c}{r + 1}", "", border=BOX)
        lo, hi = ("F8696B", "63BE7B") if hib else ("63BE7B", "F8696B")
        ws.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color=lo, mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color=hi))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"ISTEXT(C{r})"], font=Font(italic=True, color=MID), fill=PatternFill("solid", fgColor="FFFFFF")))
        metric_rows[key] = r
        r += 2
    # composite
    r += 1
    put(ws, f"A{r}", "Composite rank (average of five core ranks)", font=Font(name="Calibri", size=11, bold=True), align=Alignment(indent=1, vertical="center", wrap_text=True), border=BOX)
    put(ws, f"B{r}", "Lower", font=F_N, align=CENTER, border=BOX)
    core = [metric_rows[k] for k in ("GM", "NM", "CR", "DE", "GR")]
    for i, (t, d, pre) in enumerate(COMPANIES):
        col = L(3 + i)
        parts = ",".join(f"IFERROR(_xlfn.RANK.EQ({col}{rr},C{rr}:G{rr},{0 if hib else 1}),\"\")" for rr, hib in zip(core, [True, True, True, False, True]))
        put(ws, f"{col}{r}", f"=IFERROR(AVERAGE({parts}),\"n/a\")", font=Font(name="Calibri", size=11, bold=True), fmt="0.0", align=CENTER, border=BOX)
    put(ws, f"H{r}", f"=IFERROR(INDEX($C$5:$G$5,MATCH(MIN(C{r}:G{r}),C{r}:G{r},0)),\"n/a\")", font=Font(name="Calibri", size=10, bold=True, color="375623"), align=CENTER, border=BOX)
    put(ws, f"I{r}", f"=IFERROR(INDEX($C$5:$G$5,MATCH(MAX(C{r}:G{r}),C{r}:G{r},0)),\"n/a\")", font=Font(name="Calibri", size=10, bold=True, color="9C0006"), align=CENTER, border=BOX)
    put(ws, f"J{r}", "Computed only for the three full filers, which have all five core metrics. Sam's Club and Whole Foods lack the balance-sheet metrics and are not ranked.",
        font=F_NOTE, align=Alignment(wrap_text=True, vertical="center"), border=BOX)
    ws.row_dimensions[r].height = 40
    ws.conditional_formatting.add(f"C{r}:G{r}", ColorScaleRule(start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="F8696B"))
    ws.conditional_formatting.add(f"C{r}:G{r}", FormulaRule(formula=[f"ISTEXT(C{r})"], font=Font(italic=True, color=MID), fill=PatternFill("solid", fgColor="FFFFFF")))
    r += 2
    put(ws, f"A{r}", "Sam's Club: net margin, current ratio and debt-to-equity are n/a (segment reporting). Whole Foods: n/a after FY2017 (acquired). "
                     "Operating margin is supplementary and excluded from the composite.", font=F_NOTE)
    return {"ws": ws, "rows": metric_rows}


# ------------------------------------------------------------------ Trend
def build_trend(wb, R):
    ws = wb.create_sheet("Trend")
    sheet_setup(ws, NAVY, gridlines=False)
    title_band(ws, 6, "Trend", "FY2016–FY2025. Shaded band = FY2020–FY2023 (COVID-19 and the inflation spike). Gaps = data not available (Sam's Club balance sheet, "
                               "Whole Foods after FY2017); charts read the helper block via NA() so missing years are skipped, never plotted as zero.", 20)
    rws = R["ws"].title
    ry = R["ycol"]
    helper_top = 70
    put(ws, f"A{helper_top - 1}", "Chart helper block (formulas; do not edit): =IF(ISNUMBER(ratio), ratio, NA()) — one row per series, plus the shaded-window series",
        font=F_B)
    put(ws, f"A{helper_top}", "Series", font=F_H, fill=FILL_H)
    for i, y in enumerate(RYEARS):
        put(ws, f"{L(2 + i)}{helper_top}", f"FY{y}", font=F_H, fill=FILL_H, align=CENTER)
    cats = Reference(ws, min_col=2, max_col=1 + len(RYEARS), min_row=helper_top)
    r = helper_top + 1
    charts = [("GM", "Gross profit margin", "0%"), ("NM", "Net profit margin", "0.0%"), ("OM", "Operating margin", "0.0%"),
              ("CR", "Current ratio", '0.00"x"'), ("DE", "Debt-to-equity", '0.0"x"'), ("GR", "Revenue growth YoY", "0%")]
    anchors = ["A4", "K4", "A25", "K25", "A46", "K46"]
    for (key, title, nf), anchor in zip(charts, anchors):
        first = r
        members = [(t, d) for t, d, _ in COMPANIES if t in R["rows"][key] and not (t == "SAMS" and key in ("NM", "CR", "DE"))]
        for t, d in members:
            put(ws, f"A{r}", d, font=F_N)
            src = R["rows"][key][t]
            for i, y in enumerate(RYEARS):
                put(ws, f"{L(2 + i)}{r}", f"=IF(ISNUMBER('{rws}'!{ry[y]}{src}),'{rws}'!{ry[y]}{src},NA())", font=F_NOTE, fmt=nf)
            r += 1
        put(ws, f"A{r}", "FY2020–23 window", font=F_NOTE)
        for i, y in enumerate(RYEARS):
            put(ws, f"{L(2 + i)}{r}", f"=IF(AND({L(2 + i)}${helper_top}>=\"FY2020\",{L(2 + i)}${helper_top}<=\"FY2023\"),1,NA())", font=F_NOTE)
        shock = r
        r += 2

        # Combo chart written canonically (shared axes). src/recalc_excel.py then moves the
        # band series to a hidden secondary axis scaled 0-1 via Excel COM - openpyxl's own
        # secondary-axis XML is read by Excel as "deleted primary", so it is done in Excel.
        bar = BarChart()
        bar.type = "col"
        bar.grouping = "standard"
        bar.gapWidth = 0
        bar.add_data(Reference(ws, min_col=1, max_col=1 + len(RYEARS), min_row=shock), from_rows=True, titles_from_data=True)
        bar.set_categories(cats)
        s = bar.series[0]
        s.graphicalProperties = GraphicalProperties(solidFill=BAND)
        s.graphicalProperties.line.noFill = True
        bar.x_axis.delete = False
        bar.x_axis.tickLblPos = "low"
        bar.y_axis.delete = False
        bar.y_axis.number_format = nf
        bar.y_axis.majorGridlines = ChartLines()
        bar.y_axis.title = title
        bar.y_axis.title.overlay = False

        line = LineChart()
        line.add_data(Reference(ws, min_col=1, max_col=1 + len(RYEARS), min_row=first, max_row=first + len(members) - 1), from_rows=True, titles_from_data=True)
        line.set_categories(cats)
        for (t, _), ls in zip(members, line.series):
            col = SERIES_COLOR[t]
            ls.graphicalProperties = GraphicalProperties()
            ls.graphicalProperties.line.solidFill = col
            ls.graphicalProperties.line.width = 28575
            ls.marker.symbol = "circle"
            ls.marker.size = 5
            ls.marker.graphicalProperties = GraphicalProperties(solidFill=col)
            ls.marker.graphicalProperties.line.solidFill = col
            ls.smooth = False
        line.display_blanks = "gap"
        bar += line
        bar.title = f"{title}, FY2016–FY2025"
        bar.title.overlay = False
        bar.legend.position = "b"
        bar.legend.overlay = False
        bar.plot_area.layout = Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge", x=0.09, y=0.12, w=0.88, h=0.66))
        bar.height = 10
        bar.width = 16
        ws.add_chart(bar, anchor)
    # chart grid: each chart is ~16 cm wide; columns are sized so the second column of
    # charts (anchored at K) starts clear of the first, and rows so the next row of
    # charts (21 rows later) starts clear of the one above.
    ws.print_area = "A1:T66"
    ws.column_dimensions["A"].width = 18
    for i in range(2, 24):
        ws.column_dimensions[L(i)].width = 9
    for r in range(4, helper_top - 2):
        ws.row_dimensions[r].height = 15
    return ws


# ------------------------------------------------------------------ Executive Summary
def R_(y, nm):      # ratio lookup
    return f"_xlfn.XLOOKUP({y},RatioYears,{nm})"


def V_(y, nm):      # raw value lookup ($M)
    return f"_xlfn.XLOOKUP({y},Years,{nm})"


def W(nm):        # FY2020-FY2023 slice of a ratio row (positions 5-8 of FY2016-FY2025)
    return f"INDEX({nm},1,5):INDEX({nm},1,8)"


def pct(expr):
    return f'TEXT({expr},"0.0%")'


def x2(expr):
    return f'TEXT({expr},"0.00")&"x"'


def bn(expr):
    return f'"$"&TEXT({expr}/1000,"0.0")&"B"'


def build_summary(wb, RAW, S):
    ws = wb.create_sheet("Executive Summary")
    sheet_setup(ws, NAVY, gridlines=False)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 150
    title_band(ws, 2, "Executive Summary", "Findings are stated against the three key questions. Every figure quoted below is a live formula into the Ratio "
                                           "Calculations and Raw Data tabs; cited transactions (debt issued, buybacks, charges) are on Raw Data under 'Supporting facts'.", 2, text_col=2)
    ws.row_dimensions[3].height = 30
    ws["B3"].alignment = WRAP

    # at-a-glance table (rows 4-10) in columns B.. we need more columns; use a small table to the right? keep in B with merged? simpler: dedicated columns C..H are hidden by width; instead use rows.
    r = 4
    dn = RAW["drivers"]
    kr_debt = dn[("KR", "Proceeds from issuance of long-term debt", 2024)]
    kr_bb24 = dn[("KR", "Share repurchases", 2024)]
    kr_bb25 = dn[("KR", "Share repurchases", 2025)]
    kr_chg = dn[("KR", "Restructuring and asset impairment charges", 2025)]
    kr_lifo = dn[("KR", "LIFO reserve", 2025)]
    kr_lifo16 = dn[("KR", "LIFO reserve", 2016)]
    wmt_nci = dn[("WMT", "Noncontrolling interest (in total equity)", 2025)]
    cost_fees = dn[("COST", "Membership fee revenue (included in total revenue)", 2025)]

    sections = [
        ("1. Capital structure & financing strategy — who is financed most conservatively, and has that changed?", [
            "BOTTOM LINE: Costco. Its leverage is the lowest in the group and has fallen every year since FY2021. Walmart has not changed in a decade. Kroger has tripled its leverage since FY2023 to buy back stock after the Albertsons deal failed – deliberate, but the one red flag here.",
            f'="Costco: debt-to-equity "&{x2(R_(2016, "DE_Costco"))}&" in FY2016, "&{x2(R_(2025, "DE_Costco"))}&" in FY2025. Debt has stayed near "&{bn(V_(2025, "Costco_TotalDebt"))}&"; equity has grown from "&{bn(V_(2016, "Costco_TotalEquity"))}&" to "&{bn(V_(2025, "Costco_TotalEquity"))}&". It funds growth and buybacks from cash flow, not borrowing."',
            f'="Walmart: debt-to-equity has stayed between "&{x2("MIN(DE_Walmart)")}&" and "&{x2("MAX(DE_Walmart)")}&" for ten years. Debt is up "&{pct(f"{V_(2025, 'Walmart_TotalDebt')}/{V_(2016, 'Walmart_TotalDebt')}-1")}&" since FY2016 on revenue that is up "&{pct(f"{V_(2025, 'Walmart_Revenue')}/{V_(2016, 'Walmart_Revenue')}-1")}&". Leverage is falling relative to the size of the business."',
            f'="Kroger: debt-to-equity fell from "&{x2(R_(2016, "DE_Kroger"))}&" (FY2016) to "&{x2(R_(2023, "DE_Kroger"))}&" (FY2023), then rose to "&{x2(R_(2024, "DE_Kroger"))}&" in FY2024 and "&{x2(R_(2025, "DE_Kroger"))}&" in FY2025. Cause: "&{bn(kr_debt)}&" of debt issued in FY2024 for the Albertsons merger (terminated December 2024), "&{bn(kr_bb24)}&" of buybacks that year and "&{bn(kr_bb25)}&" more in FY2025. Equity fell from "&{bn(V_(2023, "Kroger_TotalEquity"))}&" to "&{bn(V_(2025, "Kroger_TotalEquity"))}&" in two years."',
            '="Why it matters: Kroger is borrowing to return capital, not to grow – revenue has been flat since FY2022. That is a choice, not a mistake, but it leaves Kroger with the least room of the three if margins slip."',
            f'="Including operating leases (ASC 842) changes no ranking: FY2025 debt-to-equity is "&{x2(R_(2025, "DEL_Costco"))}&" Costco, "&{x2(R_(2025, "DEL_Walmart"))}&" Walmart, "&{x2(R_(2025, "DEL_Kroger"))}&" Kroger."',
        ]),
        ("2. Resilience through 2020–2023 — who absorbed COVID and the inflation spike best, and who took the biggest hit?", [
            "BOTTOM LINE: Costco absorbed it best – sales, net margin and liquidity all rose through the window. Walmart took the biggest hit, in FY2022, and was back to normal within two years. Kroger's margin swings were fuel mix, not stress.",
            f'="Costco: net sales grew "&{pct(R_(2020, "GR_Costco"))}&", "&{pct(R_(2021, "GR_Costco"))}&" and "&{pct(R_(2022, "GR_Costco"))}&" in FY2020–FY2022. Net margin rose from "&{pct(R_(2019, "NM_Costco"))}&" to "&{pct(R_(2023, "NM_Costco"))}&". Current ratio peaked at "&{x2(R_(2020, "CR_Costco"))}&" and never fell below "&{x2("MIN(CR_Costco)")}&". Gross margin dipped to "&{pct(R_(2022, "GM_Costco"))}&" in FY2022 as Costco held prices through inflation, then recovered to "&{pct(R_(2025, "GM_Costco"))}&"."',
            f'="Walmart: FY2022 (year ended January 2023) was the low point of the decade – gross margin "&{pct(R_(2022, "GM_Walmart"))}&", operating margin "&{pct(R_(2022, "OM_Walmart"))}&" – as excess inventory was marked down. Current ratio fell from "&{x2(R_(2020, "CR_Walmart"))}&" to "&{x2(R_(2022, "CR_Walmart"))}&". By FY2024 operating margin was "&{pct(R_(2024, "OM_Walmart"))}&" and FY2025 net margin "&{pct(R_(2025, "NM_Walmart"))}&" – fully recovered."',
            f'="Kroger: gross margin jumped to "&{pct(R_(2020, "GM_Kroger"))}&" in FY2020 when low-margin fuel sales collapsed, then fell to "&{pct(R_(2022, "GM_Kroger"))}&" in FY2022 when fuel prices spiked. Net margin held between "&{pct(f"MIN({W('NM_Kroger')})")}&" and "&{pct(f"MAX({W('NM_Kroger')})")}&". Its current ratio ("&{x2(f"MIN({W('CR_Kroger')})")}&"–"&{x2(f"MAX({W('CR_Kroger')})")}&") is the lowest of the three by design, not because of the shock."',
            f'="Sam\'s Club: net sales grew "&{pct(R_(2020, "GR_SamsClub"))}&", "&{pct(R_(2021, "GR_SamsClub"))}&" and "&{pct(R_(2022, "GR_SamsClub"))}&" in FY2020–FY2022 with operating margin between "&{pct(f"MIN({W('OM_SamsClub')})")}&" and "&{pct(f"MAX({W('OM_SamsClub')})")}&". The club format won the pandemic and the trade-down that followed."',
        ]),
        ("3. Margin pressure since 2016 — who has felt it most, and is it cost, pricing, or mix?", [
            "BOTTOM LINE: Walmart – and it is pricing and mix, not cost. Gross margin has given up ground while SG&A has held. Kroger's operating margin has fallen further, but that is cost growth and one-off charges below gross profit. Costco shows no pressure. Sam's Club is improving.",
            f'="Walmart: gross margin "&{pct(R_(2016, "GM_Walmart"))}&" in FY2016, "&{pct(R_(2022, "GM_Walmart"))}&" at the FY2022 low, "&{pct(R_(2025, "GM_Walmart"))}&" in FY2025; operating margin "&{pct(R_(2016, "OM_Walmart"))}&" to "&{pct(R_(2025, "OM_Walmart"))}&". The gross-margin give-up matches Walmart\'s price investment and the shift to lower-margin e-commerce and grocery. The gap between gross and operating margin has not widened, so costs are under control."',
            f'="Kroger: operating margin "&{pct(R_(2016, "OM_Kroger"))}&" in FY2016 to "&{pct(R_(2025, "OM_Kroger"))}&" in FY2025 – the largest decline in the group. FY2025 includes "&{bn(kr_chg)}&" of restructuring and impairment charges; without them operating margin would be about "&{pct(f"({V_(2025, 'Kroger_OperatingIncome')}+{kr_chg})/{V_(2025, 'Kroger_Revenue')}")}&". Gross margin is unchanged ("&{pct(R_(2016, "GM_Kroger"))}&" to "&{pct(R_(2025, "GM_Kroger"))}&"), so the pressure is cost, not pricing."',
            f'="Costco: gross margin flat by design ("&{pct(R_(2016, "GM_Costco"))}&" to "&{pct(R_(2025, "GM_Costco"))}&"); net margin up from "&{pct(R_(2016, "NM_Costco"))}&" to "&{pct(R_(2025, "NM_Costco"))}&". Membership fees ("&{bn(cost_fees)}&" in FY2025, "&{pct(f"{cost_fees}/{V_(2025, 'Costco_Revenue')}")}&" of net sales, no cost of sales) and "&{pct(f"{V_(2025, 'Costco_Revenue')}/{V_(2016, 'Costco_Revenue')}-1")}&" cumulative sales growth do the work."',
            f'="Sam\'s Club: segment gross margin up from "&{pct(R_(2022, "GM_SamsClub"))}&" (FY2022) to "&{pct(R_(2025, "GM_SamsClub"))}&" (FY2025) as mix shifts away from fuel. Whole Foods was already under pricing pressure before Amazon bought it: gross margin "&{pct(R_(2016, "GM_WholeFoods"))}&" to "&{pct(R_(2017, "GM_WholeFoods"))}&", operating margin "&{pct(R_(2016, "OM_WholeFoods"))}&" to "&{pct(R_(2017, "OM_WholeFoods"))}&" in its last two standalone years."',
        ]),
        ("Notable trend shifts since FY2016", [
            f'="Costco has changed weight class. Net sales grew "&{pct(R_(2021, "GR_Costco"))}&" and "&{pct(R_(2022, "GR_Costco"))}&" in FY2021–FY2022, taking it from "&{bn(V_(2016, "Costco_Revenue"))}&" to "&{bn(V_(2025, "Costco_Revenue"))}&". It is now "&TEXT({V_(2025, "Costco_Revenue")}/{V_(2025, "Kroger_Revenue")},"0.0")&"x Kroger\'s size; it was "&TEXT({V_(2016, "Costco_Revenue")}/{V_(2016, "Kroger_Revenue")},"0.0")&"x in FY2016."',
            f'="Kroger has reversed course. Five years of de-leveraging (FY2018–FY2023) were undone in two: debt-to-equity "&{x2(R_(2023, "DE_Kroger"))}&" to "&{x2(R_(2025, "DE_Kroger"))}&", on revenue growth of "&{pct(R_(2024, "GR_Kroger"))}&" and "&{pct(R_(2025, "GR_Kroger"))}&"."',
            '="ASC 842 (FY2019; Costco FY2020) put operating leases on every balance sheet. Total debt here excludes them so FY2016–FY2018 stay comparable; the lease-inclusive ratio above shows the effect."',
        ]),
        ("Data caveats", [
            "Sam's Club is a segment of Walmart (ASC 280), not a filer. Net sales, operating income and total assets are available FY2016–FY2025; cost of revenue only from FY2022 (ASU 2023-07); no net income, current ratio or debt exists at segment level.",
            "Whole Foods filed its last 10-K for FY2017 (Amazon closed the acquisition 2017-08-28). From FY2018 the only public figure is Amazon's 'Physical stores' net sales, which also includes Amazon Fresh and Amazon Go. Whole Foods' cost line included occupancy costs, so its ~34% gross margin is not comparable to peers.",
            "Costco revenue is NET SALES, excluding membership fees, to match Walmart's and Sam's Club's basis. Total revenue and the fee line are memo rows on Raw Data.",
            f'="Inventory method: Walmart U.S., Sam\'s Club, Costco U.S. and ~91% of Kroger use LIFO; Costco Canada and Walmart International use FIFO. Kroger\'s LIFO reserve is "&{bn(kr_lifo)}&" at FY2025 ("&{bn(kr_lifo16)}&" at FY2016), so its reported inventory and cost of sales are below FIFO values. Walmart reported LIFO ≈ FIFO through January 2022 and gives only sensitivity language since; Costco took a LIFO charge worth 19 bp of gross margin in FY2022. Ratios are as reported; the Methodology tab quantifies the effect."',
            f'="Total equity includes noncontrolling interests (Walmart: "&{bn(wmt_nci)}&" at FY2025); net income is the amount attributable to the parent. Standard presentation; it slightly lowers Walmart\'s debt-to-equity versus a parent-only basis."',
            "53-week years: Costco and Kroger FY2017 and FY2023 each had 53 weeks, which flatters growth in those years and depresses it the year after by roughly 2 points. Fiscal years follow the SEC frame convention, so Walmart's 'fiscal 2026' (ended January 2026) is FY2025 here.",
            "Every input carries its XBRL tag, accession number and filing date on the Data Lineage tab. Seven values the SEC API could not return were taken from the filing's XBRL instance or derived from two tagged facts; they are shaded yellow on Raw Data.",
        ]),
    ]
    for title, bullets in sections:
        header_band(ws, r, title, 2, 2, fill=FILL_H, font=F_H, height=22)
        r += 1
        for b in bullets:
            b = b.replace('="• ', '="').replace("• ", "", 1) if b.startswith('="• ') or b.startswith("• ") else b
            if b.startswith("BOTTOM LINE:"):
                put(ws, f"B{r}", b, font=Font(name="Calibri", size=10.5, bold=True, color=NAVY), align=WRAP, fill=FILL_P)
                ws.row_dimensions[r].height = 14.5 * max(1, math.ceil(len(b) / 200))
                r += 1
                continue
            txt = '="• ' + b[2:] if b.startswith('="') else "• " + b
            put(ws, f"B{r}", txt, font=F_N, align=WRAP)
            if b.startswith("="):   # rendered length = string literals + ~6 chars per formatted number
                rendered = sum(len(m) for m in re.findall(r'"([^"]*)"', b)) + 6 * b.count("TEXT(")
            else:
                rendered = len(b)
            ws.row_dimensions[r].height = 14.5 * max(1, math.ceil(rendered / 205))
            r += 1
        r += 1
    return ws


# ------------------------------------------------------------------ Methodology
def build_method(wb, counts):
    ws = wb.create_sheet("Methodology & Tools")
    sheet_setup(ws, TEAL, gridlines=False)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 140
    title_band(ws, 7, "Methodology & Tools", "How the workbook was built, what was assumed, and where accounting policy limits comparability.", 2)
    items = [
        ("DATA SOURCES", None),
        ("SEC EDGAR", "Primary source for every number. Three endpoints: (1) the XBRL companyfacts API (data.sec.gov/api/xbrl/companyfacts) for undimensioned "
                      "10-K facts; (2) the filing index and XBRL instance documents under sec.gov/Archives for dimensional facts (segments, product lines) "
                      "and for values the API omits; (3) the submissions API for the 10-K filing list. Access follows the SEC fair-use rules: declared "
                      "User-Agent, ≤5 requests/second, every download cached and logged to data/raw/manifest.csv."),
        ("Companies / CIKs", "Walmart Inc. 104169 · Costco Wholesale Corp. 909832 · The Kroger Co. 56873 · Whole Foods Market, Inc. 865436 (through FY2017) · "
                             "Amazon.com, Inc. 1018724 (successor, 'Physical stores' revenue only) · Sam's Club = Walmart segment (no CIK)."),
        ("Value selection", "Fiscal-year label follows the SEC frame convention (the calendar year containing most of the period). Flow items use 10-K facts "
                            "spanning 340–380 days (53-week years qualify); balance-sheet items match the period-end date. Where a period is re-tagged in later "
                            "filings as a comparative, the latest-filed (restated) value is used and the first-filed value is retained on Data Lineage with a "
                            "restatement flag. Values the API could not supply (7 of 900+) were read from the filing's XBRL instance or derived from two tagged "
                            "facts, each with its own citation, and are shaded yellow on Raw Data. No value is estimated."),
        ("EXCEL TECHNIQUES", None),
        ("Named ranges", f"{counts['names']} defined names (e.g. Walmart_Revenue, Kroger_TotalDebt, GM_Costco, ScorecardYear, Years, RatioYears). "
                         "Every cross-tab formula reads as a sentence and survives row insertions."),
        ("XLOOKUP + IFERROR", "All ratio cells: =IFERROR((XLOOKUP(year,Years,Company_Revenue)-XLOOKUP(year,Years,Company_COGS))/XLOOKUP(year,Years,Company_Revenue),\"n/a\"). "
                              "Missing data returns the text n/a rather than #DIV/0! or #VALUE!, and is styled gray-italic by conditional formatting."),
        ("Formula-built totals", "Total debt is a formula over the extracted components under a written policy (below), not a typed number; the Scorecard "
                                 "leader/laggard and takeaway text are built with INDEX/MATCH, RANK.EQ and TEXT so they re-write when the year input changes."),
        ("Scorecard year input", "Data-validated list (FY2016–FY2025) driving XLOOKUPs on the Scorecard and the at-a-glance figures; ranks and color scales recompute."),
        ("Conditional formatting", "Three-color scales per metric row (direction-aware: low debt-to-equity is green), text-detection rules for n/a, and a reversed scale on the composite rank."),
        ("Native combo charts", "Each Trend chart is an Excel column+line combination: a zero-gap column series on a hidden secondary axis shades FY2020–FY2023; "
                                "ratio lines sit on top. The helper block converts n/a to NA() so Excel draws gaps rather than zeros."),
        ("Model conventions", "Blue font = hard-coded SEC input; black = formula; yellow shading = instance/derived value; gray band = debt components; "
                              "freeze panes, print titles, fit-to-width and footers on every tab; document properties set."),
        ("Verification", f"{counts['formulas']} formulas written; the workbook is recalculated in Excel via COM after the build and scanned for error values "
                         "(#REF!, #DIV/0!, #NAME?, #VALUE!) – the only #N/A values are the intentional NA() chart gaps. Extraction is covered by 12 pytest checks "
                         "(coverage, accounting identities, derivation tie-outs, segment sanity)."),
        ("ACCOUNTING POLICY & COMPARABILITY", None),
        ("Revenue (ASC 606)", "Revenue = net sales per the income statement. Tag names changed at ASC 606 adoption (SalesRevenueNet / SalesRevenueGoodsNet → "
                              "RevenueFromContractWithCustomerExcludingAssessedTax), so tags are tried in a per-company priority list held in config/companies.yaml. "
                              "Costco reports 'Total revenue' = net sales + membership fees; this workbook uses NET SALES (the ProductMember disaggregated-revenue fact, SalesRevenueNet before 2018) so Costco's basis matches Walmart's and Sam's Club's net sales, which also exclude membership income. Membership fees ($5.3B FY2025) are a memo row: they carry no cost of sales, so on a total-revenue basis Costco's gross margin would read about two points higher."),
        ("Leases (ASC 842)", "Adopted FY2019 (Costco FY2020). Operating lease liabilities are EXCLUDED from Total debt – including them would create a step-up "
                             "in FY2019 with no financing event – and shown as a memo line and in the supplementary lease-inclusive debt-to-equity. Finance/capital "
                             "leases are INCLUDED throughout because they were liabilities under ASC 840 too. Policy per company in config/debt_policy.yaml."),
        ("Segments (ASC 280 / ASU 2023-07)", "Sam's Club figures are Walmart's reportable-segment disclosures (StatementBusinessSegmentsAxis = SamsClubUSMember, "
                                             "renamed from SamsClubMember in the FY2024 10-K). ASU 2023-07 required significant segment expenses from FY2024, which is "
                                             "why cost of revenue – and so gross margin – exists for FY2022–FY2025 only. Amazon's 'Physical stores' line is a "
                                             "disaggregated-revenue disclosure, not a segment, and includes more than Whole Foods."),
        ("Inventory (LIFO vs FIFO)", "Walmart U.S. (retail method, LIFO), Sam's Club (weighted-average, LIFO), Costco U.S. (LIFO), Kroger (~91% LIFO, link-chain "
                                     "dollar-value) and Whole Foods (LIFO) all report cost of sales after a LIFO charge in inflationary years; Costco Canada and "
                                     "Walmart International are FIFO. In 2021–2023 LIFO raised reported cost of sales and depressed gross margin and inventory "
                                     "for the LIFO filers – Kroger's reserve grew from $1.3B (FY2016) to $2.6B (FY2025) and its LIFO charge was $157M in FY2025; Costco's FY2022 LIFO charge cost 19 bp of gross margin (~$0.4B), FY2023–FY2024 were immaterial, FY2025 7 bp. Effects: (a) gross margins for FY2021–FY2023 "
                                     "are understated relative to a FIFO basis, more so for Kroger; (b) current assets and therefore the current ratio are "
                                     "understated for LIFO filers; (c) only Kroger discloses a reserve balance – Walmart stated that LIFO approximated FIFO through "
                                     "January 2022 (FY2021) and has given only sensitivity language since; Costco discloses the annual charge but no reserve balance. Ratios are presented as reported and trends, rather than levels, "
                                     "are emphasized in the findings."),
        ("Equity and net income", "Total equity includes noncontrolling interests (Walmart $6.3B FY2025); net income is the amount attributable to the parent. "
                                  "Debt-to-equity on a parent-only equity basis would be marginally higher for Walmart; no other company has material NCI."),
        ("Whole Foods cost line", "'Cost of goods sold and occupancy costs' includes store rent, so its gross margin is structurally lower than a pure "
                                  "merchandise margin and is not compared against peers on level – only its own two-year trend is used."),
        ("53-week years", "Costco and Kroger FY2023 had 53 weeks; growth in FY2023 is overstated and FY2024 understated by about 2 points. Walmart's January 31 "
                          "year-end is fixed (no 53-week years)."),
        ("ASSUMPTIONS & LIMITS", None),
        ("Assumptions", "Latest-filed values are the best available (restatements are accepted); segment operating income is comparable to consolidated "
                        "operating income for margin purposes; Amazon 'Physical stores' is used as a revenue proxy for Whole Foods and labeled as such."),
        ("Out of scope", "FIFO restatement of LIFO filers; adjustments for 53-week years; currency effects in Walmart International and Costco's non-U.S. "
                         "operations; pro-forma treatment of Kroger's terminated Albertsons merger."),
        ("TOOLS", None),
        ("Pipeline", "Python 3 · requests (EDGAR) · lxml (XBRL instance parsing) · pandas (tidy tables) · PyYAML (company register and debt policy) · "
                     "openpyxl (workbook generation) · pywin32 (Excel recalculation and error scan) · pytest · Git. Re-running src/pull_data.py, "
                     "src/pull_segments.py and src/build_workbook.py regenerates this file from EDGAR end-to-end."),
        ("Author", f"Ryan Barry – personal portfolio project, {TODAY.strftime('%B %Y')}. Public filings only; not an H-E-B work product."),
    ]
    r = 4
    for k, v in items:
        if v is None:
            header_band(ws, r, k, 1, 2, fill=FILL_H, font=F_H, height=20)
            r += 1
            continue
        put(ws, f"A{r}", k, font=F_B, align=Alignment(vertical="top", indent=1, wrap_text=True), border=BOX)
        put(ws, f"B{r}", v, font=F_N, align=WRAP, border=BOX)
        ws.row_dimensions[r].height = 15.5 * max(1, math.ceil(len(v) / 165))
        r += 1
    return ws


# ------------------------------------------------------------------ Data Lineage
def build_lineage(wb, D):
    ws = wb.create_sheet("Data Lineage")
    sheet_setup(ws, MID)
    title_band(ws, 8, "Data Lineage", "Every input value with its SEC citation – one row per value on Raw Data. 'Restated' = the latest-filed value differs from the "
                                      "first-filed value; both are shown. Links open the filing index on sec.gov.", 17)
    cols = ["Company", "Line item", "Fiscal year", "Period start", "Period end", "Value (USD)", "XBRL tag / basis", "Dimension", "Accession no.", "Filed",
            "Status", "Restated", "First-filed value", "First filed", "Versions", "Note", "EDGAR link"]
    widths = [12, 30, 10, 12, 12, 18, 52, 40, 22, 11, 10, 9, 18, 11, 8, 70, 40]
    for i, (c, w) in enumerate(zip(cols, widths)):
        put(ws, f"{L(1 + i)}4", c, font=F_H, fill=FILL_H, align=CENTER)
        ws.column_dimensions[L(1 + i)].width = w
    ws.freeze_panes = "C5"
    labels = {k: lab for lab, k, _ in INPUT_ROWS} | {k: lab for lab, k in DEBT_ROWS} | {k: lab for lab, k, _ in MEMO_ROWS}
    labels |= {"net_sales": "Revenue (net sales) - Costco", "total_revenue": "Total revenue incl. membership fees (reference)", "operating_lease_current": "Operating lease liability – current", "operating_lease_noncurrent": "Operating lease liability – non-current",
               "operating_lease_total": "Operating lease liability – total", "equity_parent_only": "Stockholders' equity – parent only (reference)"}
    facts = D["facts"][D["facts"].status.isin(["ok", "instance", "derived"])].copy()
    facts.loc[(facts.ticker == "COST") & (facts.metric == "revenue"), "metric"] = "total_revenue"
    facts["dimension"] = ""
    segs = D["segs"].copy()
    segs["dimension"] = segs.axis + " = " + segs.member
    segs["note"] = "Dimensional fact parsed from the 10-K XBRL instance (" + segs.instance + ")"
    segs["ticker"] = segs.ticker.replace({"WFM": "WFM"})
    segs.loc[segs.ticker == "WFM", "ticker"] = "AMZN→WFM"
    allrows = pd.concat([facts, segs], ignore_index=True).sort_values(["ticker", "metric", "fiscal_year"])
    r = 5
    for _, f in allrows.iterrows():
        cik = int(f.cik)
        acc = str(f.accession)
        link = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/"
        vals = [DISPLAY.get(f.ticker, f.ticker), labels.get(f.metric, f.metric), int(f.fiscal_year), f.get("period_start", "") or "", f.period_end, float(f.value),
                f.tag, f.dimension, acc, f.filed, f.status, "yes" if bool(f.get("restated", False)) else "", f.get("first_value", ""), f.get("first_filed", ""),
                int(f.n_versions) if pd.notna(f.get("n_versions")) else "", f.get("note", "") if pd.notna(f.get("note", "")) else "", link]
        for i, v in enumerate(vals):
            c = put(ws, f"{L(1 + i)}{r}", v, font=F_NOTE)
            if i == 5:
                c.number_format = "#,##0"
            if i == 12 and isinstance(v, float) and not math.isnan(v):
                c.number_format = "#,##0"
            if i == 16:
                c.hyperlink = link
                c.font = Font(name="Calibri", size=9, color=TEAL, underline="single")
        if f.status in ("instance", "derived"):
            for i in range(len(vals)):
                ws.cell(row=r, column=1 + i).fill = PatternFill("solid", fgColor="FFF2CC")
        r += 1
    ws.auto_filter.ref = f"A4:{L(len(cols))}{r - 1}"
    return ws, r - 5


# ------------------------------------------------------------------ main
def main() -> None:
    global n_formulas
    D = load()
    wb = Workbook()
    build_cover(wb)
    RAW = build_raw(wb, D)
    R = build_ratios(wb)
    S = build_scorecard(wb, R)
    build_trend(wb, R)
    build_summary(wb, RAW, S)
    lineage, n_lineage = build_lineage(wb, D)
    counts = {"names": len(wb.defined_names), "formulas": n_formulas}
    build_method(wb, counts)
    order = ["Cover", "Executive Summary", "Raw Data", "Ratio Calculations", "Scorecard", "Trend", "Methodology & Tools", "Data Lineage"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for ws in wb.worksheets:
        if ws.title in ("Raw Data", "Ratio Calculations", "Data Lineage"):
            ws.print_title_rows = "4:4"
    wb.properties.title = "Retail Peer Financial Benchmarking: Walmart, Costco, Kroger & Sam's Club"
    wb.properties.creator = "Ryan Barry"
    wb.properties.subject = "Ten-year financial comparison from SEC 10-K filings"
    wb.properties.keywords = "SEC EDGAR, XBRL, financial ratios, competitor analysis, ASC 842, ASC 280, LIFO"
    wb.properties.description = "Built programmatically from SEC EDGAR; all analytical cells are live formulas."
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"wrote {OUT}: {len(wb.worksheets)} tabs, {counts['formulas']} formulas, {counts['names']} named ranges, {n_lineage} lineage rows")


if __name__ == "__main__":
    main()
