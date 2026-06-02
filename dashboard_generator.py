#!/usr/bin/env python3
# =============================================================================
#  dashboard_generator.py — CRM Sahibabad Narrow Complex
#  Usage:  python dashboard_generator.py [--day 29]
#  Reads:  input/MILL MIS.xlsx  input/Annealing and 2HI SPM MIS.xlsx
#          input/CRS MIS.xlsx   input/Target.xlsx
#  Writes: output/Daily_Dashboard.png  .pdf  .pptx  dashboard_data.xlsx
# =============================================================================

import sys, os, datetime, json, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import (MILL_FILE, ANNEALING_FILE, CRS_FILE, TARGET_FILE,
                    OUTPUT_PNG, OUTPUT_PDF, OUTPUT_PPT, OUTPUT_XLSX,
                    OUTPUT_DIR, INPUT_DIR)

def _safe_load(loader, path, label, *args, **kwargs):
    if not os.path.exists(path):
        print(f"  [SKIP] {label} not found at {path}")
        return {}
    try:
        result = loader(path, *args, **kwargs)
        print(f"  [OK]   {label}")
        return result
    except Exception as e:
        print(f"  [ERR]  {label}: {e}")
        return {}

def build_report_data(day=None):
    """Load all 4 MIS files and return unified report_data dict."""
    from extractors.targets   import parse_targets
    from extractors.crs       import parse_crs
    from generate_dashboard   import xlsx_to_tsv, parse_mis, parse_anneal_mis

    print("Reading input files...")
    targets = _safe_load(parse_targets, TARGET_FILE, "Target.xlsx") or {}

    # Rolling
    rolling = {}
    if os.path.exists(MILL_FILE):
        try:
            tsv, _ = xlsx_to_tsv(MILL_FILE)
            rolling = parse_mis(tsv, report_day=day)
            for mn, m in rolling.items():
                t = targets.get("rolling_day")
                if t: m["day_target"] = t / max(len(rolling), 1)
            print(f"  [OK]   MILL MIS.xlsx  ({', '.join(rolling.keys())})")
        except Exception as e:
            print(f"  [ERR]  MILL MIS.xlsx: {e}")
    else:
        print(f"  [SKIP] MILL MIS.xlsx not found")

    # Annealing + 2HI
    annealing = {}
    if os.path.exists(ANNEALING_FILE):
        try:
            tsv, _ = xlsx_to_tsv(ANNEALING_FILE)
            annealing = parse_anneal_mis(tsv, report_day=day)
            print(f"  [OK]   Annealing MIS.xlsx")
        except Exception as e:
            print(f"  [ERR]  Annealing MIS.xlsx: {e}")
    else:
        print(f"  [SKIP] Annealing MIS.xlsx not found")

    # CRS
    crs = _safe_load(parse_crs, CRS_FILE, "CRS MIS.xlsx") or {}

    data = {
        "date":      crs.get("report_date") or datetime.date.today().strftime("%d.%m.%Y"),
        "rolling":   rolling,
        "annealing": annealing,
        "crs":       crs,
    }
    return data, targets

def render_png(html_path, png_path):
    """Playwright → PNG at 1920×1080."""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(viewport={"width": 1920, "height": 1080})
            pg.goto("file://" + os.path.abspath(html_path))
            pg.wait_for_load_state("networkidle")
            pg.screenshot(path=png_path, full_page=False,
                          clip={"x":0,"y":0,"width":1920,"height":1080})
            b.close()
        print(f"  [PNG]  {png_path}")
        return png_path
    except Exception as e:
        print(f"  [PNG]  FAILED: {e}")
        print("         Install: pip install playwright && python -m playwright install chromium")
        return None

def render_pdf(html_path, pdf_path):
    """WeasyPrint → PDF."""
    try:
        import weasyprint
        weasyprint.HTML(filename=html_path).write_pdf(pdf_path)
        print(f"  [PDF]  {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"  [PDF]  FAILED: {e}  (pip install weasyprint)")
        return None

def render_pptx(png_path, pptx_path):
    """Embed PNG into a 16:9 branded PPTX slide."""
    if not png_path or not os.path.exists(png_path):
        print("  [PPT]  Skipped (no PNG)")
        return None
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        prs = Presentation()
        prs.slide_width  = Inches(16)
        prs.slide_height = Inches(9)
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        sl.shapes.add_picture(png_path, Inches(0), Inches(0), Inches(16), Inches(9))
        prs.save(pptx_path)
        print(f"  [PPT]  {pptx_path}")
        return pptx_path
    except Exception as e:
        print(f"  [PPT]  FAILED: {e}")
        return None

def export_xlsx(data, targets, path):
    """Save extracted KPIs to an xlsx for audit/history."""
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active; ws.title = "Dashboard Data"
        rows = [
            ["Section","KPI","Value","Target","Unit"],
        ]
        for mn, m in data.get("rolling",{}).items():
            rows += [
                [mn, "Day Total",  m.get("day_total",0),  m.get("day_target",""), "MT"],
                [mn, "Day Roll",   m.get("day_roll",0),   "", "MT"],
                [mn, "Day R/R",    m.get("day_rr",0),     "", "MT"],
                [mn, "Yield",      m.get("yield",0),      targets.get(f"{mn.lower()}_yield_mtd",""), "%"],
                [mn, "Util",       m.get("day_util",0),   targets.get(f"{mn.lower()}_util_mtd",""), "%"],
                [mn, "Cumm",       m.get("cumm_out",0),   "", "MT"],
            ]
        ann = data.get("annealing",{}).get("ann02",{})
        rows += [["Annealing","Day Prod",ann.get("day_prod",0),targets.get("ann_day",""),"MT"],
                 ["Annealing","Charges", ann.get("charges",0),"",""]]
        spm = data.get("annealing",{}).get("skin_pass",{})
        rows += [["2HI SKP","Day Prod",spm.get("day_prod",0),targets.get("spm_day",""),"MT"]]
        crs = data.get("crs",{})
        gr  = crs.get("gr",{})
        rows += [
            ["CRS GR","Tube",  gr.get("tube",{}).get("yesterday",0), targets.get("tube_gr_day",""),"T"],
            ["CRS GR","OEM",   gr.get("oem",{}).get("yesterday",0),  targets.get("oem_gr_day",""), "T"],
            ["CRS GR","Total", gr.get("total",{}).get("yesterday",0),targets.get("total_gr_day",""),"T"],
            ["CRS",   "Hold",  crs.get("hold",{}).get("value",0),"","MT"],
            ["CRS",   "SKP WIP",crs.get("skp_wip",0),"","MT"],
            ["CRS",   "Total at CRS",crs.get("total_at_crs",0),"","MT"],
        ]
        for r in rows:
            ws.append(r)
        # Style header
        from openpyxl.styles import Font, PatternFill
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="005CB9")
        wb.save(path)
        print(f"  [XLSX] {path}")
    except Exception as e:
        print(f"  [XLSX] FAILED: {e}")

def main():
    parser = argparse.ArgumentParser(description="CRM Dashboard Generator")
    parser.add_argument("--day", type=int, default=None,
                        help="Day of month to extract (default: latest filled row)")
    args = parser.parse_args()

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(INPUT_DIR,  exist_ok=True)

    print(f"\n{'='*55}")
    print(f"  CRM SAHIBABAD — Dashboard Generator v2.0")
    print(f"  {datetime.datetime.now():%d-%m-%Y %H:%M}")
    print(f"{'='*55}")

    data, targets = build_report_data(day=args.day)

    from dashboard.alerts import generate_alerts
    from dashboard.notes  import generate_notes
    from dashboard.layout import build_html

    alerts = generate_alerts(data, targets)
    notes  = generate_notes(data, targets, alerts)

    print(f"\nAlerts:  {len(alerts)}  "
          f"({sum(1 for a in alerts if a['priority']=='critical')} critical)")
    print(f"Notes:   {len(notes)}")

    html = build_html(data, targets, alerts, notes)

    html_path = os.path.join(OUTPUT_DIR, "Daily_Dashboard.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n  [HTML] {html_path}")

    png_path  = render_png(html_path, OUTPUT_PNG)
    pdf_path  = render_pdf(html_path, OUTPUT_PDF)
    pptx_path = render_pptx(png_path, OUTPUT_PPT)
    export_xlsx(data, targets, OUTPUT_XLSX)

    print(f"\n{'='*55}")
    print(f"  Done.  Outputs in: {OUTPUT_DIR}/")
    print(f"{'='*55}\n")

if __name__ == "__main__":
    main()
