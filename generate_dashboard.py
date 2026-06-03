# generate_dashboard.py — ALL MODULES MERGED FOR SINGLE-FILE DEPLOY

# Do not split this file. Upload as-is to GitHub.


# =============================================================================
#  config.py — CRM Sahibabad Narrow Complex Dashboard
#  All colours, RAG thresholds, file paths, section titles in one place.
# =============================================================================

import os

# ── Tata Steel brand colours (from PDF spec v2.0) ────────────────────────────
COLORS = {
    "primary":   "#005CB9",   # Tata blue
    "secondary": "#0087E0",   # lighter blue
    "bg":        "#F4F7FB",   # page background
    "card":      "#FFFFFF",   # card background
    "success":   "#00A651",   # green ≥ 100 %
    "warning":   "#F5A623",   # amber 95–99 %
    "critical":  "#D0021B",   # red < 95 %
    "muted":     "#6B7280",
    "border":    "#E2E8F0",
    "navy":      "#0F2A5C",
    "text":      "#1E293B",
}

# ── RAG logic ─────────────────────────────────────────────────────────────────
def rag_color(actual, target):
    """Return (hex_color, label, pct) tuple. pct=None when target is missing."""
    if not target or target in (0, "NA", "-", None):
        return COLORS["muted"], "—", None
    pct = round(100 * actual / target, 1)
    if pct >= 100:
        return COLORS["success"], "✓", pct
    elif pct >= 95:
        return COLORS["warning"], "~", pct
    else:
        return COLORS["critical"], "✗", pct

# ── File paths ────────────────────────────────────────────────────────────────
BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

MILL_FILE      = os.path.join(INPUT_DIR, "MILL MIS.xlsx")
ANNEALING_FILE = os.path.join(INPUT_DIR, "Annealing and 2HI SPM MIS.xlsx")
CRS_FILE       = os.path.join(INPUT_DIR, "CRS MIS.xlsx")
TARGET_FILE    = os.path.join(INPUT_DIR, "Target.xlsx")

OUTPUT_PNG  = os.path.join(OUTPUT_DIR, "Daily_Dashboard.png")
OUTPUT_PDF  = os.path.join(OUTPUT_DIR, "Daily_Dashboard.pdf")
OUTPUT_PPT  = os.path.join(OUTPUT_DIR, "Daily_Dashboard.pptx")
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "dashboard_data.xlsx")

# ── Dashboard resolution ──────────────────────────────────────────────────────
WIDTH  = 1920
HEIGHT = 1080

# ── CRS MIS column mapping ────────────────────────────────────────────────────
# Row labels (col 1) map to semantic keys; values are in col 4; MTD/cumm in col 6
CRS_ROWS = {
    # label in col1     : (key, sub_col2_filter or None)
    "FOR SLITTING":  "slitting",    # sub-rows: Oem, Tube, STRAPING, Total  (col 2)
    "FOR PACKING":   "packing",     # sub-rows: Oem, Tube, Total
    "NO PLAN":       "no_plan",
    "HROP S.PASS":   "hrop_skp",
    "HOLD":          "hold",
    "S.PASS WIP":    "skp_wip",
    "TOTAL AT CRS":  "total_at_crs",
    "G.R STATUS":    "gr",          # sub-rows: TUBE/…, OEM/…, TOTAL
    "CRCA SLITTING": "crca_slitting",
}

# ── Target.xlsx row-key mapping ───────────────────────────────────────────────
TARGET_ROWS = {
    "Rolling Total":        ("rolling_day",  "rolling_mtd"),
    "CRM04 Utilisation":    ("crm04_util_day", "crm04_util_mtd"),
    "CRM04 Yield":          ("crm04_yield_day", "crm04_yield_mtd"),
    "CRM06 Utilisation":    ("crm06_util_day", "crm06_util_mtd"),
    "CRM06 Yield":          ("crm06_yield_day", "crm06_yield_mtd"),
    "Annealing Production": ("ann_day",  "ann_mtd"),
    "2HI Production":       ("spm_day",  "spm_mtd"),
    "Tube GR":              ("tube_gr_day", "tube_gr_mtd"),
    "OEM GR":               ("oem_gr_day",  "oem_gr_mtd"),
    "Total GR":             ("total_gr_day", "total_gr_mtd"),
    "Hold Material Max":    ("hold_max",  None),
    "Skinpass WIP Max":     ("skp_wip_max", None),
}


# ===========================================================================
# CONFIG END / CORE PARSERS START

# ===========================================================================

#!/usr/bin/env python3
# =============================================================================
#  Narrow Complex C-Shift Dashboard Generator  (Tata Steel CRM Sahibabad)
# =============================================================================
#  Ingests messy daily shift-production data (pasted WhatsApp-style text and/or
#  OCR'd images) and renders a Tata-branded dashboard in HTML, PNG and PPTX.
#
#  INSTALL (Python):
#     pip install python-pptx Pillow jinja2
#  INSTALL (OCR, only if using images):
#     pip install pytesseract pdf2image
#     # system:  sudo apt-get install tesseract-ocr poppler-utils
#  INSTALL (rendering, for the PNG — gives COLOR emojis like the reference):
#     pip install playwright
#     python -m playwright install chromium
#     # system (Linux): sudo apt-get install fonts-noto-color-emoji
#     # Fallback if no Playwright: sudo apt-get install wkhtmltopdf
#     #   (wkhtmltoimage works but renders emojis monochrome)
#
#  RUN:
#     python generate_dashboard.py            # interactive: paste text / give image paths
#     python generate_dashboard.py --demo     # render with built-in sample data
#     python generate_dashboard.py --text report.txt          # parse a text file
#     python generate_dashboard.py --image board.jpg slit.png # OCR images (review step)
#
#  OUTPUT (written to ./output/):
#     dashboard_<date>.html   dashboard_<date>.png   dashboard_<date>.pptx
# =============================================================================

import os, re, sys, json, subprocess, datetime, shutil

# ----------------------------------------------------------------------------- 
# CONFIG  — edit freely
# -----------------------------------------------------------------------------
CONFIG = {
    "navy":   "#143B6B",
    "orange": "#F47920",
    "green":  "#1E8E3E",
    "red":    "#C5221F",
    "blue":   "#1A56A8",
    "grey":   "#5F6368",
    "card_bg": "#FFFFFF",
    "page_bg": "#EEF1F5",
    "logo_path": "",            # absolute path to a Tata logo PNG (optional)
    "gr_target": 7000,          # MTD GR target (T) for the headline gauge
    "rolling_targets": {        # daily ROLL+RR+SKP target per mill (edit these)
        "CRM04": 150, "CRM06": 260, "CRM07": 260,
    },
    "gr_headline": {"gr_today": 243, "gr_mtd": 6547, "gr_target": 7000},
    "output_dir": "output",
    "plant": "Tata Steel CRM Sahibabad — Narrow Complex",
}

# =============================================================================
#  report_data SCHEMA  (this dict is the single contract between parse + render)
# =============================================================================
#  {
#    "date": "29-05-2026",
#    "shift": "C",
#    "rolling": {
#        "CRM04": {"shift_total":, "roll":, "rr":, "skp":, "rr_note":,
#                  "delay": {"operational":"30 Mins","mechanical":"25 Min"},
#                  "day_total":, "day_roll":, "day_rr":,
#                  "cumm_total":, "cumm_roll":, "cumm_rr":, "cumm_skp":},
#        "CRM06": {...same...},
#        "REWINDING06": {"shift_total":, "cumm_total":}
#    },
#    "annealing": {"ann02": {field:value,...}, "rwl02": {...}, "skin_pass": {...}},
#    "slitting":  {"lines":[{name,shift,day,cumm,remark},...], "total":{...},
#                  "packing_gr":{...}, "dispatch":{...}, "permits":[...], "notes":[...]},
#    "headline":  {"gr_today":, "gr_mtd":, "gr_target":}      # optional
#  }
# =============================================================================


def n(v):
    """Normalize a numeric-ish token to float or 0.0. 'Nil','00','' -> 0."""
    if v is None:
        return 0.0
    s = str(v).strip().lower()
    s = s.replace("*", "").replace("mt", "").replace(",", "").strip()
    if s in ("", "nil", "00", "0", "n/a", "na", "-"):
        return 0.0
    m = re.search(r"-?\d+\.?\d*", s)
    return float(m.group()) if m else 0.0


def grab(text, pattern, group=1, default=""):
    m = re.search(pattern, text, re.IGNORECASE)
    return m.group(group).strip() if m else default


# -----------------------------------------------------------------------------
#  PARSER 1 — Rolling text (CRM04 / CRM06 / Rewinding 06)
# -----------------------------------------------------------------------------
def parse_rolling_text(text):
    """Tolerant parser for the WhatsApp-style rolling report."""
    text = text.replace("*", "")
    out = {}

    def block(label, stop_labels):
        # capture text between `label` and the next section header
        stop = "|".join(re.escape(s) for s in stop_labels)
        m = re.search(rf"{re.escape(label)}(.*?)(?={stop}|$)", text, re.IGNORECASE | re.DOTALL)
        return m.group(1) if m else ""

    def parse_mill(seg):
        d = {}
        d["shift_total"] = n(grab(seg, r"Total\s*-?\s*([\d.]+)"))
        d["roll"]  = n(grab(seg, r"ROLL\s*-?\s*([\d.]+)"))
        d["rr"]    = n(grab(seg, r"R/R\s*-?\s*([\d.]+)"))
        d["skp"]   = n(grab(seg, r"SKP\s*-?\s*([\d.]+)"))
        d["rr_note"] = grab(seg, r"R/R[^\n]*\(([^)]+)\)")
        d["delay_op"]  = grab(seg, r"Operational\s*[-:]?\s*([\w. ]+?)(?:\n|Mechanical|$)")
        d["delay_mech"]= grab(seg, r"Mechanical\s*[-:]?\s*([\w. ]+?)(?:\n|$)")
        day = re.search(r"Total Day [Pp]roduction\s*:?\s*([\d.]+)(.*?)(?=CUMM|$)", seg, re.DOTALL | re.IGNORECASE)
        if day:
            d["day_total"] = n(day.group(1))
            d["day_roll"]  = n(grab(day.group(2), r"ROLL\s*-?\s*([\d.]+)"))
            d["day_rr"]    = n(grab(day.group(2), r"R/R\s*-?\s*([\d.]+)"))
        cum = re.search(r"CUMM[ULATIVE.]*\s*-?\s*([\d.]+)(.*?)$", seg, re.DOTALL | re.IGNORECASE)
        if cum:
            d["cumm_total"] = n(cum.group(1))
            d["cumm_roll"]  = n(grab(cum.group(2), r"ROLL\s*-?\s*([\d.]+)"))
            d["cumm_rr"]    = n(grab(cum.group(2), r"R/R\s*-?\s*([\d.]+)"))
            d["cumm_skp"]   = n(grab(cum.group(2), r"SKP\s*-?\s*([\d.]+)"))
        return d

    crm04_seg = block("CRM04", ["CRM06"])
    crm06_seg = block("CRM06", ["Rewinding"])
    rew_seg   = block("Rewinding 06", ["$$$"])

    if crm04_seg.strip():
        out["CRM04"] = parse_mill(crm04_seg)
    if crm06_seg.strip():
        out["CRM06"] = parse_mill(crm06_seg)
    if rew_seg.strip():
        out["REWINDING06"] = {
            "shift_total": n(grab(rew_seg, r"Total\s*-?\s*([\d.]+)")),
            "cumm_total":  n(grab(rew_seg, r"CUMM[ULATIVE.]*\s*-?\s*([\d.]+)")),
        }

    date = grab(text, r"(\d{2}[-/]\d{2}[-/]\d{4})")
    return out, date


# -----------------------------------------------------------------------------
#  PARSER 3 — MIS sheet (tab-separated). Two stacked mill tables, one row/day.
#  Returns a rolling dict keyed by the mill label found in each summary block
#  (e.g. CRM06, CRM07). Picks the row whose DATE column == report_day.
# -----------------------------------------------------------------------------
# Column indices within each mill's data row (verified against the MIS layout):
MIS_COL = {
    "date": 0, "roll": 2, "rr": 4, "skp": 6, "in_gauge": 7, "tot_width": 8,
    "out_gauge": 10, "coils": 11, "delay": 12, "avg_gauge": 17, "avg_width": 18,
    "yield": 22, "day_util": 23, "cumm_roll": 26, "cumm_out": 27,
    "cumm_yield": 28, "util_tilldate": 29,
}

def parse_mis(text, report_day=None):
    """Parse the MIS tab-separated dump. report_day is an int (day of month)."""
    lines = text.split("\n")
    rolling = {}
    current_rows = []          # accumulates data rows for the current table
    pending_label = None       # mill label discovered in a summary block

    def flush(label, rows):
        if not label or not rows:
            return
        nonempty = [r for r in rows if len(r) > 6 and any(c.strip() for c in r[1:7])]
        chosen = None
        # match by calendar day if the table uses 1..31 numbering
        if report_day is not None:
            for r in rows:
                d = r[0].strip()
                if d.isdigit() and int(d) == report_day and len(r) > 6 and any(c.strip() for c in r[1:7]):
                    chosen = r
                    break
        # fallback: the last non-empty data row (latest filled day)
        if chosen is None and nonempty:
            chosen = nonempty[-1]
        if chosen is None:
            return
        g = lambda k: chosen[MIS_COL[k]] if MIS_COL[k] < len(chosen) else ""
        roll, rr, skp = n(g("roll")), n(g("rr")), n(g("skp"))
        rolling[label] = {
            "day_total": round(roll + rr + skp, 3),
            "day_roll": roll, "day_rr": rr, "day_skp": skp,
            "coils": n(g("coils")), "delay_hrs": n(g("delay")),
            "avg_gauge": n(g("avg_gauge")), "avg_width": n(g("avg_width")),
            "yield": n(g("yield")), "day_util": n(g("day_util")),
            "cumm_roll": n(g("cumm_roll")), "cumm_out": n(g("cumm_out")),
            "cumm_yield": n(g("cumm_yield")), "util_tilldate": n(g("util_tilldate")),
            "row_date": g("date"),
        }

    for ln in lines:
        cells = ln.split("\t")
        first = cells[0].strip()
        # a data row starts with a pure integer (calendar day OR continuous index)
        if first.isdigit():
            current_rows.append(cells)
            continue
        # summary block reveals the mill label, e.g. a line containing 'CRM06'
        m = re.search(r"\b(CRM\d{2})\b", ln)
        if m:
            pending_label = m.group(1)
            flush(pending_label, current_rows)
            current_rows = []
    # flush any trailing table
    if current_rows and pending_label:
        flush(pending_label, current_rows)
    return rolling


# -----------------------------------------------------------------------------
#  PARSER 3b — Annealing + 2HI/SPM MIS sheet ("PRODUCTION & GAS DETAILS").
#  One wide table, one row per date (col A = M/D/YYYY). Joins A-F (annealing
#  production + charges) with N-T (gas, water, 2HI). Matches by calendar date.
# -----------------------------------------------------------------------------
# 0-based column indices in the wide sheet:
ANN_COL = {
    "date": 0, "prod_old": 1, "prod_new": 2, "chg_old": 3, "chg_new": 4, "total": 5,
    "lng_nm3": 13, "lng_m3mt": 14, "lng_nm3mt": 15,
    "water": 16, "hi_prod": 17, "id_change": 18, "hrop": 19,
    "rpp": 20, "mes": 21, "carol_drum": 22,
}

def _match_date(cell, report_day):
    """cell like '5/18/2026' -> day int; return True if matches report_day."""
    m = re.match(r"\s*\d{1,2}[/-](\d{1,2})[/-]\d{2,4}", cell)
    if not m:
        return False
    return report_day is None or int(m.group(1)) == report_day

def parse_anneal_mis(text, report_day=None):
    """Parse the Annealing + 2HI/SPM sheet; pick the row for report_day."""
    rows = []
    for ln in text.split("\n"):
        cells = ln.split("\t")
        if re.match(r"\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", cells[0]):
            rows.append(cells)
    if not rows:
        return {}
    chosen = None
    if report_day is not None:
        for r in rows:
            if _match_date(r[0], report_day):
                chosen = r; break
    if chosen is None:
        chosen = rows[-1]      # latest filled date
    g = lambda k: chosen[ANN_COL[k]] if ANN_COL[k] < len(chosen) else ""
    prod_old, prod_new = n(g("prod_old")), n(g("prod_new"))
    total = n(g("total")) or round(prod_old + prod_new, 3)
    out = {
        "ann02": {
            "day_prod": total, "prod_old": prod_old, "prod_new": prod_new,
            "chg_old": n(g("chg_old")), "chg_new": n(g("chg_new")),
            "charges": int(n(g("chg_old")) + n(g("chg_new"))),
            "water": n(g("water")), "lng_nm3": n(g("lng_nm3")),
            "lng_m3mt": n(g("lng_m3mt")), "row_date": g("date").strip(),
        },
        "skin_pass": {     # 2HI / SPM block
            "hi_prod": n(g("hi_prod")), "id_change": n(g("id_change")),
            "hrop": n(g("hrop")),
            "day_prod": round(n(g("hi_prod")) + n(g("id_change")) + n(g("hrop")), 3),
            "carol_drum": g("carol_drum").strip(),
        },
    }
    return out


# -----------------------------------------------------------------------------
#  Excel / SharePoint loader
#  Turn an .xlsx (local path, URL, or bytes) + sheet name into the tab-separated
#  text that parse_mis / parse_anneal_mis already understand. Lets the user
#  point at a shared SharePoint file instead of pasting.
# -----------------------------------------------------------------------------
def _sharepoint_download_url(url):
    """Convert a SharePoint/OneDrive 'view' share link into a direct-download URL.
    Works only for links shared as 'Anyone with the link'."""
    if "download=1" in url or url.lower().endswith(".xlsx"):
        return url
    sep = "&" if "?" in url else "?"
    return url + sep + "download=1"

def xlsx_to_tsv(source, sheet_name=None):
    """source: local path, http(s) URL, or raw bytes. Returns (tsv_text, sheet_names).
    If sheet_name is None, uses the first sheet."""
    import io
    import openpyxl
    data = None
    if isinstance(source, (bytes, bytearray)):
        data = bytes(source)
    elif isinstance(source, str) and source.lower().startswith("http"):
        import requests
        u = _sharepoint_download_url(source)
        r = requests.get(u, allow_redirects=True, timeout=30)
        r.raise_for_status()
        ct = r.headers.get("content-type", "")
        if "html" in ct.lower():
            raise ValueError("The link returned a web page, not an Excel file. "
                             "Make sure it is shared as 'Anyone with the link' and "
                             "is a direct file link.")
        data = r.content
    else:
        with open(source, "rb") as f:
            data = f.read()
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    names = wb.sheetnames
    ws = wb[sheet_name] if (sheet_name and sheet_name in names) else wb[names[0]]
    lines = []
    for row in ws.iter_rows(values_only=True):
        cells = ["" if v is None else str(v) for v in row]
        lines.append("\t".join(cells))
    return "\n".join(lines), names




# -----------------------------------------------------------------------------
#  OCR (optional)
# -----------------------------------------------------------------------------
def ocr_extract(image_path):
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        print("  [OCR] pytesseract/Pillow not installed — skipping", image_path)
        return ""
    try:
        txt = pytesseract.image_to_string(Image.open(image_path))
        print(f"  [OCR] extracted {len(txt)} chars from {os.path.basename(image_path)}")
        return txt
    except Exception as e:
        print(f"  [OCR] failed on {image_path}: {e}")
        return ""


def review_loop(label, value):
    """Let the user confirm or correct an OCR'd value."""
    resp = input(f"    {label} = '{value}'  [Enter=keep / type correction]: ").strip()
    return resp if resp else value



# -----------------------------------------------------------------------------
#  RENDER — HTML  (fixed-size landscape INFOGRAPHIC, matches Tata Colors style)
# -----------------------------------------------------------------------------
# Professional line-style SVG icons (monochrome, tinted to accent color).
# Industrial dashboard aesthetic — no toy emojis in the KPI tiles.
_SVG = {
    "production": '<path d="M3 21V8l6 4V8l6 4V5l6 3v13z"/>',                      # factory/output
    "roll": '<ellipse cx="6" cy="12" rx="2.5" ry="8"/><path d="M6 4h11v16H6" fill="none" stroke-width="1.6"/><ellipse cx="17" cy="12" rx="2.5" ry="8"/>',  # coil
    "rr": '<path d="M4 12a8 8 0 0114-5M20 12a8 8 0 01-14 5" fill="none" stroke-width="1.8"/><path d="M18 5v3h-3M6 19v-3h3" fill="none" stroke-width="1.8"/>',  # re-roll arrows
    "skp": '<rect x="3" y="9" width="18" height="6" rx="1"/><path d="M3 12h18" stroke="#fff" stroke-width="1.2"/>',  # skin-pass roll bite
    "target": '<circle cx="12" cy="12" r="8.5" fill="none" stroke-width="1.8"/><circle cx="12" cy="12" r="4.5" fill="none" stroke-width="1.8"/><circle cx="12" cy="12" r="1.4"/>',
    "yield": '<path d="M4 13l5 5L20 6" fill="none" stroke-width="2.4"/>',          # check
    "util": '<circle cx="12" cy="12" r="8.5" fill="none" stroke-width="1.8"/><path d="M12 12V6M12 12l4 2.5" fill="none" stroke-width="1.8"/>',  # clock/util
    "gauge": '<path d="M4 18a8 8 0 1116 0" fill="none" stroke-width="1.8"/><path d="M12 18l4-5" fill="none" stroke-width="1.8"/><circle cx="12" cy="18" r="1.4"/>',  # gauge dial
    "width": '<path d="M3 12h18M3 12l3-3M3 12l3 3M21 12l-3-3M21 12l-3 3" fill="none" stroke-width="1.8"/>',  # measure
    "coils": '<circle cx="8" cy="9" r="3.2" fill="none" stroke-width="1.6"/><circle cx="15" cy="9" r="3.2" fill="none" stroke-width="1.6"/><circle cx="11.5" cy="15" r="3.2" fill="none" stroke-width="1.6"/>',  # stacked coils
    "cumm": '<path d="M3 17l5-5 4 3 6-7" fill="none" stroke-width="2"/><path d="M16 8h3v3" fill="none" stroke-width="2"/>',  # trend up
    "delay": '<path d="M12 3l9.5 16.5H2.5z" fill="none" stroke-width="1.8" stroke-linejoin="round"/><path d="M12 9v5" stroke-width="1.8"/><circle cx="12" cy="17" r="0.6"/>',  # warning triangle
    "factory": '<path d="M3 21V8l6 4V8l6 4V5l6 3v13z"/>',
}

_ALIAS = {
    "truck": "production", "gear": "rr", "chart": "cumm", "coil": "roll",
    "box": "skp", "stack": "production", "shield": "yield", "check": "yield",
    "warn": "delay", "pin": "gauge", "bar": "width", "up": "cumm",
    "badge": "yield", "note": "yield", "sum": "cumm", "fire": "delay",
}

def _svg(name, color, size):
    name = _ALIAS.get(name, name)
    body = _SVG.get(name, '<circle cx="12" cy="12" r="3"/>')
    return (f'<svg class="svgic" width="{size}" height="{size}" viewBox="0 0 24 24" '
            f'fill="{color}" stroke="{color}" stroke-linecap="round">{body}</svg>')

# Header/footer accents may still use a couple of restrained unicode marks.
EMOJI = {"factory": "", "calendar": "", "check": ""}

def render_html(data, path):
    """Build the dashboard HTML and write it to `path`. Returns path."""
    html = build_html(data)
    with open(path, "w") as f:
        f.write(html)
    return path

def build_html(data):
    c = CONFIG
    d = data

    def ic(name, color=None, size=20):
        return _svg(name, color or CONFIG["navy"], size)

    def fmt(v):
        if isinstance(v, (int, float)):
            return f"{v:g}"
        return v if v not in (None, "") else "—"

    def rag(actual, target):
        """Return color by achievement %: >=100 green, >=85 amber, else red."""
        if not target:
            return c["grey"], None
        p = round(100 * actual / target, 1)
        col = c["green"] if p >= 100 else (c["orange"] if p >= 85 else c["red"])
        return col, p

    # ---- primary stat card: icon badge + value + caption (top KPI row)
    def stat(icon, label, value, color, unit="MT"):
        disp = fmt(value)
        u = f'<span class="s-unit">{unit}</span>' if (unit and isinstance(value,(int,float))) else ""
        return f"""<div class="stat">
          <div class="s-badge" style="background:{color}1A;color:{color}">{ic(icon, color, 20)}</div>
          <div class="s-body">
            <div class="s-lbl">{label}</div>
            <div class="s-val" style="color:{color}">{disp}{u}</div>
          </div></div>"""

    # ---- target stat card: actual vs target with slim achievement bar
    def tstat(icon, label, actual, target):
        col, p = rag(actual, target)
        if p is not None:
            meta = f"""<div class="s-bar"><div class="s-bar-f" style="width:{min(p,100)}%;background:{col}"></div></div>
              <div class="s-meta" style="color:{col}">{p}% of {fmt(target)} MT target</div>"""
        else:
            meta = '<div class="s-meta s-muted">no target set</div>'
        return f"""<div class="stat stat-wide">
          <div class="s-badge" style="background:{col}1A;color:{col}">{ic(icon, col, 20)}</div>
          <div class="s-body">
            <div class="s-lbl">{label}</div>
            <div class="s-val" style="color:{col}">{fmt(actual)}<span class="s-unit">MT</span></div>
            {meta}
          </div></div>"""

    # ---- compact metric cell (for the grouped Process/Quality grid)
    def cell(icon, label, value, unit="", accent=None):
        accent = accent or c["navy"]
        return f"""<div class="cell">
          <div class="c-top">{ic(icon, c['grey'], 15)}<span class="c-lbl">{label}</span></div>
          <div class="c-val">{fmt(value)}<span class="c-unit">{unit}</span></div></div>"""

    # ---- process panel: professional grouped KPI layout
    def mill_panel(name, m, head_color):
        if not m:
            return ""
        delay = ""
        if m.get("delay_hrs"):
            delay = f'<div class="delaybar">{ic("delay",c["red"],15)} Delay {fmt(m["delay_hrs"])} hrs</div>'
        else:
            dl = []
            if m.get("delay_op"):   dl.append(f"Op {m['delay_op']}")
            if m.get("delay_mech"): dl.append(f"Mech {m['delay_mech']}")
            if dl:
                delay = f'<div class="delaybar">{ic("delay",c["red"],15)} Delay: {" · ".join(dl)}</div>'
        day = m.get("day_total", 0); tgt = m.get("day_target", 0)
        col, p = rag(day, tgt)
        gap = (day - tgt) if tgt else None
        gaptxt = ""
        if gap is not None:
            sign = "▲" if gap >= 0 else "▼"
            gaptxt = f'<span class="ph-gap" style="color:{"#bfe3c8" if gap>=0 else "#f6c6c2"}">{sign} {fmt(abs(round(gap,3)))} MT</span>'
        cumm_total = m.get("cumm_out") or m.get("cumm_total", 0)

        # grouped metric grid — only cells that exist
        cells = []
        if m.get("yield") is not None and m.get("yield"):
            cells.append(("Quality", cell("yield", "Day Yield", round(m["yield"],2), "%")))
        if m.get("cumm_yield"):
            cells.append(("Quality", cell("yield", "Cumm Yield", round(m["cumm_yield"],2), "%")))
        if m.get("day_util"):
            cells.append(("Process", cell("util", "Utilization", round(m["day_util"],1), "%")))
        if m.get("util_tilldate"):
            cells.append(("Process", cell("util", "Util Till Date", round(m["util_tilldate"],1), "%")))
        if m.get("avg_gauge"):
            cells.append(("Process", cell("gauge", "Avg Gauge", round(m["avg_gauge"],3), " mm")))
        if m.get("avg_width"):
            cells.append(("Process", cell("width", "Avg Width", round(m["avg_width"],1), " mm")))
        if m.get("coils"):
            cells.append(("Process", cell("coils", "Coils", round(m["coils"]), "")))
        grid = ""
        if cells:
            grid = ('<div class="grp-lbl">Process &amp; Quality Metrics</div>'
                    '<div class="cellgrid">' + "".join(x[1] for x in cells) + '</div>')

        return f"""
        <div class="panel">
          <div class="phead" style="background:{head_color}">{ic('roll','#dfe6ee',20)}<span>{name}</span>
            <span class="ph-tot">Day {fmt(day)}{(' / Target '+fmt(tgt)) if tgt else ''} MT {gaptxt}</span></div>
          <div class="statrow">
            {tstat('target','Day Production', day, tgt)}
            {stat('roll','Roll', m.get('day_roll',0), c['blue'])}
            {stat('rr','Re-Roll', m.get('day_rr',0), c['blue'])}
            {stat('skp','Skin Pass', m.get('day_skp', m.get('skp',0)), c['blue'])}
            {stat('cumm','Cumm Output', cumm_total, c['orange'])}
          </div>
          {grid}
          <div class="achv-bar" style="border-color:{col};background:{col}0F">{ic('yield' if (p or 0)>=100 else 'target',col,15)}
            <b style="color:{col}">{(str(p)+'% of day target achieved') if p is not None else 'Day target not set'}</b></div>
          {delay}
        </div>"""

    rolling = d.get("rolling", {})
    head_colors = [c["green"], c["blue"], c["navy"], c["orange"]]
    mill_names = [k for k in rolling.keys() if k != "REWINDING06"]
    panels_html = ""
    for i, mn in enumerate(mill_names):
        panels_html += mill_panel(mn, rolling.get(mn), head_colors[i % len(head_colors)])

    # ---- headline KPI cards (top-right of masthead)
    hl = d.get("headline", {})
    target = hl.get("gr_target", c["gr_target"]); mtd = hl.get("gr_mtd", 0); today = hl.get("gr_today", 0)
    pct = round(100 * mtd / target, 1) if target else 0
    headline = f"""
      <div class="hcard hcard-wide">
        <div class="hc-t">{ic('target',c['navy'],14)} TOTAL GR &nbsp;·&nbsp; TODAY / MTD / TARGET</div>
        <div class="hc-v"><b style="color:{c['blue']}">{fmt(today)}</b> / <b style="color:{c['green']}">{mtd}</b> / <b style="color:{c['orange']}">{target}</b> <small>T</small></div>
        <div class="hbar"><div class="hbar-f" style="width:{min(pct,100)}%"></div></div>
        <div class="hc-pct">{ic('yield',c['green'],12)} {pct}% of MTD target achieved</div>
      </div>"""

    # ---- annealing panel (from PRODUCTION & GAS DETAILS sheet)
    ann = d.get("annealing", {}).get("ann02", {})
    ann_panel = ""
    if ann:
        cells = []
        if ann.get("prod_old") is not None:
            cells.append(cell("roll", "Prod OLD", round(ann.get("prod_old",0),3), " MT"))
        if ann.get("prod_new") is not None:
            cells.append(cell("cumm", "Prod NEW", round(ann.get("prod_new",0),3), " MT"))
        if ann.get("charges"):
            cells.append(cell("production", "Charges", ann.get("charges",0), f"  ({int(ann.get('chg_old',0))}+{int(ann.get('chg_new',0))})"))
        if ann.get("water"):
            cells.append(cell("width", "Water Cons", round(ann.get("water",0)), " m³"))
        if ann.get("lng_nm3"):
            cells.append(cell("util", "LNG", round(ann.get("lng_nm3",0)), " Nm³"))
        grid = ('<div class="grp-lbl">Production · Charges · Utilities</div>'
                '<div class="cellgrid">' + "".join(cells) + '</div>') if cells else ""
        ann_panel = f"""
        <div class="midcard">
          <div class="mh" style="background:{c['navy']}">{ic('production','#dfe6ee',18)} Annealing (ANN) — Day Summary</div>
          <div class="statrow statrow-mid">
            {stat('production','Day Production', ann.get('day_prod',0), c['green'])}
            {stat('production','Total Charges', ann.get('charges',0), c['blue'], unit='')}
          </div>
          {grid}
        </div>"""

    # ---- 2HI / SPM panel (PROD + ID CHANGE + HROP)
    sp = d.get("annealing", {}).get("skin_pass", {})
    sp_panel = ""
    if sp:
        note = f'<div class="notebar">{ic("yield",c["navy"],14)} {sp.get("carol_drum")}</div>' if sp.get("carol_drum") else ""
        sp_panel = f"""
        <div class="midcard">
          <div class="mh" style="background:{c['navy']}">{ic('skp','#dfe6ee',18)} 2HI / Skin Pass — Day Summary</div>
          <div class="statrow statrow-mid">
            {stat('skp','2HI Prod', sp.get('hi_prod',0), c['green'])}
            {stat('rr','ID Change', sp.get('id_change',0), c['blue'])}
            {stat('cumm','HROP', sp.get('hrop',0), c['orange'])}
          </div>
          <div class="grp-lbl">Total Day Output</div>
          <div class="cellgrid">{cell('production','Day Total', round(sp.get('day_prod',0),3), ' MT')}</div>
          {note}
        </div>"""

    # ---- rewinding mini card
    rew = rolling.get("REWINDING06")
    rew_card = ""
    if rew:
        rew_card = f"""
        <div class="minicard">
          <div class="mh" style="background:{c['navy']}">{ic('roll','#dfe6ee',18)} Rewinding 06 — Day Summary</div>
          <div class="statrow statrow-mid">
            {stat('production','Day Total', rew.get('day_total', rew.get('shift_total',0)), c['green'])}
            {stat('cumm','Cumulative', rew.get('cumm_total',0), c['orange'])}
          </div>
        </div>"""

    # ---- packing GR pie + tiles
    sl = d.get("slitting", {})
    pg = sl.get("packing_gr", {})
    pack_card = ""
    if pg:
        tube = pg.get("tube_gr", 0); oem = pg.get("oem_gr", 0); tot = (tube + oem) or 1
        tpct = round(100 * tube / tot, 1); opct = round(100 * oem / tot, 1)
        pack_card = f"""
        <div class="minicard">
          <div class="mh" style="background:{c['navy']}">{ic('box','#cfd8e3',20)} Packing (GR) Breakup</div>
          <div class="pie-wrap">
            <div class="pie" style="background:conic-gradient({c['green']} 0 {tpct}%, {c['orange']} {tpct}% 100%)"></div>
            <div class="legend">
              <div><span class="dot" style="background:{c['green']}"></span> Tube GR <b>{fmt(tube)} MT</b> ({tpct}%)</div>
              <div><span class="dot" style="background:{c['orange']}"></span> OEM GR <b>{fmt(oem)} MT</b> ({opct}%)</div>
              <div class="sap">SAP Cumm — Tube {fmt(pg.get('tube_sap',0))} · OEM {fmt(pg.get('oem_sap',0))}</div>
            </div>
          </div>
        </div>"""

    # ---- slitting table (Day summary)
    sl_rows = ""
    for ln in sl.get("lines", []):
        sl_rows += f"<tr><td>{ln.get('name','')}</td><td>{fmt(ln.get('day',''))}</td><td>{fmt(ln.get('cumm',''))}</td><td class='rmk'>{ln.get('remark','')}</td></tr>"
    slit_card = ""
    if sl_rows:
        t = sl.get("total", {})
        slit_card = f"""
        <div class="botcard">
          <div class="mh" style="background:{c['navy']}">{ic('chart','#cfd8e3',20)} Slitting — Day Summary</div>
          <table class="slt">
            <tr><th>Line</th><th>Day</th><th>Cumm</th><th>Remark</th></tr>
            {sl_rows}
            <tr class="tot"><td>TOTAL</td><td>{fmt(t.get('day',''))}</td><td>{fmt(t.get('cumm',''))}</td><td></td></tr>
          </table>
        </div>"""

    # ---- notes card
    notes = sl.get("notes", [])
    notes_card = ""
    if notes:
        lis = "".join(f"<li>{x}</li>" for x in notes)
        notes_card = f"""
        <div class="botcard">
          <div class="mh" style="background:{c['green']}">{ic('shield','#cfd8e3',20)} Key Notes & Remarks</div>
          <ul class="notes">{lis}</ul>
        </div>"""

    logo = ""
    if c["logo_path"] and os.path.exists(c["logo_path"]):
        logo = f'<img src="file://{c["logo_path"]}" class="logo">'
    else:
        logo = f'<div class="logo-txt">{ic("factory","#ffffff",28)}<div><b>TATA STEEL</b><div class="logo-sub">Narrow Complex</div></div></div>'

    html = f"""<!doctype html><html><head><meta charset="utf-8"><style>
    *{{box-sizing:border-box;margin:0;padding:0;font-family:Calibri,'Segoe UI',Arial,sans-serif}}
    body{{width:1500px;height:1000px;background:{c['page_bg']};padding:14px;overflow:hidden}}
    .svgic{{display:inline-block;vertical-align:middle;flex-shrink:0}}
    .em{{display:none}}
    /* masthead */
    .top{{background:{c['navy']};border-radius:10px;display:flex;align-items:stretch;color:#fff;height:92px;overflow:hidden}}
    .logo-txt{{display:flex;align-items:center;gap:10px;padding:0 18px;border-right:2px solid rgba(255,255,255,.2)}}
    .logo-txt b{{font-size:20px;letter-spacing:1px}}.logo-sub{{font-size:11px;opacity:.8}}
    .ttl{{flex:1;display:flex;flex-direction:column;justify-content:center;padding-left:22px}}
    .ttl h1{{font-size:27px;letter-spacing:.5px}}.ttl .dt{{font-size:13px;opacity:.85;margin-top:3px}}
    .hcard{{background:#fff;color:{c['navy']};border-radius:8px;margin:9px;padding:8px 18px;display:flex;flex-direction:column;justify-content:center;width:340px;flex-shrink:0}}
    .hc-t{{font-size:10.5px;font-weight:bold;letter-spacing:.6px;color:{c['grey']}}}
    .hc-v{{font-size:27px;font-weight:bold;margin:3px 0;letter-spacing:.5px}}.hc-v small{{font-size:14px}}
    .hbar{{height:9px;background:#e4e8ee;border-radius:5px;overflow:hidden;margin-top:3px}}
    .hbar-f{{height:100%;background:linear-gradient(90deg,{c['blue']},{c['green']})}}
    .hc-pct{{font-size:11.5px;font-weight:bold;color:{c['green']};text-align:right;margin-top:2px}}
    /* process panels */
    .prow{{display:flex;gap:12px;margin-top:12px}}
    .prow>.panel{{flex:1}}
    .panel{{background:#fff;border-radius:12px;box-shadow:0 1px 6px rgba(20,40,80,.08);overflow:hidden;border:1px solid #eaeef3}}
    .phead{{color:#fff;font-size:15px;font-weight:bold;padding:9px 16px;display:flex;align-items:center;gap:10px;letter-spacing:.3px}}
    .ph-tot{{margin-left:auto;font-size:12.5px;font-weight:600;background:rgba(255,255,255,.16);padding:3px 12px;border-radius:14px;display:flex;align-items:center;gap:7px}}
    .ph-gap{{font-size:12px;font-weight:bold}}
    /* primary stat row */
    .statrow{{display:flex;padding:13px 10px 9px;gap:0}}
    .stat{{flex:1;display:flex;align-items:center;gap:7px;padding:0 8px;border-right:1px solid #eef1f5;min-width:0}}
    .stat:last-child{{border-right:none;padding-right:4px}}
    .stat-wide{{flex:1.3}}
    .s-badge{{width:34px;height:34px;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0}}
    .s-body{{min-width:0;overflow:hidden}}
    .s-lbl{{font-size:9px;font-weight:bold;color:{c['grey']};text-transform:uppercase;letter-spacing:.4px;white-space:nowrap}}
    .s-val{{font-size:19px;font-weight:bold;line-height:1.15;letter-spacing:.2px;white-space:nowrap}}
    .s-unit{{font-size:10px;font-weight:600;color:{c['grey']};margin-left:2px}}
    .s-bar{{height:4px;background:#e7ebf0;border-radius:3px;overflow:hidden;margin:3px 0 2px;max-width:130px}}
    .s-bar-f{{height:100%}}
    .s-meta{{font-size:10px;font-weight:bold}}.s-muted{{color:#9aa3ab;font-weight:normal}}
    /* grouped metric grid */
    .grp-lbl{{font-size:10px;font-weight:bold;color:{c['grey']};text-transform:uppercase;letter-spacing:.7px;padding:2px 16px 6px;border-top:1px solid #eef1f5;margin-top:2px}}
    .cellgrid{{display:flex;flex-wrap:wrap;padding:0 12px 8px;gap:1px}}
    .cell{{flex:1;min-width:96px;padding:6px 10px;background:#f7f9fc;border-radius:7px;margin:2px}}
    .c-top{{display:flex;align-items:center;gap:5px;margin-bottom:2px}}
    .c-lbl{{font-size:9.5px;font-weight:600;color:{c['grey']};text-transform:uppercase;letter-spacing:.3px}}
    .c-val{{font-size:15px;font-weight:bold;color:{c['navy']}}}
    .c-unit{{font-size:9.5px;font-weight:600;color:{c['grey']};margin-left:1px}}
    .achv-bar{{margin:4px 14px 9px;padding:6px 12px;border:1px solid;border-radius:7px;font-size:12.5px;display:flex;align-items:center;gap:8px}}
    .delaybar{{margin:0 12px 9px;background:#fdeceb;color:{c['red']};font-size:11.5px;font-weight:bold;padding:5px 10px;border-radius:6px;display:flex;align-items:center;gap:6px}}
    /* middle row */
    .midrow{{display:flex;gap:10px;margin-top:10px}}
    .midrow>.midcard{{flex:1.3}}.midrow>.minicard{{flex:0.9}}
    .midcard,.minicard,.botcard{{background:#fff;border-radius:12px;box-shadow:0 1px 6px rgba(20,40,80,.08);overflow:hidden;border:1px solid #eaeef3}}
    .statrow-mid{{flex-wrap:wrap;padding:12px 8px 8px}}
    .statrow-mid .stat{{flex:1 1 40%;border-right:1px solid #eef1f5}}
    .notebar{{margin:2px 14px 10px;padding:5px 11px;background:#eef3f9;border-radius:7px;font-size:11.5px;color:{c['navy']};display:flex;align-items:center;gap:6px}}
    .mh{{color:#fff;font-size:13px;font-weight:bold;padding:7px 13px;display:flex;align-items:center;gap:8px}}
    .permit{{font-size:11px;color:{c['grey']};padding:6px 13px 0}}
    .mtiles{{display:flex;justify-content:space-around;padding:8px;gap:4px;flex-wrap:wrap}}
    .kvlist{{font-size:12px;padding:0 13px 7px}}.kvlist div{{margin:2px 0}}
    /* bottom row */
    .botrow{{display:flex;gap:10px;margin-top:10px}}
    .botrow>.botcard:nth-child(1){{flex:1.1}}.botrow>.minicard{{flex:0.85}}.botrow>.botcard:nth-child(3){{flex:0.9}}
    .slt{{width:100%;border-collapse:collapse;font-size:12px}}
    .slt th{{background:#eef2f7;color:{c['navy']};padding:5px 9px;text-align:left;font-size:11px}}
    .slt td{{padding:4px 9px;border-bottom:1px solid #eee}}
    .slt .rmk{{color:{c['grey']};font-size:10.5px}}
    .slt .tot td{{font-weight:bold;background:#fff6ec;border-top:2px solid {c['orange']}}}
    .notes{{margin:8px 18px;font-size:12.5px}}.notes li{{margin:5px 0}}
    .pie-wrap{{display:flex;align-items:center;gap:14px;padding:12px 14px}}
    .pie{{width:90px;height:90px;border-radius:50%;flex-shrink:0}}
    .legend{{font-size:12px}}.legend div{{margin:4px 0}}
    .dot{{display:inline-block;width:11px;height:11px;border-radius:2px;margin-right:5px}}
    .sap{{margin-top:6px;font-size:10.5px;color:{c['grey']}}}
    .foot{{background:{c['orange']};color:#fff;font-size:11px;text-align:center;padding:6px;border-radius:8px;margin-top:11px}}
    </style></head><body>
    <div class="top">
      {logo}
      <div class="ttl"><h1>NARROW COMPLEX — DAY SUMMARY REPORT</h1>
        <div class="dt">{c['plant']} &nbsp;|&nbsp; Date: {d.get('date','—')} &nbsp;|&nbsp; Production · GR · Cumulative</div></div>
      {headline}
    </div>

    <div class="prow">{panels_html}</div>

    <div class="midrow">{ann_panel}{sp_panel}{rew_card}</div>

    <div class="botrow">{slit_card}{pack_card}{notes_card}</div>

    <div class="foot">Tata Steel — Narrow Complex  |  Auto-generated Day Summary  |  {datetime.date.today().isoformat()}</div>
    </body></html>"""

    return html


# -----------------------------------------------------------------------------
def render_png(html_path, png_path):
    # Primary: Playwright/Chromium — renders COLOR emoji + pie chart correctly.
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(viewport={"width": 1500, "height": 1000})
            pg.goto("file://" + os.path.abspath(html_path))
            pg.screenshot(path=png_path, clip={"x": 0, "y": 0, "width": 1500, "height": 1000})
            b.close()
        return png_path
    except Exception as e:
        print("  [PNG] Playwright unavailable (", e, ") — install with:")
        print("        pip install playwright && python -m playwright install chromium")
        print("  [PNG] falling back to wkhtmltoimage (emojis may appear monochrome)")
    # Fallback: wkhtmltoimage (older WebKit; emojis render monochrome)
    if shutil.which("wkhtmltoimage"):
        try:
            subprocess.run(["wkhtmltoimage", "--width", "1500", "--height", "1000",
                            "--disable-smart-width", "--enable-local-file-access",
                            html_path, png_path], check=True,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return png_path
        except Exception as e:
            print("  [PNG] wkhtmltoimage failed:", e)
    print("  [PNG] no working HTML->image binary; skipping PNG")
    return None


# -----------------------------------------------------------------------------
#  RENDER — PPTX  (python-pptx, single branded slide)
# -----------------------------------------------------------------------------
def render_pptx(data, path):
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    def C(hexs):
        return RGBColor.from_string(hexs.lstrip("#"))

    c = CONFIG
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])

    def box(x, y, w, h, fill=None, line=None):
        from pptx.enum.shapes import MSO_SHAPE
        sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid(); sh.fill.fore_color.rgb = C(fill or "FFFFFF")
        if line: sh.line.color.rgb = C(line); sh.line.width = Pt(0.75)
        else: sh.line.fill.background()
        sh.shadow.inherit = False
        return sh

    def text(x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Pt(2); tf.margin_top = tf.margin_bottom = Pt(1)
        first = True
        for (txt, size, color, bold) in runs:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False; p.alignment = align
            r = p.add_run(); r.text = txt
            r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = C(color)
            r.font.name = "Calibri"
        return tb

    d = data
    # Title bar
    box(0, 0, 13.333, 1.05, fill=c["navy"].lstrip("#"))
    text(0.35, 0.12, 9, 0.5, [("C-SHIFT PRODUCTION DASHBOARD", 26, "FFFFFF", True)])
    text(0.37, 0.62, 9, 0.35,
         [(f"{c['plant']}  •  Date: {d.get('date','—')}  •  Shift {d.get('shift','C')}", 11, "D7E1F0", False)])

    hl = d.get("headline", {})
    if hl:
        target = hl.get("gr_target", c["gr_target"]); mtd = hl.get("gr_mtd", 0)
        pct = round(100 * mtd / target, 1) if target else 0
        text(9.6, 0.18, 3.4, 0.7,
             [(f"GR  {hl.get('gr_today','—')} / {mtd} / {target} T", 15, "FFFFFF", True),
              (f"Achievement: {pct}%", 11, "F8C9A0", False)], align=PP_ALIGN.RIGHT)

    # --- card layout grid ---
    gx, gy, gw, gh, gap = 0.35, 1.25, 4.1, 2.55, 0.2
    col = [gx, gx + gw + gap, gx + 2 * (gw + gap)]

    def mill_block(x, y, name, m):
        if not m: return
        box(x, y, gw, gh, fill="FFFFFF", line="E2E6EC")
        box(x, y, gw, 0.42, fill=c["navy"].lstrip("#"))
        text(x + 0.12, y + 0.05, gw - 0.2, 0.34, [(name + "  — C-Shift", 14, "FFFFFF", True)])
        rows = [
            ("Shift", f"{m.get('shift_total',0):g}", f"R {m.get('roll',0):g} / RR {m.get('rr',0):g} / SKP {m.get('skp',0):g}", c["blue"]),
            ("Day",   f"{m.get('day_total',0):g}",   f"R {m.get('day_roll',0):g} / RR {m.get('day_rr',0):g}", c["green"]),
            ("Cumm",  f"{m.get('cumm_total',0):g}",   f"R {m.get('cumm_roll',0):g} / RR {m.get('cumm_rr',0):g} / SKP {m.get('cumm_skp',0):g}", c["orange"]),
        ]
        yy = y + 0.55
        for lbl, val, sub, color in rows:
            text(x + 0.15, yy, 1.45, 0.55, [(lbl, 10, "5F6368", True), (val + " MT", 14, color.lstrip("#"), True)])
            text(x + 1.7, yy + 0.18, gw - 1.85, 0.45, [(sub, 9, "5F6368", False)])
            yy += 0.6
        if m.get("delay_op") or m.get("delay_mech"):
            dly = []
            if m.get("delay_op"): dly.append(f"Op {m['delay_op']}")
            if m.get("delay_mech"): dly.append(f"Mech {m['delay_mech']}")
            text(x + 0.15, y + gh - 0.34, gw - 0.3, 0.3, [("⚠ Delay: " + " · ".join(dly), 10, c["red"].lstrip("#"), True)])

    rolling = d.get("rolling", {})
    mill_block(col[0], gy, "CRM04", rolling.get("CRM04"))
    mill_block(col[1], gy, "CRM06", rolling.get("CRM06"))

    # Annealing card (col 3)
    ann = d.get("annealing", {}).get("ann02", {})
    if ann:
        x, y = col[2], gy
        box(x, y, gw, gh, fill="FFFFFF", line="E2E6EC")
        box(x, y, gw, 0.42, fill=c["navy"].lstrip("#"))
        text(x + 0.12, y + 0.05, gw - 0.2, 0.34, [("Annealing — ANN02", 14, "FFFFFF", True)])
        kv = [("Shift Prod", ann.get('shift_prod',0)), ("Day Prod", ann.get('day_prod',0)),
              ("Cumm Prod", ann.get('cumm_prod',0)), ("H&T", ann.get('ht',0)),
              ("Await", ann.get('await',0)), ("Furnace", ann.get('furnace','—'))]
        yy = y + 0.55
        for i, (k, v) in enumerate(kv):
            xx = x + 0.15 + (i % 2) * (gw/2)
            if i % 2 == 0 and i: yy += 0.52
            disp = f"{v:g} MT" if isinstance(v,(int,float)) else str(v)
            text(xx, yy, gw/2 - 0.2, 0.5, [(k, 9, "5F6368", True), (disp, 14, c["orange"].lstrip("#"), True)])
        yy += 0.52

    # Bottom row: Skin/RWL, Rewinding, Slitting summary
    by, bh = gy + gh + 0.2, 2.45
    sp = d.get("annealing", {}).get("skin_pass", {}); rwl = d.get("annealing", {}).get("rwl02", {})
    if sp or rwl:
        x = col[0]
        box(x, by, gw, bh, fill="FFFFFF", line="E2E6EC")
        box(x, by, gw, 0.42, fill=c["navy"].lstrip("#"))
        text(x + 0.12, by + 0.05, gw - 0.2, 0.34, [("Skin Pass (2HI) & RWL02", 14, "FFFFFF", True)])
        text(x + 0.15, by + 0.55, gw - 0.3, bh - 0.6,
             [(f"RWL02 Shift Prod: {rwl.get('shift_prod',0):g} MT", 12, "143B6B", True),
              (f"SKP C-Shift: {sp.get('shift_prod','—')}", 11, "5F6368", False),
              (f"SKP Day Prod: {sp.get('day_prod','—')}", 11, "5F6368", False),
              (f"SKP Await: {sp.get('await',0):g} MT", 11, "5F6368", False),
              (f"SKP Cumm: {sp.get('cumm_prod',0):g} MT", 11, c["orange"].lstrip("#"), True),
              (f"RWL02 Delay: {rwl.get('delay','—')}", 10, c["red"].lstrip("#"), False)])

    rew = rolling.get("REWINDING06")
    if rew:
        x = col[1]
        box(x, by, gw, bh, fill="FFFFFF", line="E2E6EC")
        box(x, by, gw, 0.42, fill=c["navy"].lstrip("#"))
        text(x + 0.12, by + 0.05, gw - 0.2, 0.34, [("Rewinding 06 — C-Shift", 14, "FFFFFF", True)])
        text(x + 0.15, by + 0.7, gw - 0.3, 1.5,
             [("Shift Total", 11, "5F6368", True), (f"{rew.get('shift_total',0):g} MT", 26, "143B6B", True),
              ("Cumulative", 11, "5F6368", True), (f"{rew.get('cumm_total',0):g} MT", 22, c["orange"].lstrip("#"), True)])

    sl = d.get("slitting", {})
    if sl.get("lines") or sl.get("notes"):
        x = col[2]
        box(x, by, gw, bh, fill="FFFFFF", line="E2E6EC")
        box(x, by, gw, 0.42, fill=c["navy"].lstrip("#"))
        text(x + 0.12, by + 0.05, gw - 0.2, 0.34, [("Slitting & Notes", 14, "FFFFFF", True)])
        tot = sl.get("total", {})
        runs = [(f"Slitting Day Total: {tot.get('day','—')}", 12, "143B6B", True)]
        for nt in sl.get("notes", [])[:4]:
            runs.append(("• " + nt, 10, "5F6368", False))
        text(x + 0.15, by + 0.55, gw - 0.3, bh - 0.6, runs)

    # Footer
    box(0, 7.12, 13.333, 0.38, fill=c["orange"].lstrip("#"))
    text(0, 7.16, 13.333, 0.3,
         [("Tata Steel — Narrow Complex  |  Auto-generated shift report", 10, "FFFFFF", False)],
         align=PP_ALIGN.CENTER)

    prs.save(path)
    return path


# -----------------------------------------------------------------------------
#  DEMO DATA  (real values from the three C-shift reports)
# -----------------------------------------------------------------------------
def demo_data():
    return {
        "date": "29-05-2026", "shift": "C",
        "headline": {"gr_today": 243, "gr_mtd": 6547, "gr_target": 7000},
        "rolling": {
            "CRM04": {"day_total":138.800,"day_roll":90.525,"day_rr":48.275,"day_target":150,
                      "delay_op":"30 Mins","delay_mech":"25 Min",
                      "cumm_total":4571.988},
            "CRM06": {"day_total":280.985,"day_roll":239.520,"day_rr":41.465,"day_target":260,
                      "delay_op":"60 Min","delay_mech":"",
                      "cumm_total":5265.165},
            "REWINDING06": {"day_total":43.670,"cumm_total":1351.560},
        },
        "annealing": {
            "ann02": {"e_permit":"110021653577 (CLOSED)","ht":362.280,"await":305.235,
                      "day_prod":187.823,"cumm_prod":4745.498,
                      "furnace":"5","delay":"HH#2 - 8 Hrs (No Material)"},
            "rwl02": {"day_prod":35.225,"delay":"2 Hrs (No Material)"},
            "skin_pass": {"await":234.605,
                          "day_prod":"78.617 (TUBE) + 42.245 (ID CHANGE)","cumm_prod":1439.945},
        },
        "slitting": {
            "lines": [
                {"name":"CRS09","day":0,"cumm":0,"remark":"8 Hrs - No Man Power"},
                {"name":"CRS10","day":85.38,"cumm":1934.803,"remark":"1.5 Hrs Operational"},
                {"name":"CRS11","day":76.54,"cumm":1900.133,"remark":"No Delay"},
                {"name":"CRS13","day":13.52,"cumm":463.743,"remark":"8 Hrs - No Man Power"},
                {"name":"CRS14","day":80.545,"cumm":1471.315,"remark":"2 Hrs - Preprty Check"},
            ],
            "total": {"day":256,"cumm":5769.99},
            "packing_gr": {"tube_gr":112.83,"oem_gr":67.208,"tube_sap":4354,"oem_sap":1006},
            "notes": ["3 High Carbon coil packed, wt. pending",
                      "Dispatch: 2 trucks OEM (L.G-1, VAISH-1)"],
        },
    }


# -----------------------------------------------------------------------------
#  ORCHESTRATION
# -----------------------------------------------------------------------------
def build_outputs(data):
    os.makedirs(CONFIG["output_dir"], exist_ok=True)
    tag = (data.get("date") or datetime.date.today().isoformat()).replace("/", "-")
    base = os.path.join(CONFIG["output_dir"], f"dashboard_{tag}")
    html = base + ".html"; png = base + ".png"; pptx = base + ".pptx"
    results = {}
    for fmt, fn in [("HTML", lambda: render_html(data, html)),
                    ("PPTX", lambda: render_pptx(data, pptx))]:
        try:
            results[fmt] = fn(); print(f"  [{fmt}] -> {results[fmt]}")
        except Exception as e:
            print(f"  [{fmt}] FAILED: {e}")
    try:
        r = render_png(html, png)
        if r: results["PNG"] = r; print(f"  [PNG] -> {r}")
    except Exception as e:
        print(f"  [PNG] FAILED: {e}")
    return results


def interactive():
    print("Paste the rolling report text, then Ctrl-D (blank line + Ctrl-D to skip):")
    try:
        raw = sys.stdin.read()
    except KeyboardInterrupt:
        raw = ""
    data = demo_data() if not raw.strip() else {}
    if raw.strip():
        rolling, date = parse_rolling_text(raw)
        data = {"date": date or datetime.date.today().isoformat(), "shift": "C", "rolling": rolling}
    build_outputs(data)


def main():
    args = sys.argv[1:]
    if "--demo" in args or not args:
        print("Rendering with built-in sample data..." if "--demo" in args else "No args; using demo. (use --help)")
        build_outputs(demo_data()); return
    if "--text" in args:
        p = args[args.index("--text") + 1]
        rolling, date = parse_rolling_text(open(p).read())
        build_outputs({"date": date, "shift": "C", "rolling": rolling}); return
    if "--image" in args:
        imgs = args[args.index("--image") + 1:]
        merged = "\n".join(ocr_extract(p) for p in imgs)
        print("\n--- OCR text (review parsed values below) ---")
        ann = parse_annealing_text(merged)
        rolling, date = parse_rolling_text(merged)
        build_outputs({"date": date, "shift": "C", "rolling": rolling, "annealing": ann}); return
    if "--mis" in args or "--anneal" in args:
        day = None
        if "--day" in args:
            try: day = int(args[args.index("--day") + 1])
            except Exception: day = None
        data = {"date": str(day) if day else "—", "shift": "Day",
                "headline": CONFIG.get("gr_headline", {})}
        if "--mis" in args:
            rolling = parse_mis(open(args[args.index("--mis") + 1]).read(), report_day=day)
            for mill, m in rolling.items():
                t = CONFIG.get("rolling_targets", {}).get(mill)
                if t: m["day_target"] = t
            data["rolling"] = rolling
        if "--anneal" in args:
            data["annealing"] = parse_anneal_mis(open(args[args.index("--anneal") + 1]).read(), report_day=day)
        build_outputs(data); return
    interactive()




# ===========================================================================
# EXTRACTOR: targets

# ===========================================================================

# extractors/targets.py
# Reads Target.xlsx  (KPI | Day Target | MTD Target)
# Returns a flat dict:  targets["rolling_day"] = 200, etc.

import openpyxl
import sys, os

# from config import TARGET_ROWS  # merged inline

def _to_num(v, default=None):
    if v is None or str(v).strip() in ("-", "NA", ""):
        return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default

def parse_targets(filepath):
    """Return dict of all target values. Missing/NA → None."""
    targets = {}
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        kpi = str(row[0]).strip() if row[0] else ""
        if kpi in TARGET_ROWS:
            day_key, mtd_key = TARGET_ROWS[kpi]
            targets[day_key] = _to_num(row[1] if len(row) > 1 else None)
            if mtd_key:
                targets[mtd_key] = _to_num(row[2] if len(row) > 2 else None)
    wb.close()
    return targets


# ── standalone test ───────────────────────────────────────────────────────────


# ===========================================================================
# EXTRACTOR: crs

# ===========================================================================

# extractors/crs.py
# Reads CRS MIS.xlsx — Narrow Finishing status sheet.
#
# Layout (verified against uploaded file):
#   Col 1 = section label, Col 2 = sub-label, Col 4 = value, Col 6 = note/cumm
#   The report date is embedded in Row 1, Col 1.

import re
import openpyxl
import datetime

def _n(v, default=0.0):
    if v is None or str(v).strip() in ("", "-", "NA"):
        return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default

def parse_crs(filepath):
    """
    Returns a dict:
    {
      "report_date": "30.05.2026",
      "slitting":  {"oem":95, "tube":277, "strapping":0, "total":372},
      "packing":   {"oem":158, "tube":60, "total":218, "note":"108 L.G"},
      "no_plan":   12,
      "hrop_skp":  128,
      "hold":      {"value":95, "note":"46 LG"},
      "skp_wip":   178,
      "total_at_crs": 825,
      "gr": {
          "tube":  {"yesterday":110, "cumm":4425, "target_label":"TUBE/5000"},
          "oem":   {"yesterday":58,  "cumm":1045, "target_label":"OEM/1200"},
          "total": {"yesterday":168, "cumm":5470},
      },
      "crca_slitting": {"yesterday":256, "cumm":5770},
    }
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active

    out = {
        "report_date": "",
        "slitting":    {"oem": 0, "tube": 0, "strapping": 0, "total": 0},
        "packing":     {"oem": 0, "tube": 0, "total": 0, "note": ""},
        "no_plan":     0,
        "hrop_skp":    0,
        "hold":        {"value": 0, "note": ""},
        "skp_wip":     0,
        "total_at_crs": 0,
        "gr": {
            "tube":  {"yesterday": 0, "cumm": 0, "target_label": ""},
            "oem":   {"yesterday": 0, "cumm": 0, "target_label": ""},
            "total": {"yesterday": 0, "cumm": 0},
        },
        "crca_slitting": {"yesterday": 0, "cumm": 0},
    }

    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Row 1: extract date
    if rows:
        m = re.search(r"(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4})", str(rows[0][0] or ""))
        if m:
            out["report_date"] = m.group(1)

    section = None
    for row in rows[1:]:
        col1 = str(row[0] or "").strip()
        col2 = str(row[1] or "").strip()
        col4 = row[3] if len(row) > 3 else None   # 0-based index 3 = col D
        col6 = row[5] if len(row) > 5 else None   # col F

        if col1:
            section = col1.upper()

        # --- SLITTING ---
        if section == "FOR SLITTING":
            sub = col2.upper()
            if sub == "OEM":       out["slitting"]["oem"]       = _n(col4)
            elif sub == "TUBE":    out["slitting"]["tube"]      = _n(col4)
            elif sub in ("STRAPING","STRAPPING"):
                                   out["slitting"]["strapping"] = _n(col4)
            elif sub == "TOTAL":   out["slitting"]["total"]     = _n(col4)

        # --- PACKING ---
        elif section == "FOR PACKING":
            sub = col2.upper()
            if sub == "OEM":       out["packing"]["oem"]   = _n(col4)
            elif sub == "TUBE":    out["packing"]["tube"]  = _n(col4)
            elif sub == "TOTAL":   out["packing"]["total"] = _n(col4)
            if col6:               out["packing"]["note"]  = str(col6).strip()

        # --- SINGLE-VALUE ROWS ---
        elif section == "NO PLAN":     out["no_plan"]      = _n(col4)
        elif section == "HROP S.PASS": out["hrop_skp"]     = _n(col4)
        elif section == "S.PASS WIP":  out["skp_wip"]      = _n(col4)
        elif section == "TOTAL AT CRS":out["total_at_crs"] = _n(col4)

        # HOLD
        elif section == "HOLD":
            out["hold"]["value"] = _n(col4)
            if col6: out["hold"]["note"] = str(col6).strip()

        # --- GR STATUS ---
        elif section == "G.R STATUS":
            sub = col2.upper()
            if "TUBE" in sub:
                out["gr"]["tube"]["yesterday"]    = _n(col4)
                out["gr"]["tube"]["cumm"]         = _n(col6)
                out["gr"]["tube"]["target_label"] = col2
            elif "OEM" in sub:
                out["gr"]["oem"]["yesterday"]     = _n(col4)
                out["gr"]["oem"]["cumm"]          = _n(col6)
                out["gr"]["oem"]["target_label"]  = col2
            elif "TOTAL" in sub:
                out["gr"]["total"]["yesterday"]   = _n(col4)
                out["gr"]["total"]["cumm"]        = _n(col6)

        # --- CRCA SLITTING ---
        elif "CRCA SLITTING" in section:
            out["crca_slitting"]["yesterday"] = _n(col4)
            out["crca_slitting"]["cumm"]      = _n(col6)

    return out


# ── standalone test ───────────────────────────────────────────────────────────


# ===========================================================================
# EXTRACTOR: mill

# ===========================================================================

# extractors/mill.py
# Reads MILL MIS.xlsx — CRM04 / CRM06 daily + MTD production data.
# Wraps the xlsx_to_tsv + parse_mis pipeline already proven in generate_dashboard.py

import sys, os


def parse_mill(filepath, report_day=None):
    """
    Returns dict keyed by mill name (e.g. "CRM06", "CRM07"):
    {
      "CRM06": {
        "day_total", "day_roll", "day_rr", "day_skp",
        "yield", "day_util", "avg_gauge", "avg_width",
        "coils", "delay_hrs",
        "cumm_roll", "cumm_out", "cumm_yield", "util_tilldate",
        "day_target": <from config if set>
      }, ...
    }
    """
    # Reuse the xlsx_to_tsv + parse_mis from generate_dashboard
    proj_root = os.path.dirname(os.path.dirname(__file__))
    
    from generate_dashboard import xlsx_to_tsv, parse_mis

    tsv, sheets = xlsx_to_tsv(filepath)
    mills = parse_mis(tsv, report_day=report_day)
    return mills




# ===========================================================================
# EXTRACTOR: annealing

# ===========================================================================

# extractors/annealing.py
# Reads "Annealing and 2HI SPM MIS.xlsx" — wraps parse_anneal_mis.

import sys, os

def parse_annealing(filepath, report_day=None):
    """
    Returns:
    {
      "ann02": {
        "day_prod", "prod_old", "prod_new",
        "chg_old", "chg_new", "charges",
        "water", "lng_nm3", "lng_m3mt", "row_date"
      },
      "skin_pass": {
        "hi_prod", "id_change", "hrop", "day_prod", "carol_drum"
      }
    }
    """
    
    from generate_dashboard import xlsx_to_tsv, parse_anneal_mis

    tsv, sheets = xlsx_to_tsv(filepath)
    return parse_anneal_mis(tsv, report_day=report_day)




# ===========================================================================
# DASHBOARD: alerts

# ===========================================================================

# dashboard/alerts.py
# Rule-based alert engine.  Returns a list of alert dicts, sorted by priority.
# Priority levels: "critical", "warning", "info"

def generate_alerts(data, targets):
    """
    data    = unified report_data dict (from build_report_data)
    targets = dict from parse_targets
    Returns list of {"priority": str, "title": str, "detail": str}
    """
    alerts = []

    def add(priority, title, detail=""):
        alerts.append({"priority": priority, "title": title, "detail": detail})

    rolling = data.get("rolling", {})
    ann     = data.get("annealing", {}).get("ann02", {})
    spm     = data.get("annealing", {}).get("skin_pass", {})
    crs     = data.get("crs", {})
    gr      = crs.get("gr", {})

    # ── Rolling targets ────────────────────────────────────────────────────
    roll_tgt = targets.get("rolling_day")
    total_roll = sum(m.get("day_total", 0) for m in rolling.values())
    if roll_tgt:
        pct = 100 * total_roll / roll_tgt
        if pct < 90:
            add("critical", f"Rolling target critical — {pct:.1f}%",
                f"Actual {total_roll:.1f} MT vs target {roll_tgt} MT")
        elif pct < 100:
            add("warning", f"Rolling below target — {pct:.1f}%",
                f"Actual {total_roll:.1f} MT vs target {roll_tgt} MT")
        else:
            add("info", f"Rolling on target — {pct:.1f}%",
                f"Combined {total_roll:.1f} MT")

    # ── Per-mill checks ────────────────────────────────────────────────────
    for mill, m in rolling.items():
        ytgt = targets.get(f"{mill.lower()}_yield_mtd")
        utgt = targets.get(f"{mill.lower()}_util_mtd")

        y = m.get("yield")
        if ytgt and y and y < ytgt:
            add("critical", f"{mill} Yield below target",
                f"{y:.2f}% vs {ytgt}% norm")

        u = m.get("day_util")
        if utgt and u and u < utgt:
            severity = "critical" if u < utgt - 10 else "warning"
            add(severity, f"{mill} Utilisation low — {u:.1f}%",
                f"Target {utgt}%")

        if m.get("delay_hrs", 0) > 2:
            add("warning", f"{mill} significant delay",
                f"{m['delay_hrs']:.1f} hrs lost")

    # ── Annealing ──────────────────────────────────────────────────────────
    ann_tgt = targets.get("ann_day")
    ann_prod = ann.get("day_prod", 0)
    if ann_tgt:
        pct = 100 * ann_prod / ann_tgt
        if pct < 90:
            add("critical", f"Annealing target critical — {pct:.1f}%",
                f"{ann_prod:.1f} MT vs {ann_tgt} MT")
        elif pct < 100:
            add("warning", f"Annealing below target — {pct:.1f}%",
                f"{ann_prod:.1f} MT vs {ann_tgt} MT")
        else:
            add("info", "Annealing on target", f"{ann_prod:.1f} MT")

    # ── 2HI / SPM ──────────────────────────────────────────────────────────
    spm_tgt  = targets.get("spm_day")
    spm_prod = spm.get("day_prod", 0)
    if spm_tgt and spm_prod < spm_tgt * 0.9:
        add("warning", f"2HI Skin Pass below target",
            f"{spm_prod:.1f} MT vs {spm_tgt:.1f} MT")

    # ── GR ────────────────────────────────────────────────────────────────
    tgr_tgt  = targets.get("total_gr_day")
    tgr_act  = gr.get("total", {}).get("yesterday", 0)
    if tgr_tgt:
        pct = 100 * tgr_act / tgr_tgt
        if pct < 90:
            add("critical", f"Total GR critical — {pct:.1f}%",
                f"{tgr_act} T vs {tgr_tgt:.0f} T target")
        elif pct < 100:
            add("warning", f"Total GR below target — {pct:.1f}%",
                f"{tgr_act} T vs {tgr_tgt:.0f} T target")
        else:
            add("info", f"GR target achieved — {pct:.1f}%",
                f"{tgr_act} T")

    tube_tgt = targets.get("tube_gr_day")
    tube_act = gr.get("tube", {}).get("yesterday", 0)
    if tube_tgt and tube_act < tube_tgt * 0.9:
        add("warning", "Tube GR below target",
            f"{tube_act} T vs {tube_tgt:.0f} T")

    oem_tgt = targets.get("oem_gr_day")
    oem_act = gr.get("oem", {}).get("yesterday", 0)
    if oem_tgt and oem_act < oem_tgt * 0.9:
        add("warning", "OEM GR below target",
            f"{oem_act} T vs {oem_tgt:.0f} T")

    # ── CRS inventory ─────────────────────────────────────────────────────
    hold = crs.get("hold", {}).get("value", 0)
    hold_max = targets.get("hold_max")
    if hold_max and hold > hold_max:
        add("critical", f"Hold material above threshold",
            f"{hold:.0f} MT (limit {hold_max:.0f} MT)")
    elif hold > 50:
        add("warning", f"Hold material elevated — {hold:.0f} MT")

    skp_wip = crs.get("skp_wip", 0)
    skp_max = targets.get("skp_wip_max")
    if skp_max and skp_wip > skp_max:
        add("warning", f"Skinpass WIP high — {skp_wip:.0f} MT",
            f"Max {skp_max:.0f} MT")
    elif skp_wip > 200:
        add("info", f"Skinpass WIP: {skp_wip:.0f} MT")

    no_plan = crs.get("no_plan", 0)
    if no_plan > 30:
        add("warning", f"No-plan material: {no_plan:.0f} MT",
            "Requires planning attention")

    # ── Sort: critical first, then warning, then info ──────────────────────
    order = {"critical": 0, "warning": 1, "info": 2}
    alerts.sort(key=lambda a: order.get(a["priority"], 3))
    return alerts


# ===========================================================================
# DASHBOARD: notes

# ===========================================================================

# dashboard/notes.py
# Auto-generates up to 5 management notes from KPI values.
# Returns a list of strings; caller renders them in the dashboard.

def generate_notes(data, targets, alerts):
    """
    data    = unified report_data dict
    targets = dict from parse_targets
    alerts  = list from generate_alerts
    Returns list of up to 5 concise management note strings.
    """
    notes = []
    rolling = data.get("rolling", {})
    ann     = data.get("annealing", {}).get("ann02", {})
    spm     = data.get("annealing", {}).get("skin_pass", {})
    crs     = data.get("crs", {})
    gr      = crs.get("gr", {})

    def pct(actual, target_key):
        t = targets.get(target_key)
        if not t: return None
        return round(100 * actual / t, 1)

    # 1. Rolling combined
    total_roll = sum(m.get("day_total", 0) for m in rolling.values())
    p = pct(total_roll, "rolling_day")
    if p is not None:
        word = "achieved" if p >= 100 else ("near" if p >= 95 else "below")
        notes.append(f"Combined rolling {word} {p}% of daily target at {total_roll:.1f} MT.")

    # 2. Best-performing mill
    if rolling:
        best = max(rolling.items(), key=lambda x: (x[1].get("day_total", 0) /
                   (x[1].get("day_target") or 1)))
        bname, bm = best
        bp = pct(bm.get("day_total", 0), f"rolling_day")
        if bm.get("yield"):
            notes.append(f"{bname} leading with yield {bm['yield']:.2f}% "
                         f"and {bm.get('day_total', 0):.1f} MT day production.")

    # 3. Annealing
    ann_prod = ann.get("day_prod", 0)
    p = pct(ann_prod, "ann_day")
    if p is not None:
        word = "exceeded" if p >= 100 else ("met" if p >= 95 else "missed")
        notes.append(f"Annealing {word} target at {p}% — {ann_prod:.1f} MT "
                     f"({ann.get('charges', 0)} charge(s) processed).")

    # 4. GR summary
    tgr = gr.get("total", {}).get("yesterday", 0)
    tgr_cumm = gr.get("total", {}).get("cumm", 0)
    p = pct(tgr, "total_gr_day")
    if p is not None:
        notes.append(f"Total GR stands at {tgr} T ({p}% of daily target); "
                     f"MTD cumulative {tgr_cumm} T.")

    # 5. CRS / hold / WIP concern or positive note
    hold = crs.get("hold", {}).get("value", 0)
    skp_wip = crs.get("skp_wip", 0)
    no_plan = crs.get("no_plan", 0)
    crs_total = crs.get("total_at_crs", 0)

    critical_count = sum(1 for a in alerts if a["priority"] == "critical")
    if hold > 50:
        notes.append(f"Hold material at {hold:.0f} MT — requires immediate disposition. "
                     f"Total CRS inventory {crs_total:.0f} MT.")
    elif skp_wip > 150:
        notes.append(f"Skinpass WIP elevated at {skp_wip:.0f} MT; "
                     f"No-plan material {no_plan:.0f} MT. "
                     f"Planning review recommended.")
    elif critical_count == 0:
        notes.append(f"No critical alerts today. "
                     f"CRS total at {crs_total:.0f} MT with normal WIP levels.")

    return notes[:5]


# ===========================================================================
# DASHBOARD: layout (build_html)

# ===========================================================================

# dashboard/layout.py  — CRM Sahibabad Narrow Complex
# Modern dark-navy + white-card style matching the Pickling/HRS + Tata Colors references.
# Fixed 1920×1080, Tata brand colours, IBM Plex fonts.

import datetime, sys, os

# from config import COLORS, rag_color  # merged inline

# ─── colour aliases ────────────────────────────────────────────────────────
P   = COLORS["primary"]      # #005CB9
S   = COLORS["secondary"]    # #0087E0
G   = COLORS["success"]      # #00A651
W   = COLORS["warning"]      # #F5A623
R   = COLORS["critical"]     # #D0021B
MUT = COLORS["muted"]        # #6B7280
NAV = "#0F2A5C"              # dark navy (header / section headers)
BG  = "#F0F4FA"              # page background
CB  = "#FFFFFF"              # card background
BRD = "#DDE4EE"              # card border
TXT = "#1A2B45"              # primary text

# ─── circular icon badges (reference image style) ──────────────────────────
BADGE_ICONS = {
    "roll":    ("⟳",   P,     "#EBF3FF"),
    "ann":     ("🔥",  W,     "#FFF4E6"),
    "spm":     ("≡",   "#7B2FBE", "#F3EBFF"),
    "gr":      ("▣",   G,     "#E8F7EF"),
    "tube":    ("◎",   "#0F6E56", "#E3F5EE"),
    "oem":     ("□",   S,     "#E8F3FF"),
    "total_gr":("✦",   P,     "#EBF3FF"),
    "factory": ("⚙",   P,     "#EBF3FF"),
    "flame":   ("🌡",  W,     "#FFF4E6"),
    "target":  ("◎",   G,     "#E8F7EF"),
    "warn":    ("⚠",   R,     "#FFEBEB"),
    "truck":   ("🚚",  "#0F6E56", "#E3F5EE"),
    "chart":   ("📊",  S,     "#E8F3FF"),
}

def _badge(key, size=38):
    ico, fg, bg = BADGE_ICONS.get(key, ("●", P, "#EBF3FF"))
    return (f'<div style="width:{size}px;height:{size}px;border-radius:50%;'
            f'background:{bg};display:flex;align-items:center;justify-content:center;'
            f'font-size:{int(size*0.45)}px;color:{fg};flex-shrink:0">{ico}</div>')

def _fmt(v, d=1):
    if v is None: return "—"
    if isinstance(v, float): return f"{v:.{d}f}"
    return str(v)

# ─── achievement bar (like in references — % shown inside the bar) ─────────
def _achv_bar(pct, color, height=18):
    if pct is None: pct = 0
    pct_c = min(pct, 100)
    label = f"{pct:.1f}%"
    return f"""
    <div style="height:{height}px;background:#E2EAF4;border-radius:{height//2}px;
                overflow:hidden;position:relative;margin-top:6px">
      <div style="height:100%;width:{pct_c:.1f}%;background:{color};border-radius:{height//2}px;
                  display:flex;align-items:center;justify-content:flex-end;padding-right:8px;
                  min-width:48px;transition:width .5s">
        <span style="font-size:11px;font-weight:800;color:#fff;white-space:nowrap">{label}</span>
      </div>
    </div>"""

# ─── 6-card KPI strip (matches reference header row) ──────────────────────
def _kpi_strip_card(label, actual, target, badge_key, unit="MT"):
    color, _, pct = rag_color(actual, target)
    pct_str = f"{pct}%" if pct is not None else "—"
    return f"""
    <div style="flex:1;background:{CB};border-radius:12px;padding:14px 16px;
                border:1px solid {BRD};border-top:4px solid {color};min-width:0">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
        {_badge(badge_key, 36)}
        <div>
          <div style="font-size:9px;font-weight:700;letter-spacing:.9px;text-transform:uppercase;color:{MUT}">{label}</div>
          <div style="font-size:22px;font-weight:900;color:{TXT};line-height:1.1;
                      font-family:'IBM Plex Mono',monospace">{_fmt(actual,1)}<span style="font-size:11px;color:{MUT};font-weight:500;margin-left:3px">{unit}</span></div>
        </div>
      </div>
      <div style="display:flex;justify-content:space-between;align-items:center;
                  font-size:10px;color:{MUT}">
        <span>Target {_fmt(target,0)} {unit}</span>
        <span style="font-size:13px;font-weight:900;color:{color}">{pct_str}</span>
      </div>
      {_achv_bar(pct, color, 14)}
    </div>"""

# ─── mini KPI card ──────────────────────────────────────────────────────────
def _mk(label, val, color=None, unit="MT", badge=None):
    color = color or TXT
    b = _badge(badge, 30) if badge else ""
    return f"""
    <div style="background:{BG};border:1px solid {BRD};border-radius:8px;padding:10px 12px;flex:1;min-width:0">
      <div style="display:flex;align-items:center;gap:8px">
        {b}
        <div style="min-width:0">
          <div style="font-size:9px;text-transform:uppercase;letter-spacing:.6px;color:{MUT};margin-bottom:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{label}</div>
          <div style="font-size:17px;font-weight:800;color:{color};font-family:'IBM Plex Mono',monospace;line-height:1">{_fmt(val)}<span style="font-size:9px;color:{MUT};font-weight:500;margin-left:2px">{unit}</span></div>
        </div>
      </div>
    </div>"""

# ─── horizontal bar row (GR performance style) ─────────────────────────────
def _hbar_row(label, actual, target, color, unit="T"):
    pct = min(100 * actual / target, 100) if target else 0
    bar_color, _, _ = rag_color(actual, target)
    return f"""
    <div style="margin-bottom:7px">
      <div style="display:flex;justify-content:space-between;font-size:10px;margin-bottom:4px">
        <span style="color:{MUT};font-weight:700;text-transform:uppercase;letter-spacing:.4px">{label}</span>
        <span style="font-family:'IBM Plex Mono',monospace;font-weight:800;color:{TXT}">{_fmt(actual,0)} {unit}</span>
      </div>
      <div style="height:20px;background:#E2EAF4;border-radius:10px;overflow:hidden">
        <div style="height:100%;width:{pct:.1f}%;background:{color};border-radius:10px;
                    display:flex;align-items:center;padding:0 10px;min-width:56px">
          <span style="font-size:10px;font-weight:800;color:#fff">{round(pct)}%</span>
        </div>
      </div>
      <div style="font-size:9px;color:{MUT};text-align:right;margin-top:2px">Target {_fmt(target,0)} {unit}</div>
    </div>"""

# ─── bullet bar (mill metrics) ──────────────────────────────────────────────
def _bullet(label, actual, target, unit=""):
    color, _, pct = rag_color(actual, target)
    pct_c = min(pct or 0, 100)
    return f"""
    <div style="margin-bottom:8px">
      <div style="display:flex;justify-content:space-between;font-size:10px;margin-bottom:3px">
        <span style="color:{MUT};font-weight:600;text-transform:uppercase;letter-spacing:.4px">{label}</span>
        <span style="font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:800;color:{TXT}">{_fmt(actual)}{unit}</span>
      </div>
      <div style="position:relative;height:12px;background:#E2EAF4;border-radius:6px;overflow:visible">
        <div style="position:absolute;top:1px;left:0;height:10px;width:{pct_c:.1f}%;
                    background:{color};border-radius:5px"></div>
        <div style="position:absolute;top:-4px;right:0;width:2px;height:20px;
                    background:{TXT}40;border-radius:1px"></div>
      </div>
      <div style="font-size:9px;color:{color};text-align:right;margin-top:1px;font-weight:700">
        {(str(pct)+'% of '+_fmt(target,0)+unit) if pct is not None else '—'}
      </div>
    </div>"""

# ─── section header (dark navy bar like reference) ─────────────────────────
def _sh(title, icon_key=None, extra=""):
    ico = _badge(icon_key, 28) if icon_key else ""
    return f"""
    <div style="background:{NAV};color:#fff;border-radius:8px 8px 0 0;
                padding:8px 14px;display:flex;align-items:center;gap:10px">
      {ico}
      <span style="font-size:11px;font-weight:800;letter-spacing:.8px;text-transform:uppercase">{title}</span>
      {extra}
    </div>"""

# ─── mill panel ─────────────────────────────────────────────────────────────
def _mill(name, m, targets, accent):
    if not m:
        return f'<div style="background:{CB};border-radius:10px;border:1px solid {BRD};padding:14px;color:{MUT}">No data</div>'
    day = m.get("day_total", 0)
    tgt = m.get("day_target") or 100
    color, _, pct = rag_color(day, tgt)
    ytgt = targets.get("crm06_yield_mtd" if "6" in name or "7" in name else "crm04_yield_mtd", 99)
    utgt = targets.get("crm06_util_mtd"  if "6" in name or "7" in name else "crm04_util_mtd",  80)
    badge_str = f'<span style="margin-left:auto;font-size:12px;font-weight:900;background:rgba(255,255,255,.2);padding:3px 12px;border-radius:12px">{pct}%</span>' if pct else ""
    return f"""
    <div style="background:{CB};border-radius:10px;border:1px solid {BRD};
                overflow:hidden;height:100%">
      <div style="background:{accent};color:#fff;padding:10px 14px;
                  display:flex;align-items:center;gap:10px">
        {_badge("factory", 32)}
        <div>
          <div style="font-size:13px;font-weight:900;letter-spacing:.3px">{name}</div>
          <div style="font-size:10px;opacity:.8">Day Production Summary</div>
        </div>
        {badge_str}
      </div>
      <div style="padding:12px 14px">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px">
          {_mk("Day Roll",  m.get("day_roll",0),  P,   "MT")}
          {_mk("Day R/R",   m.get("day_rr",0),    S,   "MT")}
          {_mk("Cumulative",m.get("cumm_roll",0), MUT, "MT")}
          {_mk("Coils",     int(m.get("coils",0)) if m.get("coils") else "—", TXT, "")}
        </div>
        {_bullet("Day Production", day, tgt, " MT")}
        {_bullet("Yield %",   m.get("yield",0),     ytgt, "%")}
        {_bullet("Utilisation %", m.get("day_util",0), utgt, "%")}
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;
                    padding-top:10px;border-top:1px solid {BRD};margin-top:8px">
          <div style="font-size:9px;color:{MUT}">Avg Gauge<div style="font-size:14px;font-weight:800;color:{TXT};margin-top:1px">{m.get('avg_gauge','—')} mm</div></div>
          <div style="font-size:9px;color:{MUT}">Avg Width<div style="font-size:14px;font-weight:800;color:{TXT};margin-top:1px">{m.get('avg_width','—')} mm</div></div>
          <div style="font-size:9px;color:{MUT}">TPOH<div style="font-size:14px;font-weight:800;color:{TXT};margin-top:1px">{m.get('tpoh','—')} T/h</div></div>
          <div style="font-size:9px;color:{MUT}">Delay<div style="font-size:14px;font-weight:800;color:{R};margin-top:1px">{m.get('delay_hrs',0):.1f} hr</div></div>
        </div>
        <div style="margin-top:8px;padding-top:8px;border-top:1px solid {BRD};
                    display:flex;justify-content:space-between;font-size:9px;color:{MUT}">
          <span>Cumm Yield <b style="color:{G}">{m.get('cumm_yield','—')}%</b></span>
          <span>Util TillDate <b style="color:{S}">{m.get('util_tilldate','—')}%</b></span>
        </div>
      </div>
    </div>"""

# ─── line-wise production table (reference image style) ────────────────────
def _linewise_table(rolling, targets):
    rows = ""
    colors = [P, "#7B2FBE", G, W, S]
    for i, (mn, m) in enumerate(list(rolling.items())[:5]):
        day = m.get("day_total", 0)
        tgt = m.get("day_target") or 100
        color, _, pct = rag_color(day, tgt)
        pct_c = min(pct or 0, 100)
        rows += f"""
        <tr style="border-bottom:1px solid {BRD}">
          <td style="padding:7px 10px;font-size:11px;font-weight:700;color:{TXT}">{mn}</td>
          <td style="padding:7px 10px;font-size:11px;font-family:'IBM Plex Mono',monospace;font-weight:700;color:{TXT}">{_fmt(day,1)} MT</td>
          <td style="padding:7px 10px;font-size:11px;font-family:'IBM Plex Mono',monospace;color:{MUT}">{_fmt(m.get('cumm_roll',0),0)} MT</td>
          <td style="padding:7px 10px;width:120px">
            <div style="height:14px;background:#E2EAF4;border-radius:7px;overflow:hidden">
              <div style="height:100%;width:{pct_c:.0f}%;background:{color};border-radius:7px;
                          display:flex;align-items:center;padding:0 6px;min-width:36px">
                <span style="font-size:9px;color:#fff;font-weight:800">{pct or 0:.0f}%</span>
              </div>
            </div>
          </td>
        </tr>"""
    return f"""
    <table style="width:100%;border-collapse:collapse">
      <thead>
        <tr style="background:{BG}">
          <th style="padding:7px 10px;font-size:9px;color:{MUT};text-align:left;font-weight:700;letter-spacing:.6px;text-transform:uppercase;border-bottom:2px solid {BRD}">Mill</th>
          <th style="padding:7px 10px;font-size:9px;color:{MUT};text-align:left;font-weight:700;letter-spacing:.6px;text-transform:uppercase;border-bottom:2px solid {BRD}">Day</th>
          <th style="padding:7px 10px;font-size:9px;color:{MUT};text-align:left;font-weight:700;letter-spacing:.6px;text-transform:uppercase;border-bottom:2px solid {BRD}">Cumm</th>
          <th style="padding:7px 10px;font-size:9px;color:{MUT};text-align:left;font-weight:700;letter-spacing:.6px;text-transform:uppercase;border-bottom:2px solid {BRD}">Achievement</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>"""

# ─── finishing/CRS inventory rows ──────────────────────────────────────────
def _fin_row(label, value, threshold=None, unit="MT", note=""):
    v = value if isinstance(value, (int,float)) else 0
    color = G
    if threshold and v > threshold: color = R
    elif threshold and v > threshold * 0.8: color = W
    note_html = f'<span style="font-size:9px;color:{MUT};margin-left:6px">{note}</span>' if note else ""
    return f"""
    <div style="display:flex;justify-content:space-between;align-items:center;
                padding:8px 12px;border-bottom:1px solid {BRD}">
      <span style="font-size:10px;font-weight:700;color:{TXT};text-transform:uppercase;letter-spacing:.4px">{label}</span>
      <div style="display:flex;align-items:center">
        <span style="font-size:15px;font-weight:900;color:{color};
                     font-family:'IBM Plex Mono',monospace">{_fmt(v,0)} <span style="font-size:9px;color:{MUT};font-weight:500">{unit}</span></span>
        {note_html}
      </div>
    </div>"""

# ─── alert cards ────────────────────────────────────────────────────────────
def _alert_card(a):
    pri = a["priority"]
    color = {  "critical": R, "warning": W, "info": S }.get(pri, MUT)
    ico   = {  "critical": "⚠", "warning": "⚠", "info": "ℹ" }.get(pri, "●")
    bg    = f"{color}0F"
    brd   = f"{color}35"
    return f"""
    <div style="background:{bg};border:1px solid {brd};border-left:4px solid {color};
                border-radius:8px;padding:8px 12px;margin-bottom:6px">
      <div style="font-size:11px;font-weight:800;color:{color};display:flex;align-items:center;gap:5px">
        <span>{ico}</span>{a['title']}
      </div>
      <div style="font-size:10px;color:{MUT};margin-top:2px;margin-left:16px">{a.get('detail','')}</div>
    </div>"""

# ════════════════════════════════════════════════════════════════════════════
#  MAIN build_html
# ════════════════════════════════════════════════════════════════════════════
def build_html(data, targets, alerts, notes, report_date=None, report_time=None):
    rolling   = data.get("rolling", {})
    ann       = data.get("annealing", {}).get("ann02", {})
    spm       = data.get("annealing", {}).get("skin_pass", {})
    crs       = data.get("crs", {})
    gr        = crs.get("gr", {})

    today     = report_date or datetime.date.today().strftime("%d %B %Y")
    now       = report_time or datetime.datetime.now().strftime("%H:%M")

    total_roll = sum(m.get("day_total",0) for m in rolling.values())
    ann_prod   = ann.get("day_prod", 0)
    spm_prod   = spm.get("day_prod", 0)
    tgr_today  = gr.get("total",{}).get("yesterday", 0)
    tgr_cumm   = gr.get("total",{}).get("cumm", 0)
    tube_today = gr.get("tube",{}).get("yesterday", 0)
    tube_cumm  = gr.get("tube",{}).get("cumm", 0)
    oem_today  = gr.get("oem",{}).get("yesterday", 0)
    oem_cumm   = gr.get("oem",{}).get("cumm", 0)

    # Donut SVG (annealing base mix)
    new_b = ann.get("prod_new", 0) or 0.001
    old_b = ann.get("prod_old", 0) or 0
    tot_b = new_b + old_b or 1
    np_   = new_b / tot_b
    op_   = old_b / tot_b
    r_    = 45
    cxy   = 55
    def arc(frac, color, offset):
        circ = 2 * 3.14159 * r_
        dash = circ * frac
        gap  = circ * (1 - frac)
        return f'<circle cx="{cxy}" cy="{cxy}" r="{r_}" fill="none" stroke="{color}" stroke-width="18" stroke-dasharray="{dash:.2f} {gap:.2f}" stroke-dashoffset="{offset:.2f}" transform="rotate(-90 {cxy} {cxy})"/>'
    circ_ = 2 * 3.14159 * r_
    donut_svg = f"""
    <svg width="110" height="110" viewBox="0 0 110 110" style="overflow:visible">
      {arc(np_, G,     0)}
      {arc(op_, BRD,   -circ_ * np_)}
      <text x="{cxy}" y="{cxy+4}" text-anchor="middle" font-size="11" font-weight="900" fill="{TXT}" font-family="IBM Plex Mono,monospace">{round(new_b,0):.0f}</text>
    </svg>"""

    # Mill panels
    mill_items = list(rolling.items())
    mill_accents = [P, "#7B2FBE"]
    mill_html = ""
    for i, (mn, mv) in enumerate(mill_items[:2]):
        mv["day_target"] = mv.get("day_target") or (targets.get("rolling_day",200)/max(len(mill_items),1))
        mill_html += f'<div style="flex:1;min-width:0">{_mill(mn, mv, targets, mill_accents[i%2])}</div>'

    # Alerts
    alert_html = "".join(_alert_card(a) for a in alerts) or f'<div style="color:{G};font-size:11px;padding:8px 0;font-weight:700">✓ No critical alerts. Operations normal.</div>'

    # Notes
    notes_html = "".join(f'<div style="padding:6px 0;border-bottom:1px solid {BRD};font-size:11px;color:{TXT};line-height:1.5;display:flex;gap:8px"><span style="color:{G};font-weight:900;flex-shrink:0">✓</span>{n}</div>' for n in notes) or ""

    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700;800&family=IBM+Plex+Sans:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'IBM Plex Sans',system-ui,sans-serif}}
body{{background:{BG};color:{TXT};width:1920px;min-height:1080px;padding:12px;font-size:12px}}
.card{{background:{CB};border:1px solid {BRD};border-radius:10px;overflow:hidden}}
</style></head><body>

<!-- ══ HEADER — dark navy like references ══ -->
<div style="background:{NAV};color:#fff;border-radius:12px;padding:10px 24px;margin-bottom:10px;
            display:flex;align-items:center;justify-content:space-between;min-height:66px">
  <div style="display:flex;align-items:center;gap:20px">
    <div style="background:#fff;border-radius:8px;width:52px;height:52px;
                display:flex;align-items:center;justify-content:center;
                font-size:22px;font-weight:900;color:{NAV}">T</div>
    <div>
      <div style="font-size:11px;opacity:.7;letter-spacing:.5px">TATA STEEL · SAHIBABAD</div>
      <div style="font-size:20px;font-weight:900;letter-spacing:.3px;line-height:1.2">CRM NARROW COMPLEX — DAILY OPERATIONS DASHBOARD</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:8px">
    <div style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);
                border-radius:8px;padding:8px 18px;text-align:center">
      <div style="font-size:10px;opacity:.7;letter-spacing:.5px">DATE</div>
      <div style="font-size:17px;font-weight:800;font-family:'IBM Plex Mono',monospace">{today}</div>
    </div>
    <div style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);
                border-radius:8px;padding:8px 18px;text-align:center">
      <div style="font-size:10px;opacity:.7;letter-spacing:.5px">REPORT TIME</div>
      <div style="font-size:17px;font-weight:800;font-family:'IBM Plex Mono',monospace">{now} IST</div>
    </div>
    <div style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);
                border-radius:8px;padding:8px 18px;text-align:center">
      <div style="font-size:10px;opacity:.7;letter-spacing:.5px">ROLLING</div>
      <div style="font-size:17px;font-weight:900;font-family:'IBM Plex Mono',monospace;color:#38BDF8">{_fmt(total_roll)} <span style="font-size:10px;font-weight:400;opacity:.8">MT</span></div>
    </div>
    <div style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);
                border-radius:8px;padding:8px 18px;text-align:center">
      <div style="font-size:10px;opacity:.7;letter-spacing:.5px">ANNEALING</div>
      <div style="font-size:17px;font-weight:900;font-family:'IBM Plex Mono',monospace;color:{W}">{_fmt(ann_prod)} <span style="font-size:10px;font-weight:400;opacity:.8">MT</span></div>
    </div>
    <div style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);
                border-radius:8px;padding:8px 18px;text-align:center">
      <div style="font-size:10px;opacity:.7;letter-spacing:.5px">SKIN PASS</div>
      <div style="font-size:17px;font-weight:900;font-family:'IBM Plex Mono',monospace;color:#A78BFA">{_fmt(spm_prod)} <span style="font-size:10px;font-weight:400;opacity:.8">MT</span></div>
    </div>
    <div style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);
                border-radius:8px;padding:8px 18px;text-align:center">
      <div style="font-size:10px;opacity:.7;letter-spacing:.5px">GR TODAY</div>
      <div style="font-size:17px;font-weight:900;font-family:'IBM Plex Mono',monospace;color:#34D399">{_fmt(tgr_today,0)} <span style="font-size:10px;font-weight:400;opacity:.8">T</span></div>
    </div>
  </div>
</div>

<!-- ══ KPI STRIP ══ -->
<div style="display:flex;gap:8px;margin-bottom:8px">
  {_kpi_strip_card("Rolling",        total_roll, targets.get("rolling_day"),  "roll")}
  {_kpi_strip_card("Annealing",      ann_prod,   targets.get("ann_day"),      "flame")}
  {_kpi_strip_card("2HI Skin Pass",  spm_prod,   targets.get("spm_day"),      "spm")}
  {_kpi_strip_card("Tube GR",   tube_today, targets.get("tube_gr_day"), "tube",  "T")}
  {_kpi_strip_card("OEM GR",    oem_today,  targets.get("oem_gr_day"),  "oem",   "T")}
  {_kpi_strip_card("Total GR",  tgr_today,  targets.get("total_gr_day"),"total_gr","T")}
</div>

<!-- ══ ROW 1: TODAY HIGHLIGHTS | MILLS | LINE-WISE ══ -->
<div style="display:flex;gap:8px;margin-bottom:8px;align-items:stretch">

  <!-- TODAY HIGHLIGHTS — large numbers, no charts (per PDF spec) -->
  <div class="card" style="width:195px;flex-shrink:0">
    {_sh("Today's Highlights","chart")}
    <div style="padding:10px;display:flex;flex-direction:column;gap:6px">
      <div style="display:flex;align-items:center;gap:10px;padding:8px;background:{BG};border-radius:8px">
        {_badge("roll",34)}
        <div><div style="font-size:8px;color:{MUT};text-transform:uppercase;letter-spacing:.5px">Rolling</div>
          <div style="font-size:20px;font-weight:900;color:{P};font-family:'IBM Plex Mono',monospace">{_fmt(total_roll,1)} <span style="font-size:9px;color:{MUT}">MT</span></div></div>
      </div>
      <div style="display:flex;align-items:center;gap:10px;padding:8px;background:{BG};border-radius:8px">
        {_badge("flame",34)}
        <div><div style="font-size:8px;color:{MUT};text-transform:uppercase;letter-spacing:.5px">Annealing</div>
          <div style="font-size:20px;font-weight:900;color:{W};font-family:'IBM Plex Mono',monospace">{_fmt(ann_prod,1)} <span style="font-size:9px;color:{MUT}">MT</span></div></div>
      </div>
      <div style="display:flex;align-items:center;gap:10px;padding:8px;background:{BG};border-radius:8px">
        {_badge("spm",34)}
        <div><div style="font-size:8px;color:{MUT};text-transform:uppercase;letter-spacing:.5px">2HI SKP</div>
          <div style="font-size:20px;font-weight:900;color:#7B2FBE;font-family:'IBM Plex Mono',monospace">{_fmt(spm_prod,1)} <span style="font-size:9px;color:{MUT}">MT</span></div></div>
      </div>
      <div style="display:flex;align-items:center;gap:10px;padding:8px;background:{BG};border-radius:8px">
        {_badge("gr",34)}
        <div><div style="font-size:8px;color:{MUT};text-transform:uppercase;letter-spacing:.5px">GR Today</div>
          <div style="font-size:20px;font-weight:900;color:{G};font-family:'IBM Plex Mono',monospace">{_fmt(tgr_today,0)} <span style="font-size:9px;color:{MUT}">T</span></div></div>
      </div>
      <div style="display:flex;align-items:center;gap:10px;padding:8px;background:{BG};border-radius:8px">
        {_badge("warn",34)}
        <div><div style="font-size:8px;color:{MUT};text-transform:uppercase;letter-spacing:.5px">Hold</div>
          <div style="font-size:20px;font-weight:900;color:{R};font-family:'IBM Plex Mono',monospace">{_fmt(crs.get('hold',{}).get('value',0),0)} <span style="font-size:9px;color:{MUT}">MT</span></div></div>
      </div>
    </div>
  </div>

  <!-- MILL PANELS -->
  <div style="flex:1;display:flex;gap:10px;min-width:0">
    {mill_html}
  </div>

  <!-- LINE-WISE PRODUCTION TABLE (right column, reference style) -->
  <div class="card" style="width:380px;flex-shrink:0">
    {_sh("Mill Production Summary","bars")}
    <div style="padding:10px">
      {_linewise_table(rolling, targets)}
    </div>
    <div style="border-top:1px solid {BRD};padding:10px;display:grid;grid-template-columns:1fr 1fr;gap:8px">
      {_mk("GR Cumm Tube",   tube_cumm,  "#0F6E56", "T")}
      {_mk("GR Cumm OEM",    oem_cumm,   P,          "T")}
      {_mk("GR Cumm Total",  tgr_cumm,   S,          "T")}
      {_mk("CRCA Slitting",  crs.get('crca_slitting', dict()).get('yesterday',0), MUT, "MT")}
    </div>
  </div>
</div>

<!-- ══ ROW 2: ANNEALING | 2HI SKP ══ -->
<div style="display:flex;gap:8px;margin-bottom:8px">
  <!-- ANNEALING -->
  <div class="card" style="flex:1">
    {_sh("Annealing Operations","flame")}
    <div style="padding:12px;display:flex;gap:12px;align-items:flex-start">
      <div style="flex:1">
        <div style="display:flex;gap:8px;margin-bottom:7px;flex-wrap:wrap">
          {_mk("New Base", ann.get("prod_new",0), G,   "MT","ann")}
          {_mk("Old Base", ann.get("prod_old",0), MUT, "MT")}
          {_mk("Total",    ann_prod,              P,   "MT")}
          {_mk("Charges",  ann.get("charges",0),  S,   "")}
          {_mk("Water",    ann.get("water",0),    "#0F6E56","m³")}
          {_mk("LNG",      ann.get("lng_nm3",0),  "#7B2FBE","Nm³")}
        </div>
        {_bullet("Day Production", ann_prod, targets.get("ann_day"), " MT")}
        <div style="display:flex;gap:4px;margin-top:8px;font-size:10px;color:{MUT}">
          <span>New charges: <b style="color:{TXT}">{int(ann.get('chg_new',0))}</b></span>
          <span style="margin-left:12px">Old charges: <b style="color:{TXT}">{int(ann.get('chg_old',0))}</b></span>
          {f'<span style="margin-left:12px;color:{R};font-weight:700">{ann.get("delay","")}</span>' if ann.get("delay") else ""}
        </div>
      </div>
      <!-- DONUT -->
      <div style="flex-shrink:0;text-align:center">
        <div style="font-size:8px;text-transform:uppercase;letter-spacing:.7px;color:{MUT};margin-bottom:4px">Base Mix</div>
        {donut_svg}
        <div style="font-size:9px;margin-top:4px">
          <span style="display:inline-flex;align-items:center;gap:4px"><span style="width:9px;height:9px;background:{G};border-radius:2px;display:inline-block"></span>New {_fmt(new_b,1)}</span>&nbsp;
          <span style="display:inline-flex;align-items:center;gap:4px"><span style="width:9px;height:9px;background:{BRD};border-radius:2px;display:inline-block"></span>Old {_fmt(old_b,1)}</span>
        </div>
      </div>
    </div>
  </div>

  <!-- 2HI SKIN PASS -->
  <div class="card" style="width:380px;flex-shrink:0">
    {_sh("2HI Skin Pass","spm")}
    <div style="padding:12px">
      <div style="display:flex;gap:8px;margin-bottom:7px">
        {_mk("2HI Prod",   spm.get("hi_prod",0),   "#7B2FBE","MT")}
        {_mk("ID C
