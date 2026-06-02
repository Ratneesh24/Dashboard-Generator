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
