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
if __name__ == "__main__":
    import json
    d = parse_crs("/mnt/user-data/uploads/CRS_MIS.xlsx")
    print(json.dumps(d, indent=2))
