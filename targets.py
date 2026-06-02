# extractors/targets.py
# Reads Target.xlsx  (KPI | Day Target | MTD Target)
# Returns a flat dict:  targets["rolling_day"] = 200, etc.

import openpyxl
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from config import TARGET_ROWS

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
if __name__ == "__main__":
    import json
    t = parse_targets("/mnt/user-data/uploads/Target__1_.xlsx")
    print(json.dumps(t, indent=2))
