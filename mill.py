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


if __name__ == "__main__":
    import json
    # Test against the existing real_mis.tsv
    tsv_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "real_mis.tsv")
    if os.path.exists(tsv_path):
        from generate_dashboard import parse_mis
        with open(tsv_path) as f:
            data = parse_mis(f.read(), report_day=19)
        print(json.dumps(data, indent=2))
    else:
        print("No real_mis.tsv found. Drop MILL MIS.xlsx into input/ and run dashboard_generator.py")
