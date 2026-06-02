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
    sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
    from generate_dashboard import xlsx_to_tsv, parse_anneal_mis

    tsv, sheets = xlsx_to_tsv(filepath)
    return parse_anneal_mis(tsv, report_day=report_day)


if __name__ == "__main__":
    import json
    d = parse_annealing("/mnt/user-data/uploads/20260530_1429598851336786811523554.jpg")
    print(json.dumps(d, indent=2))
