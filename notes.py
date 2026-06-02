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
