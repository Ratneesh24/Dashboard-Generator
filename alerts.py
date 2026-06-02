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
