# dashboard/layout.py  — CRM Sahibabad Narrow Complex
# Modern dark-navy + white-card style matching the Pickling/HRS + Tata Colors references.
# Fixed 1920×1080, Tata brand colours, IBM Plex fonts.

import datetime, sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from config import COLORS, rag_color

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
        {_mk("ID Change",  spm.get("id_change",0),  S,         "MT")}
        {_mk("HROP SKP",   spm.get("hrop",0),       W,         "MT")}
      </div>
      {_bullet("Day Production",     spm_prod, targets.get("spm_day"),  " MT")}
      {_hbar_row("Achievement vs Target", spm_prod, targets.get("spm_day") or 47, "#7B2FBE", "MT")}
    </div>
  </div>
</div>

<!-- ══ ROW 3: FINISHING SECTION | GR | ALERTS + NOTES ══ -->
<div style="display:flex;gap:10px">

  <!-- FINISHING SECTION (renamed per request) -->
  <div class="card" style="width:370px;flex-shrink:0">
    {_sh("Finishing Section","warn")}
    <div style="padding:0">
      {_fin_row("Total at CRS",   crs.get('total_at_crs',0),              unit="MT")}
      {_fin_row("Skinpass WIP",   crs.get('skp_wip',0),                   threshold=targets.get('skp_wip_max'), unit="MT")}
      {_fin_row("For Slitting",   crs.get('slitting', dict()).get('total',0), unit="MT")}
      {_fin_row("For Packing",    crs.get('packing', dict()).get('total',0),  unit="MT")}
      {_fin_row("No Plan",        crs.get('no_plan',0),                   threshold=30, unit="MT")}
      {_fin_row("HROP Skinpass",  crs.get('hrop_skp',0),                  unit="MT")}
      {_fin_row("Hold Material",  crs.get('hold', dict()).get('value',0),    threshold=targets.get('hold_max') or 50, unit="MT", note=crs.get('hold', dict()).get('note',''))}
      {_fin_row("CRCA Slitting",  crs.get('crca_slitting', dict()).get('yesterday',0), unit="MT")}
    </div>
    <div style="padding:8px 12px;background:{BG};display:flex;gap:12px;font-size:9px;color:{MUT};border-top:1px solid {BRD}">
      <span><span style="color:{G};font-weight:900">●</span> Normal</span>
      <span><span style="color:{W};font-weight:900">●</span> Attention</span>
      <span><span style="color:{R};font-weight:900">●</span> Critical</span>
    </div>
    <!-- Material readiness sub-grid -->
    <div style="padding:10px;border-top:1px solid {BRD}">
      <div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:{MUT};margin-bottom:8px">Material Readiness</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">
        {_mk("Slit OEM",   crs.get('slitting', dict()).get('oem',0),    P,         "MT")}
        {_mk("Slit Tube",  crs.get('slitting', dict()).get('tube',0),   G,         "MT")}
        {_mk("Pack OEM",   crs.get('packing', dict()).get('oem',0),     S,         "MT")}
        {_mk("Pack Tube",  crs.get('packing', dict()).get('tube',0),    "#0F6E56", "MT")}
      </div>
    </div>
  </div>

  <!-- GR PERFORMANCE -->
  <div class="card" style="flex:1">
    {_sh("Finishing / GR Performance","truck")}
    <div style="padding:12px;display:flex;gap:12px">
      <div style="flex:1">
        <div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:{MUT};margin-bottom:7px">Day vs Target</div>
        {_hbar_row("Tube GR",  tube_today, targets.get("tube_gr_day"),  "#0F6E56")}
        {_hbar_row("OEM GR",   oem_today,  targets.get("oem_gr_day"),   P)}
        {_hbar_row("Total GR", tgr_today,  targets.get("total_gr_day"), S)}
      </div>
      <div style="width:200px;flex-shrink:0">
        <div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:{MUT};margin-bottom:7px">Cumulative (MTD)</div>
        {_mk("Tube Cumm",  tube_cumm,  "#0F6E56","T")}
        <div style="height:7px"></div>
        {_mk("OEM Cumm",   oem_cumm,   P,         "T")}
        <div style="height:7px"></div>
        {_mk("Total Cumm", tgr_cumm,   S,         "T")}
      </div>
    </div>
    <!-- Slitting packing readiness bars -->
    <div style="padding:12px;border-top:1px solid {BRD}">
      <div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:{MUT};margin-bottom:7px">Slitting Material — OEM vs Tube</div>
      <div style="display:flex;gap:4px;height:28px;border-radius:8px;overflow:hidden;margin-bottom:6px">
        <div style="flex:{crs.get('slitting', dict()).get('oem',1)};background:{P};display:flex;align-items:center;justify-content:center">
          <span style="font-size:10px;color:#fff;font-weight:800">OEM {_fmt(crs.get('slitting', dict()).get('oem',0),0)}</span>
        </div>
        <div style="flex:{crs.get('slitting', dict()).get('tube',1)};background:{G};display:flex;align-items:center;justify-content:center">
          <span style="font-size:10px;color:#fff;font-weight:800">Tube {_fmt(crs.get('slitting', dict()).get('tube',0),0)}</span>
        </div>
      </div>
      <div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:{MUT};margin-bottom:8px;margin-top:10px">Packing Material — OEM vs Tube</div>
      <div style="display:flex;gap:4px;height:28px;border-radius:8px;overflow:hidden">
        <div style="flex:{crs.get('packing', dict()).get('oem',1)};background:{S};display:flex;align-items:center;justify-content:center">
          <span style="font-size:10px;color:#fff;font-weight:800">OEM {_fmt(crs.get('packing', dict()).get('oem',0),0)}</span>
        </div>
        <div style="flex:{crs.get('packing', dict()).get('tube',1)};background:"#0F6E56";display:flex;align-items:center;justify-content:center">
          <span style="font-size:10px;color:#fff;font-weight:800">Tube {_fmt(crs.get('packing', dict()).get('tube',0),0)}</span>
        </div>
      </div>
    </div>
  </div>

  <!-- ALERTS + NOTES -->
  <div style="width:330px;flex-shrink:0;display:flex;flex-direction:column;gap:10px">
    <div class="card" style="flex:1">
      {_sh("Critical Alerts","warn")}
      <div style="padding:10px">{alert_html}</div>
    </div>
    <div class="card">
      {_sh("Management Notes","target")}
      <div style="padding:10px">{notes_html or f'<div style="color:{MUT};font-size:11px">Generating...</div>'}</div>
    </div>
  </div>
</div>

</body></html>"""
