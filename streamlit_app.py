"""
CRM Sahibabad Narrow Complex — Daily Operations Dashboard
v2.1 — SINGLE FILE DEPLOY

Upload exactly 3 files to GitHub (one by one):
  1. streamlit_app.py       ← this file
  2. generate_dashboard.py  ← all parsers + dashboard logic merged in
  3. requirements.txt

No folders needed. Password via Streamlit Secrets: app_password = "..."
"""
import streamlit as st
import datetime, io, os, sys

# Everything lives in generate_dashboard.py — one import covers all
from generate_dashboard import (
    xlsx_to_tsv, parse_mis, parse_anneal_mis, demo_data,
    parse_single_mill_mis, mill_name_from_filename,
    parse_targets, parse_crs,
    generate_alerts, generate_notes,
    build_html,
)

def load_targets_from_secrets():
    """
    Load annual targets from Streamlit Secrets.
    Keys in secrets must be under [targets], e.g.:
        [targets]
        rolling_day    = 200
        ann_day        = 150
        spm_day        = 47
        tube_gr_day    = 167
        oem_gr_day     = 27
        total_gr_day   = 193
        crm04_util_mtd = 80
        crm04_yield_mtd = 99
        crm06_util_mtd = 80
        crm06_yield_mtd = 99
        hold_max       = 50
        skp_wip_max    = 200
    Returns a dict. Falls back to session-state overrides if set.
    """
    base = {}
    try:
        raw = st.secrets.get("targets", {})
        for k, v in raw.items():
            try:    base[k] = float(v)
            except: base[k] = v
    except Exception:
        pass
    # Session-state overrides applied on top
    overrides = st.session_state.get("target_overrides", {})
    base.update(overrides)
    return base


st.set_page_config(
    page_title="CRM Sahibabad — Narrow Complex",
    page_icon="🏭", layout="wide",
    initial_sidebar_state="expanded"
)

# ── Password gate ─────────────────────────────────────────────────────────────
def check_password():
    pw = st.secrets.get("app_password", None)
    if not pw: return True
    if st.session_state.get("auth"): return True
    st.title("🔒 CRM Sahibabad — Narrow Complex")
    entered = st.text_input("Access password", type="password")
    if entered:
        if entered == pw:
            st.session_state["auth"] = True; st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

check_password()

# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.title("CRM Sahibabad")
st.sidebar.caption("Narrow Complex · Daily Operations")

page = st.sidebar.radio("Navigation", [
    "📊 Dashboard", "✅ Data Validation", "🎯 Target Master", "⬇ Exports"
])
st.sidebar.divider()

day = st.sidebar.number_input("Report day (date of month)",
    min_value=1, max_value=31, value=datetime.date.today().day)

st.sidebar.markdown("**Upload MIS files**")

# ── Two separate mill files ────────────────────────────────────────────────
st.sidebar.markdown("*Rolling Mills*")
crm04_file = st.sidebar.file_uploader(
    "CRM04 MIS.xlsx  (mill name read from filename)",
    type=["xlsx"], key="crm04")
crm06_file = st.sidebar.file_uploader(
    "CRM06 MIS.xlsx  (mill name read from filename)",
    type=["xlsx"], key="crm06")

with st.sidebar.expander("⚙ Row-to-date mapping (optional)", expanded=False):
    st.caption(
        "Both files use continuous row numbers (32, 33 … 62). "
        "Enter the row number that equals **Day 1** of this month "
        "so the parser can find the exact day. "
        "Leave 0 to use the last filled row (safe default)."
    )
    crm04_start = st.number_input("CRM04 month-start row", min_value=0, value=0, step=1, key="c4s")
    crm06_start = st.number_input("CRM06 month-start row", min_value=0, value=0, step=1, key="c6s")

ann_file  = st.sidebar.file_uploader("Annealing & 2HI SPM MIS.xlsx", type=["xlsx"], key="ann")
crs_file  = st.sidebar.file_uploader("CRS MIS.xlsx",                  type=["xlsx"], key="crs")
# Target.xlsx uploader removed — targets are stored in Streamlit Secrets (see below)

st.sidebar.divider()
ann_sheet  = st.sidebar.text_input("Annealing sheet/tab", value="", key="asheet",
                                    placeholder="e.g. MAY-26")

run = st.sidebar.button("⚡ Generate Dashboard", type="primary", use_container_width=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def _tsv(uploaded, url, local, sheet=None, label=""):
    src = None
    if uploaded:  src = uploaded.getvalue()
    elif url and url.strip().startswith("http"): src = url.strip()
    elif local and os.path.exists(local): src = local
    if not src: return ""
    try:
        tsv, sheets = xlsx_to_tsv(src, sheet or None)
        if sheets: st.sidebar.caption(f"{label} tabs: {', '.join(sheets)}")
        return tsv
    except Exception as e:
        st.sidebar.error(f"{label}: {e}"); return ""

def _obj(uploaded, url, local, fn, label=""):
    import tempfile
    src = None
    if uploaded:  src = uploaded.getvalue()
    elif url and url.strip().startswith("http"): src = url.strip()
    elif local and os.path.exists(local): src = local
    if not src: return {}
    try:
        if isinstance(src, bytes):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(src); src = tmp.name
        return fn(src)
    except Exception as e:
        st.sidebar.error(f"{label}: {e}"); return {}

# ── Session state ─────────────────────────────────────────────────────────────
for k, v in [("data",{}),("targets",{}),("html",""),
              ("alerts",[]),("notes",[]),("png",None),("pptx",None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Generate ──────────────────────────────────────────────────────────────────
if run:
    with st.spinner("Reading files and building dashboard…"):
        try:
            here = os.path.dirname(os.path.abspath(__file__))
            inp  = lambda n: os.path.join(here, "input", n)

            targets  = load_targets_from_secrets()
            if not targets:
                st.warning("No targets found in Streamlit Secrets. "
                           "Go to 🎯 Target Master to set them for this session.")
            crs      = _obj(crs_file, None, inp("CRS MIS.xlsx"),
                            parse_crs, "CRS")
            ann_tsv  = _tsv(ann_file, None,
                            inp("Annealing and 2HI SPM MIS.xlsx"),
                            ann_sheet, "Annealing")

            # ── Two separate mill files ──────────────────────────────────
            rolling = {}

            def _parse_mill_file(uploaded, url, local_path, start_row):
                """Load one mill file, extract mill name from filename, parse."""
                src = name = None
                if uploaded:
                    src  = uploaded.getvalue()
                    name = mill_name_from_filename(uploaded.name)
                elif url and url.strip().startswith("http"):
                    src  = url.strip()
                    name = mill_name_from_filename(url)
                elif local_path and os.path.exists(local_path):
                    src  = local_path
                    name = mill_name_from_filename(local_path)
                if not src:
                    return {}
                try:
                    tsv, sheets = xlsx_to_tsv(src)
                    month_start = int(start_row) if start_row else None
                    result = parse_single_mill_mis(
                        tsv, name,
                        report_day=int(day),
                        month_start_row=month_start if month_start else None
                    )
                    st.sidebar.caption(
                        f"{name}: row {list(result.values())[0].get('row_date','?')} "
                        f"→ day_total {list(result.values())[0].get('day_total',0):.1f} MT"
                        if result else f"{name}: no data found"
                    )
                    return result
                except Exception as e:
                    st.sidebar.error(f"Mill file error: {e}")
                    return {}

            rolling.update(_parse_mill_file(
                crm04_file, None, inp("CRM04 MIS.xlsx"), crm04_start))
            rolling.update(_parse_mill_file(
                crm06_file, None, inp("CRM06 MIS.xlsx"), crm06_start))

            # Fallback: try the combined MILL MIS.xlsx if neither individual
            # file was provided (backwards compatible)
            if not rolling:
                combined_tsv = _tsv(None, None, inp("MILL MIS.xlsx"), None, "MILL MIS")
                if combined_tsv:
                    rolling = parse_mis(combined_tsv, report_day=int(day))

            # Apply rolling targets per mill
            for mn, m in rolling.items():
                t = targets.get("rolling_day")
                if t: m["day_target"] = t / max(len(rolling), 1)

            annealing = parse_anneal_mis(ann_tsv, report_day=int(day)) if ann_tsv else {}

            data = {
                "date":      crs.get("report_date", ""),
                "rolling":   rolling,
                "annealing": annealing,
                "crs":       crs,
            }
            alerts = generate_alerts(data, targets)
            notes  = generate_notes(data, targets, alerts)
            html   = build_html(data, targets, alerts, notes,
                                report_date=data.get("date"),
                                report_time=datetime.datetime.now().strftime("%H:%M"))
            st.session_state.update(
                data=data, targets=targets, html=html,
                alerts=alerts, notes=notes, png=None, pptx=None
            )
            crit = sum(1 for a in alerts if a["priority"] == "critical")
            mills_found = list(rolling.keys())
            st.success(
                f"✅ Dashboard ready — Mills: {', '.join(mills_found) or 'none'} · "
                f"{len(alerts)} alerts · {crit} critical · {len(notes)} notes"
            )
        except Exception as e:
            st.error(f"Error: {e}")
            import traceback; st.code(traceback.format_exc())

# ══ PAGE: DASHBOARD ══════════════════════════════════════════════════════════
if page == "📊 Dashboard":
    st.title("Daily Operations Dashboard")
    if st.session_state["html"]:
        st.components.v1.html(st.session_state["html"], height=1120, scrolling=True)
        d = st.session_state["data"]
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("⬇ Download HTML", data=st.session_state["html"],
                               file_name=f"Dashboard_{d.get('date','today')}.html",
                               mime="text/html", use_container_width=True)
        with col2:
            st.info("HTML → browser Print → Save as PDF")
    else:
        st.info("Upload MIS files in the sidebar and click **⚡ Generate Dashboard**.")
        if st.button("👁 Preview with demo data"):
            dd  = demo_data()
            tgt = {"rolling_day":200,"ann_day":150,"spm_day":47,
                   "tube_gr_day":167,"oem_gr_day":27,"total_gr_day":193,
                   "crm06_yield_mtd":99,"crm06_util_mtd":80,
                   "crm07_yield_mtd":99,"crm07_util_mtd":80}
            al  = generate_alerts(dd, tgt)
            nt  = generate_notes(dd, tgt, al)
            st.components.v1.html(build_html(dd, tgt, al, nt), height=1120, scrolling=True)

# ══ PAGE: DATA VALIDATION ════════════════════════════════════════════════════
elif page == "✅ Data Validation":
    st.title("Data Validation — Extracted KPIs")
    data    = st.session_state["data"]
    targets = st.session_state["targets"]
    if not data: st.info("Generate a dashboard first."); st.stop()

    import pandas as pd
    st.subheader("Rolling Mills")
    rows = [{"Mill":mn,"Day Total":m.get("day_total"),"Day Target":m.get("day_target"),
             "Day Roll":m.get("day_roll"),"Day R/R":m.get("day_rr"),
             "Yield %":m.get("yield"),"Util %":m.get("day_util"),
             "Avg Gauge":m.get("avg_gauge"),"Avg Width":m.get("avg_width"),
             "Cumm":m.get("cumm_roll"),"Cumm Yield %":m.get("cumm_yield")}
            for mn,m in data.get("rolling",{}).items()]
    if rows: st.dataframe(pd.DataFrame(rows), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Annealing");  st.json(data.get("annealing",{}).get("ann02",{}))
        st.subheader("2HI SKP");    st.json(data.get("annealing",{}).get("skin_pass",{}))
    with c2:
        st.subheader("CRS");        st.json(data.get("crs",{}))

    st.subheader("Alerts")
    ico = {"critical":"🔴","warning":"🟡","info":"🔵"}
    for a in st.session_state["alerts"]:
        st.write(f"{ico.get(a['priority'],'●')} **{a['title']}** — {a.get('detail','')}")
    st.subheader("Management Notes")
    for n in st.session_state["notes"]: st.write(f"✅ {n}")

# ══ PAGE: TARGET MASTER ══════════════════════════════════════════════════════
elif page == "🎯 Target Master":
    st.title("🎯 Target Master")
    st.caption(
        "Annual targets are loaded from **Streamlit Secrets** and stay fixed "
        "for the whole year. Use the form below to temporarily override any "
        "value for this session — overrides are lost when the page refreshes. "
        "To make a permanent change, update the Secrets in Streamlit Cloud settings."
    )

    base_targets = {}
    try:
        raw = st.secrets.get("targets", {})
        for k, v in raw.items():
            try:    base_targets[k] = float(v)
            except: base_targets[k] = v
    except Exception:
        pass

    overrides = st.session_state.get("target_overrides", {})

    # ── Show current effective targets ────────────────────────────────────
    st.subheader("Current targets (Secrets + any session overrides)")

    ALL_TARGETS = {
        # label : (secret_key, step, unit)
        "Rolling Day Target":        ("rolling_day",     10.0,  "MT"),
        "Annealing Day Target":      ("ann_day",          5.0,  "MT"),
        "2HI Skin Pass Day Target":  ("spm_day",          1.0,  "MT"),
        "Tube GR Day Target":        ("tube_gr_day",      5.0,  "T"),
        "OEM GR Day Target":         ("oem_gr_day",       5.0,  "T"),
        "Total GR Day Target":       ("total_gr_day",    10.0,  "T"),
        "CRM04 Utilisation % (MTD)": ("crm04_util_mtd",  1.0,  "%"),
        "CRM04 Yield % (MTD)":       ("crm04_yield_mtd", 0.1,  "%"),
        "CRM06 Utilisation % (MTD)": ("crm06_util_mtd",  1.0,  "%"),
        "CRM06 Yield % (MTD)":       ("crm06_yield_mtd", 0.1,  "%"),
        "Hold Material Max":         ("hold_max",         5.0,  "MT"),
        "Skinpass WIP Max":          ("skp_wip_max",     10.0,  "MT"),
    }

    import pandas as pd
    summary = []
    for label, (key, _, unit) in ALL_TARGETS.items():
        secret_val = base_targets.get(key, "—")
        override_val = overrides.get(key, "—")
        effective = override_val if override_val != "—" else secret_val
        source = "Session override" if override_val != "—" else (
                  "Streamlit Secrets" if secret_val != "—" else "Not set")
        summary.append({
            "KPI": label,
            "From Secrets": secret_val,
            "Session Override": override_val,
            "Effective Value": effective,
            "Unit": unit,
            "Source": source,
        })
    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

    # ── Session override form ─────────────────────────────────────────────
    st.subheader("Set session overrides")
    st.caption("Changes here apply immediately to the next dashboard generation. "
               "They are **not** saved permanently.")

    new_overrides = dict(overrides)
    cols = st.columns(3)
    for i, (label, (key, step, unit)) in enumerate(ALL_TARGETS.items()):
        current = overrides.get(key) or base_targets.get(key) or 0.0
        with cols[i % 3]:
            new_overrides[key] = st.number_input(
                f"{label} ({unit})",
                value=float(current), step=step, key=f"ov_{key}"
            )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ Apply session overrides", type="primary", use_container_width=True):
            st.session_state["target_overrides"] = new_overrides
            st.session_state["targets"] = {**base_targets, **new_overrides}
            st.success("Overrides applied — regenerate dashboard to see changes.")
    with col2:
        if st.button("🔄 Clear all overrides (use Secrets only)", use_container_width=True):
            st.session_state["target_overrides"] = {}
            st.session_state["targets"] = base_targets
            st.success("Overrides cleared.")

    st.divider()
    st.subheader("How to update targets permanently")
    st.markdown("""
1. Go to **share.streamlit.io** → your app → **⋮ menu → Settings → Secrets**
2. Edit the `[targets]` section:
```toml
[targets]
rolling_day     = 200
ann_day         = 150
spm_day         = 47
tube_gr_day     = 167
oem_gr_day      = 27
total_gr_day    = 193
crm04_util_mtd  = 80
crm04_yield_mtd = 99
crm06_util_mtd  = 80
crm06_yield_mtd = 99
hold_max        = 50
skp_wip_max     = 200
```
3. Click **Save** — the app reloads with the new values. No code change, no file upload.
    """)

# ══ PAGE: EXPORTS ════════════════════════════════════════════════════════════
elif page == "⬇ Exports":
    st.title("Exports")
    html     = st.session_state.get("html","")
    data     = st.session_state.get("data",{})
    date_str = data.get("date", datetime.date.today().strftime("%d-%m-%Y"))
    if not html: st.info("Generate a dashboard first."); st.stop()

    st.markdown("### HTML *(always works on cloud)*")
    st.download_button("⬇ Download HTML", data=html,
                       file_name=f"Dashboard_{date_str}.html",
                       mime="text/html", use_container_width=True)
    st.caption("Open in browser → Ctrl+P → Save as PDF")

    st.divider()
    st.markdown("### PNG 1920×1080 *(run locally with Playwright installed)*")
    if st.button("Render PNG", use_container_width=True):
        with st.spinner("Rendering via Chromium…"):
            try:
                import tempfile
                from playwright.sync_api import sync_playwright
                with tempfile.NamedTemporaryFile(suffix=".html",delete=False,
                                                  mode="w",encoding="utf-8") as tf:
                    tf.write(html); tf_path = tf.name
                with sync_playwright() as p:
                    b  = p.chromium.launch()
                    pg = b.new_page(viewport={"width":1920,"height":1080})
                    pg.goto("file://"+tf_path)
                    pg.wait_for_load_state("networkidle")
                    png = pg.screenshot(full_page=False,
                                        clip={"x":0,"y":0,"width":1920,"height":1080})
                    b.close()
                st.session_state["png"] = png
                st.success("PNG ready.")
            except Exception as e:
                st.error(f"PNG render failed: {e} — make sure Playwright is installed locally.")

    if st.session_state.get("png"):
        st.download_button("⬇ Download PNG", data=st.session_state["png"],
                           file_name=f"Dashboard_{date_str}.png",
                           mime="image/png", use_container_width=True)

    st.divider()
    st.markdown("### PowerPoint *(render PNG first)*")
    if st.button("Build PPTX", use_container_width=True):
        if not st.session_state.get("png"):
            st.warning("Render PNG first — the PPTX embeds it.")
        else:
            with st.spinner("Building PPTX…"):
                try:
                    from pptx import Presentation; from pptx.util import Inches
                    import tempfile
                    prs = Presentation()
                    prs.slide_width = Inches(16)
                    prs.slide_height = Inches(9)
                    sl = prs.slides.add_slide(prs.slide_layouts[6])
                    with tempfile.NamedTemporaryFile(suffix=".png",delete=False) as tmp:
                        tmp.write(st.session_state["png"]); tmp_path = tmp.name
                    sl.shapes.add_picture(tmp_path, Inches(0), Inches(0),
                                          Inches(16), Inches(9))
                    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
                    st.session_state["pptx"] = buf.read()
                    st.success("PPTX ready.")
                except Exception as e:
                    st.error(f"PPTX failed: {e}")
    if st.session_state.get("pptx"):
        st.download_button("⬇ Download PPTX", data=st.session_state["pptx"],
                           file_name=f"Dashboard_{date_str}.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                           use_container_width=True)
