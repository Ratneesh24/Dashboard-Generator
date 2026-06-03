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
    flexible_parse_mis, flexible_parse_anneal,
    get_sheet_headers, MIS_COL, ANN_COL,
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
    "📊 Dashboard", "✅ Data Validation", "🎯 Target Master", "⬇ Exports", "Column Mapper"
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
              ("alerts",[]),("notes",[]),("png",None),("pptx",None),
              ("mill_col_map",{}),("ann_col_map",{})]:
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
                    _mcm = st.session_state.get("mill_col_map", {})
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

            _acm = st.session_state.get("ann_col_map", {})
            annealing = (flexible_parse_anneal(ann_tsv, _acm, report_day=int(day))
                         if ann_tsv else {})

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

# ══ PAGE: COLUMN MAPPER ══════════════════════════════════════════════════════
elif page == "Column Mapper":
    st.title("Column Mapper")
    st.caption(
        "Upload a sheet, see what columns it contains, then map each KPI "
        "to the correct column. The mapping is saved for this session and "
        "used automatically when you Generate the dashboard. "
        "Permanent defaults are stored in Streamlit Secrets."
    )

    # ── Which sheet to map ─────────────────────────────────────────────────
    map_target = st.radio(
        "Which sheet do you want to map?",
        ["Rolling MIS (CRM04 / CRM06)", "Annealing & 2HI SPM MIS"],
        horizontal=True
    )
    is_rolling = map_target.startswith("Rolling")

    # ── Upload for preview ────────────────────────────────────────────────
    preview_file = st.file_uploader(
        f"Upload the {'Rolling' if is_rolling else 'Annealing'} MIS file to inspect",
        type=["xlsx"], key="mapper_file"
    )

    if preview_file:
        raw = preview_file.getvalue()
        headers, preview_rows, sheet_names = get_sheet_headers(raw)

        # Sheet tab selector
        sel_sheet = st.selectbox("Sheet / tab", sheet_names, key="mapper_sheet")
        if sel_sheet != sheet_names[0]:
            headers, preview_rows, _ = get_sheet_headers(raw, sel_sheet)

        # ── Data preview table ─────────────────────────────────────────────
        st.subheader("📋 Sheet preview (first 8 rows)")
        if preview_rows:
            import pandas as pd
            col_labels = [h[1].split("[")[0].strip() for h in headers]
            max_cols = max(len(r) for r in preview_rows)
            # Pad rows to same length
            padded = [r + [""] * (max_cols - len(r)) for r in preview_rows]
            df = pd.DataFrame(padded, columns=col_labels[:max_cols])
            st.dataframe(df, use_container_width=True, height=220)

        st.subheader("🎯 Map columns to KPIs")
        st.caption(
            "For each KPI, choose the column from the dropdown. "
            "The label shows the column letter and sample values from your sheet."
        )

        # Column options for dropdowns
        col_options = ["— not in this file —"] + [h[1] for h in headers]
        col_index_map = {h[1]: h[0] for h in headers}

        def col_pick(label, default_col_idx, key):
            """Show a selectbox defaulting to the column that matches default_col_idx."""
            default_label = next(
                (h[1] for h in headers if h[0] == default_col_idx),
                "— not in this file —"
            )
            choice = st.selectbox(label, col_options,
                                  index=col_options.index(default_label)
                                  if default_label in col_options else 0,
                                  key=key)
            return col_index_map.get(choice)  # None if "not in this file"

        # ── Rolling KPI mapping ───────────────────────────────────────────
        if is_rolling:
            cur = st.session_state.get("mill_col_map", {})
            st.markdown("**Production**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: roll_col = col_pick("Rolling (Day)",  cur.get("roll",  MIS_COL["roll"]),  "m_roll")
            with c2: rr_col   = col_pick("Re-Rolling",     cur.get("rr",    MIS_COL["rr"]),    "m_rr")
            with c3: skp_col  = col_pick("Skin Pass",      cur.get("skp",   MIS_COL["skp"]),   "m_skp")
            with c4: dly_col  = col_pick("Delay (hrs)",    cur.get("delay", MIS_COL["delay"]), "m_delay")

            st.markdown("**Process KPIs**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: yld_col  = col_pick("Yield %",        cur.get("yield",    MIS_COL["yield"]),    "m_yield")
            with c2: utl_col  = col_pick("Utilisation %",  cur.get("day_util", MIS_COL["day_util"]), "m_util")
            with c3: gau_col  = col_pick("Avg Gauge",      cur.get("avg_gauge",MIS_COL["avg_gauge"]),"m_gauge")
            with c4: wid_col  = col_pick("Avg Width",      cur.get("avg_width",MIS_COL["avg_width"]),"m_width")

            st.markdown("**Cumulative**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: cr_col  = col_pick("Cumm Roll",    cur.get("cumm_roll", MIS_COL["cumm_roll"]),  "m_cr")
            with c2: co_col  = col_pick("Cumm Output",  cur.get("cumm_out",  MIS_COL["cumm_out"]),   "m_co")
            with c3: cy_col  = col_pick("Cumm Yield %", cur.get("cumm_yield",MIS_COL["cumm_yield"]), "m_cy")
            with c4: ut_col  = col_pick("Util TillDate",cur.get("util_tilldate",MIS_COL["util_tilldate"]),"m_ut")

            new_map = {
                "roll": roll_col, "rr": rr_col, "skp": skp_col, "delay": dly_col,
                "yield": yld_col, "day_util": utl_col,
                "avg_gauge": gau_col, "avg_width": wid_col,
                "cumm_roll": cr_col, "cumm_out": co_col,
                "cumm_yield": cy_col, "util_tilldate": ut_col,
            }

        # ── Annealing KPI mapping ─────────────────────────────────────────
        else:
            cur = st.session_state.get("ann_col_map", {})
            st.markdown("**Annealing Production**")
            c1, c2, c3 = st.columns(3)
            with c1: pn_col  = col_pick("Prod NEW",       cur.get("prod_new", ANN_COL["prod_new"]), "a_pn")
            with c2: po_col  = col_pick("Prod OLD",       cur.get("prod_old", ANN_COL["prod_old"]), "a_po")
            with c3: tt_col  = col_pick("Total",          cur.get("total",    ANN_COL["total"]),    "a_tt")
            c1, c2, c3 = st.columns(3)
            with c1: cn_col  = col_pick("New Charges",    cur.get("chg_new",  ANN_COL["chg_new"]),  "a_cn")
            with c2: co_col  = col_pick("Old Charges",    cur.get("chg_old",  ANN_COL["chg_old"]),  "a_co")
            with c3: wt_col  = col_pick("Water Cons",     cur.get("water",    ANN_COL["water"]),    "a_wt")

            st.markdown("**2HI / Skin Pass**")
            c1, c2, c3 = st.columns(3)
            with c1: hp_col  = col_pick("2HI Prod",       cur.get("hi_prod",   ANN_COL["hi_prod"]),   "a_hp")
            with c2: ic_col  = col_pick("ID Change",      cur.get("id_change", ANN_COL["id_change"]), "a_ic")
            with c3: hr_col  = col_pick("HROP",           cur.get("hrop",      ANN_COL["hrop"]),      "a_hr")

            new_map = {
                "prod_new": pn_col, "prod_old": po_col, "total": tt_col,
                "chg_new": cn_col,  "chg_old": co_col,  "water": wt_col,
                "hi_prod": hp_col,  "id_change": ic_col, "hrop": hr_col,
            }

        # ── Save / Apply buttons ──────────────────────────────────────────
        st.divider()
        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button("✅ Apply for this session", type="primary",
                         use_container_width=True):
                key = "mill_col_map" if is_rolling else "ann_col_map"
                st.session_state[key] = {k: v for k, v in new_map.items()
                                          if v is not None}
                st.success(
                    f"Mapping saved for this session. "
                    f"Go to 📊 Dashboard and click Generate."
                )

        with col2:
            # Show the mapping as a TOML snippet to paste into Secrets
            toml_key = "mill_col_map" if is_rolling else "ann_col_map"
            toml_lines = [f"[{toml_key}]"]
            for k, v in new_map.items():
                if v is not None:
                    toml_lines.append(f"{k} = {v}")
            toml_snippet = "\n".join(toml_lines)
            st.download_button(
                "📋 Download as Secrets snippet",
                data=toml_snippet,
                file_name=f"{toml_key}.toml",
                mime="text/plain",
                use_container_width=True,
                help="Paste this into Streamlit Cloud → Settings → Secrets to save permanently"
            )

        with col3:
            if st.button("🔄 Reset to defaults", use_container_width=True):
                key = "mill_col_map" if is_rolling else "ann_col_map"
                st.session_state[key] = {}
                st.success("Reset to built-in column defaults.")

        # ── How to save permanently ────────────────────────────────────────
        with st.expander("💾 How to save this mapping permanently"):
            st.markdown("""
1. Click **Download as Secrets snippet** above
2. Go to **share.streamlit.io** → your app → **⋮ → Settings → Secrets**
3. Paste the downloaded content at the end of your existing Secrets
4. Click **Save** — the app reloads with the mapping as the permanent default

The app also auto-loads saved mappings from Secrets on startup — 
look for `[mill_col_map]` and `[ann_col_map]` sections.
            """)

        # ── Auto-load from Secrets ─────────────────────────────────────────
        # (runs once at page load to pre-populate session state from Secrets)
        for skey in ["mill_col_map", "ann_col_map"]:
            if not st.session_state.get(skey):
                try:
                    saved = dict(st.secrets.get(skey, {}))
                    if saved:
                        st.session_state[skey] = {k: int(v) for k, v in saved.items()}
                except Exception:
                    pass

    else:
        st.info(
            "Upload the MIS file above to see its columns and start mapping. "
            "You only need to do this once — or again if the column layout changes."
        )

        # Show currently active mapping
        for skey, label in [("mill_col_map","Rolling"), ("ann_col_map","Annealing")]:
            cm = st.session_state.get(skey, {})
            if cm:
                st.success(f"✅ Active {label} mapping: {len(cm)} columns mapped")
                import pandas as pd
                base = MIS_COL if skey == "mill_col_map" else ANN_COL
                rows = [{"KPI": k,
                         "Default col": base.get(k,"—"),
                         "Mapped col": cm.get(k, base.get(k,"—")),
                         "Overridden": "✓" if k in cm and cm[k] != base.get(k) else ""}
                        for k in base]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            else:
                st.info(f"ℹ {label}: using default column positions")
