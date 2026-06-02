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
    parse_targets, parse_crs,
    generate_alerts, generate_notes,
    build_html,
)

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
mill_file = st.sidebar.file_uploader("MILL MIS.xlsx",                 type=["xlsx"], key="mill")
ann_file  = st.sidebar.file_uploader("Annealing & 2HI SPM MIS.xlsx", type=["xlsx"], key="ann")
crs_file  = st.sidebar.file_uploader("CRS MIS.xlsx",                  type=["xlsx"], key="crs")
tgt_file  = st.sidebar.file_uploader("Target.xlsx",                   type=["xlsx"], key="tgt")

st.sidebar.divider()
st.sidebar.markdown("**Or paste SharePoint links**")
roll_url   = st.sidebar.text_input("Rolling Excel link",   key="roll_url",
                                    placeholder="https://tslin-my.sharepoint.com/...")
ann_url    = st.sidebar.text_input("Annealing Excel link", key="ann_url",
                                    placeholder="https://tslin-my.sharepoint.com/...")
crs_url    = st.sidebar.text_input("CRS Excel link",       key="crs_url",
                                    placeholder="https://tslin-my.sharepoint.com/...")
tgt_url    = st.sidebar.text_input("Target Excel link",    key="tgt_url",
                                    placeholder="https://tslin-my.sharepoint.com/...")
roll_sheet = st.sidebar.text_input("Rolling sheet/tab",   value="", key="rsheet",
                                    placeholder="e.g. MAY-26")
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

            targets  = _obj(tgt_file, tgt_url, inp("Target.xlsx"),
                            parse_targets, "Target")
            crs      = _obj(crs_file, crs_url, inp("CRS MIS.xlsx"),
                            parse_crs, "CRS")
            roll_tsv = _tsv(mill_file, roll_url, inp("MILL MIS.xlsx"),
                            roll_sheet, "Rolling")
            ann_tsv  = _tsv(ann_file, ann_url,
                            inp("Annealing and 2HI SPM MIS.xlsx"),
                            ann_sheet, "Annealing")

            rolling   = parse_mis(roll_tsv, report_day=int(day)) if roll_tsv else {}
            annealing = parse_anneal_mis(ann_tsv, report_day=int(day)) if ann_tsv else {}
            for mn, m in rolling.items():
                t = targets.get("rolling_day")
                if t: m["day_target"] = t / max(len(rolling), 1)

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
            st.success(f"✅ Dashboard ready — {len(alerts)} alerts, {crit} critical, "
                       f"{len(notes)} management notes")
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
    st.title("Target Master — View & Edit")
    targets = st.session_state["targets"]
    if not targets: st.info("Load Target.xlsx first."); st.stop()
    st.caption("Edits apply to this session only and do not save back to the file.")
    new_t = dict(targets)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Day Targets**")
        for k,v in {k:v for k,v in targets.items()
                    if "day" in k and v not in (None,"NA")}.items():
            new_t[k] = st.number_input(k, value=float(v), step=1.0, key=f"t_{k}")
    with c2:
        st.markdown("**MTD Targets**")
        for k,v in {k:v for k,v in targets.items()
                    if "mtd" in k and v not in (None,"NA")}.items():
            new_t[k] = st.number_input(k, value=float(v), step=10.0, key=f"t_{k}")
    if st.button("Apply overrides", type="primary"):
        st.session_state["targets"] = new_t
        st.success("Targets updated — regenerate dashboard to see changes.")

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
