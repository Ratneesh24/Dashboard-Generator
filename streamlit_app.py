"""
CRM Sahibabad Narrow Complex — Streamlit Dashboard App (v2.0)
Pages: Dashboard | Data Validation | Target Master | Exports
Password-gated via Streamlit secrets (app_password).
"""
import streamlit as st
import datetime, io, os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

st.set_page_config(page_title="CRM Sahibabad — Narrow Complex",
                   page_icon="🏭", layout="wide", initial_sidebar_state="expanded")

# ── Password gate ─────────────────────────────────────────────────────────────
def check_password():
    pw = st.secrets.get("app_password", None)
    if not pw: return True          # no secret set → open (local dev)
    if st.session_state.get("auth"): return True
    st.title("🔒 CRM Sahibabad — Narrow Complex")
    entered = st.text_input("Access password", type="password")
    if entered:
        if entered == pw: st.session_state["auth"] = True; st.rerun()
        else: st.error("Incorrect password.")
    st.stop()

check_password()

# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.image("assets/tata_logo.png", width=120) if os.path.exists("assets/tata_logo.png") else None
st.sidebar.title("CRM Sahibabad")
st.sidebar.caption("Narrow Complex · Daily Operations")

page = st.sidebar.radio("Navigation", ["Dashboard", "Data Validation", "Target Master", "Exports"])
st.sidebar.divider()
day = st.sidebar.number_input("Report day", min_value=1, max_value=31, value=datetime.date.today().day)

st.sidebar.markdown("**Upload MIS files**")
mill_file  = st.sidebar.file_uploader("MILL MIS.xlsx",               type=["xlsx"])
ann_file   = st.sidebar.file_uploader("Annealing & 2HI SPM MIS.xlsx",type=["xlsx"])
crs_file   = st.sidebar.file_uploader("CRS MIS.xlsx",                 type=["xlsx"])
tgt_file   = st.sidebar.file_uploader("Target.xlsx",                  type=["xlsx"])

st.sidebar.divider()
st.sidebar.markdown("**Or use local input/ folder**")
use_local = st.sidebar.checkbox("Use local input/ files", value=True)
run = st.sidebar.button("Generate Dashboard", type="primary")

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Reading MIS files…")
def load_data(mill_bytes, ann_bytes, crs_bytes, tgt_bytes, day_num, use_local):
    from generate_dashboard import xlsx_to_tsv, parse_mis, parse_anneal_mis
    from extractors.targets import parse_targets
    from extractors.crs     import parse_crs
    from config import MILL_FILE, ANNEALING_FILE, CRS_FILE, TARGET_FILE

    def _load(uploaded, local_path, parser, *args, **kwargs):
        if uploaded:
            data = uploaded if isinstance(uploaded, bytes) else uploaded
            if hasattr(data, "getvalue"): data = data.getvalue()
            return parser(data, *args, **kwargs)
        if use_local and os.path.exists(local_path):
            return parser(local_path, *args, **kwargs)
        return {}

    def _load_tsv(uploaded, local_path, parse_fn, *args, **kwargs):
        if uploaded:
            raw = uploaded if isinstance(uploaded, bytes) else uploaded.getvalue()
        elif use_local and os.path.exists(local_path):
            raw = open(local_path, "rb").read()
        else:
            return {}
        tsv, _ = xlsx_to_tsv(raw)
        return parse_fn(tsv, *args, **kwargs)

    targets  = _load(tgt_bytes, TARGET_FILE, parse_targets) or {}
    crs      = _load(crs_bytes, CRS_FILE,    parse_crs)     or {}
    rolling  = _load_tsv(mill_bytes, MILL_FILE, parse_mis, report_day=day_num) or {}
    for mn, m in rolling.items():
        t = targets.get("rolling_day")
        if t: m["day_target"] = t / max(len(rolling), 1)
    annealing = _load_tsv(ann_bytes, ANNEALING_FILE, parse_anneal_mis, report_day=day_num) or {}

    return {
        "date":      crs.get("report_date", datetime.date.today().strftime("%d.%m.%Y")),
        "rolling":   rolling,
        "annealing": annealing,
        "crs":       crs,
    }, targets

# ── Session state ─────────────────────────────────────────────────────────────
if "data" not in st.session_state:
    st.session_state["data"] = {}
    st.session_state["targets"] = {}
    st.session_state["html"] = ""
    st.session_state["alerts"] = []
    st.session_state["notes"] = []

if run:
    with st.spinner("Loading & processing MIS files…"):
        data, targets = load_data(
            mill_file, ann_file, crs_file, tgt_file, int(day), use_local)
        from dashboard.alerts import generate_alerts
        from dashboard.notes  import generate_notes
        from dashboard.layout import build_html
        alerts = generate_alerts(data, targets)
        notes  = generate_notes(data, targets, alerts)
        html   = build_html(data, targets, alerts, notes,
                            report_date=data.get("date"),
                            report_time=datetime.datetime.now().strftime("%H:%M"))
        st.session_state["data"]    = data
        st.session_state["targets"] = targets
        st.session_state["html"]    = html
        st.session_state["alerts"]  = alerts
        st.session_state["notes"]   = notes
    st.success(f"Dashboard ready — {len(alerts)} alert(s), {len(notes)} note(s)")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 1: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if page == "Dashboard":
    st.title("Daily Operations Dashboard")
    if st.session_state["html"]:
        st.components.v1.html(st.session_state["html"], height=1120, scrolling=True)
        st.download_button("⬇ Download HTML", data=st.session_state["html"],
                           file_name=f"Dashboard_{st.session_state['data'].get('date','today')}.html",
                           mime="text/html")
    else:
        st.info("Upload MIS files in the sidebar and click **Generate Dashboard**.")
        from generate_dashboard import demo_data
        if st.button("Preview demo dashboard"):
            from dashboard.layout import build_html
            from dashboard.alerts import generate_alerts
            from dashboard.notes  import generate_notes
            dd = demo_data()
            tgt = {"rolling_day":200,"ann_day":150,"spm_day":47,
                   "tube_gr_day":167,"oem_gr_day":27,"total_gr_day":193}
            al  = generate_alerts(dd, tgt)
            nt  = generate_notes(dd, tgt, al)
            st.components.v1.html(build_html(dd, tgt, al, nt), height=1120, scrolling=True)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 2: DATA VALIDATION
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Data Validation":
    st.title("Data Validation — Extracted KPIs")
    data    = st.session_state["data"]
    targets = st.session_state["targets"]
    if not data:
        st.info("Generate a dashboard first."); st.stop()

    st.subheader("Rolling Mills")
    rows = []
    for mn, m in data.get("rolling", {}).items():
        rows.append({"Mill": mn, "Day Total": m.get("day_total"), "Day Roll": m.get("day_roll"),
                     "Day R/R": m.get("day_rr"), "Yield%": m.get("yield"),
                     "Util%": m.get("day_util"), "Avg Gauge": m.get("avg_gauge"),
                     "Avg Width": m.get("avg_width"), "Cumm": m.get("cumm_roll")})
    if rows:
        import pandas as pd
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

    st.subheader("Annealing")
    ann = data.get("annealing", {}).get("ann02", {})
    if ann: st.json(ann)

    st.subheader("2HI Skin Pass")
    spm = data.get("annealing", {}).get("skin_pass", {})
    if spm: st.json(spm)

    st.subheader("CRS — GR & Inventory")
    crs = data.get("crs", {})
    if crs: st.json(crs)

    st.subheader("Alerts")
    for a in st.session_state["alerts"]:
        col = {"critical":"🔴","warning":"🟡","info":"🔵"}.get(a["priority"],"●")
        st.write(f"{col} **{a['title']}** — {a.get('detail','')}")

    st.subheader("Management Notes")
    for n in st.session_state["notes"]:
        st.write(f"• {n}")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 3: TARGET MASTER
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Target Master":
    st.title("Target Master — View & Edit")
    targets = st.session_state["targets"]
    if not targets:
        st.info("Load Target.xlsx first via the sidebar uploader or local input/ folder.")
    else:
        st.caption("These values are loaded from Target.xlsx. "
                   "Edit here to override for this session only.")
        new_targets = dict(targets)
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Day Targets")
            for k, v in {k: v for k, v in targets.items() if "day" in k and v is not None}.items():
                new_targets[k] = st.number_input(k, value=float(v), step=1.0, key=k)
        with col2:
            st.subheader("MTD Targets")
            for k, v in {k: v for k, v in targets.items() if "mtd" in k and v is not None}.items():
                new_targets[k] = st.number_input(k, value=float(v), step=10.0, key=k)
        if st.button("Apply overrides"):
            st.session_state["targets"] = new_targets
            st.success("Targets updated for this session. Click Generate Dashboard to refresh.")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 4: EXPORTS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Exports":
    st.title("Exports")
    html = st.session_state.get("html", "")
    data = st.session_state.get("data", {})
    date_str = data.get("date", datetime.date.today().strftime("%d-%m-%Y"))

    if not html:
        st.info("Generate a dashboard first."); st.stop()

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("### HTML")
        st.download_button("⬇ Download HTML", data=html,
                           file_name=f"Daily_Dashboard_{date_str}.html",
                           mime="text/html")
        st.caption("Open in browser → Print → Save as PDF for a clean PDF copy.")

    with col2:
        st.markdown("### PNG (1920×1080)")
        if st.button("Render PNG", key="png_btn"):
            with st.spinner("Rendering PNG via Chromium…"):
                try:
                    import tempfile
                    from playwright.sync_api import sync_playwright
                    with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w") as tf:
                        tf.write(html); tf_path = tf.name
                    with sync_playwright() as p:
                        b = p.chromium.launch()
                        pg = b.new_page(viewport={"width":1920,"height":1080})
                        pg.goto("file://"+tf_path)
                        pg.wait_for_load_state("networkidle")
                        png = pg.screenshot(full_page=False,
                                            clip={"x":0,"y":0,"width":1920,"height":1080})
                        b.close()
                    st.session_state["png"] = png
                    st.success("PNG rendered.")
                except Exception as e:
                    st.error(f"PNG failed: {e}")
        if st.session_state.get("png"):
            st.download_button("⬇ Download PNG", data=st.session_state["png"],
                               file_name=f"Daily_Dashboard_{date_str}.png", mime="image/png")

    with col3:
        st.markdown("### PowerPoint")
        if st.button("Render PPTX", key="ppt_btn"):
            if not st.session_state.get("png"):
                st.warning("Render PNG first.")
            else:
                with st.spinner("Building PPTX…"):
                    try:
                        from pptx import Presentation
                        from pptx.util import Inches
                        import tempfile
                        prs = Presentation()
                        prs.slide_width = Inches(16)
                        prs.slide_height = Inches(9)
                        sl = prs.slides.add_slide(prs.slide_layouts[6])
                        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                            tmp.write(st.session_state["png"]); png_path = tmp.name
                        sl.shapes.add_picture(png_path, Inches(0), Inches(0), Inches(16), Inches(9))
                        buf = io.BytesIO(); prs.save(buf); buf.seek(0)
                        st.session_state["pptx"] = buf.read()
                        st.success("PPTX ready.")
                    except Exception as e:
                        st.error(f"PPTX failed: {e}")
        if st.session_state.get("pptx"):
            st.download_button("⬇ Download PPTX", data=st.session_state["pptx"],
                               file_name=f"Daily_Dashboard_{date_str}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    st.divider()
    st.markdown("### Dashboard Data (Excel audit trail)")
    if st.button("Export dashboard_data.xlsx"):
        try:
            import openpyxl, io as _io
            from dashboard_generator import export_xlsx
            targets = st.session_state.get("targets", {})
            buf = _io.BytesIO()
            # write to temp path then read back
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                export_xlsx(data, targets, tmp.name)
                buf.write(open(tmp.name,"rb").read())
            buf.seek(0)
            st.download_button("⬇ Download dashboard_data.xlsx", data=buf,
                               file_name=f"dashboard_data_{date_str}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Excel export failed: {e}")
