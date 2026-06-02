"""
CRM Sahibabad Narrow Complex — Streamlit Dashboard App (v2.1)
FIXED: All imports now work from repo root (no subfolder path issues).
Pages: Dashboard | Data Validation | Target Master | Exports
Password via Streamlit secrets: app_password
"""
import streamlit as st
import datetime, io, os, sys

# ── Ensure repo root is on path (works both locally and on Streamlit Cloud) ──
ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

# ── Imports (all from root-level modules) ────────────────────────────────────
from generate_dashboard import xlsx_to_tsv, parse_mis, parse_anneal_mis
from extractors.targets  import parse_targets
from extractors.crs      import parse_crs
from dashboard.alerts    import generate_alerts
from dashboard.notes     import generate_notes
from dashboard.layout    import build_html

st.set_page_config(
    page_title="CRM Sahibabad — Narrow Complex",
    page_icon="🏭", layout="wide",
    initial_sidebar_state="expanded"
)

# ── Password gate ─────────────────────────────────────────────────────────────
def check_password():
    pw = st.secrets.get("app_password", None)
    if not pw:
        return True          # no secret → open (local dev / testing)
    if st.session_state.get("auth"):
        return True
    st.title("🔒 CRM Sahibabad — Narrow Complex")
    entered = st.text_input("Access password", type="password")
    if entered:
        if entered == pw:
            st.session_state["auth"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

check_password()

# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.title("CRM Sahibabad")
st.sidebar.caption("Narrow Complex · Daily Operations")

page = st.sidebar.radio(
    "Navigation",
    ["📊 Dashboard", "✅ Data Validation", "🎯 Target Master", "⬇ Exports"]
)
st.sidebar.divider()

day = st.sidebar.number_input(
    "Report day (date of month)",
    min_value=1, max_value=31,
    value=datetime.date.today().day
)

st.sidebar.markdown("**Upload MIS files**")
mill_file = st.sidebar.file_uploader(
    "MILL MIS.xlsx",                type=["xlsx"], key="mill")
ann_file  = st.sidebar.file_uploader(
    "Annealing & 2HI SPM MIS.xlsx", type=["xlsx"], key="ann")
crs_file  = st.sidebar.file_uploader(
    "CRS MIS.xlsx",                 type=["xlsx"], key="crs")
tgt_file  = st.sidebar.file_uploader(
    "Target.xlsx",                  type=["xlsx"], key="tgt")

st.sidebar.divider()
st.sidebar.markdown("**Or use SharePoint links**")
roll_url = st.sidebar.text_input("Rolling Excel link",    key="roll_url",
                                  placeholder="https://tslin-my.sharepoint.com/...")
ann_url  = st.sidebar.text_input("Annealing Excel link",  key="ann_url",
                                  placeholder="https://tslin-my.sharepoint.com/...")
crs_url  = st.sidebar.text_input("CRS Excel link",        key="crs_url",
                                  placeholder="https://tslin-my.sharepoint.com/...")
tgt_url  = st.sidebar.text_input("Target Excel link",     key="tgt_url",
                                  placeholder="https://tslin-my.sharepoint.com/...")

roll_sheet = st.sidebar.text_input("Rolling sheet/tab",   value="", key="rsheet",
                                    placeholder="e.g. MAY-26")
ann_sheet  = st.sidebar.text_input("Annealing sheet/tab", value="", key="asheet",
                                    placeholder="e.g. MAY-26")

run = st.sidebar.button("⚡ Generate Dashboard", type="primary", use_container_width=True)

# ── Helper: load xlsx from upload, link, or local path ────────────────────────
def _load_tsv(uploaded_file, url, local_path, sheet=None, label=""):
    """Try upload → URL → local file. Returns TSV text or ''."""
    src = None
    if uploaded_file is not None:
        src = uploaded_file.getvalue()
    elif url and url.strip().startswith("http"):
        src = url.strip()
    elif local_path and os.path.exists(local_path):
        src = local_path
    if src is None:
        return ""
    try:
        tsv, sheets = xlsx_to_tsv(src, sheet or None)
        if sheets:
            st.sidebar.caption(f"{label} tabs: {', '.join(sheets)}")
        return tsv
    except Exception as e:
        st.sidebar.error(f"{label} read error: {e}")
        return ""

def _load_parser(uploaded_file, url, local_path, parser_fn, label=""):
    """Load a single-parser file (targets, crs) from upload/url/local."""
    src = None
    if uploaded_file is not None:
        src = uploaded_file.getvalue()
    elif url and url.strip().startswith("http"):
        src = url.strip()
    elif local_path and os.path.exists(local_path):
        src = local_path
    if src is None:
        return {}
    try:
        if isinstance(src, bytes):
            # write to temp file since parsers accept path or bytes
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(src); tmp_path = tmp.name
            return parser_fn(tmp_path)
        return parser_fn(src)
    except Exception as e:
        st.sidebar.error(f"{label} parse error: {e}")
        return {}

# ── Session state defaults ────────────────────────────────────────────────────
for k, v in [("data",{}),("targets",{}),("html",""),
              ("alerts",[]),("notes",[]),("png",None),("pptx",None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Generate dashboard on button click ───────────────────────────────────────
if run:
    with st.spinner("Reading MIS files and generating dashboard…"):
        try:
            # Local fallback paths
            L = lambda name: os.path.join(ROOT, "input", name)

            targets = _load_parser(tgt_file, tgt_url, L("Target.xlsx"),
                                   parse_targets, "Target")
            crs     = _load_parser(crs_file, crs_url, L("CRS MIS.xlsx"),
                                   parse_crs, "CRS")

            roll_tsv = _load_tsv(mill_file, roll_url, L("MILL MIS.xlsx"),
                                  roll_sheet, "Rolling")
            rolling = parse_mis(roll_tsv, report_day=int(day)) if roll_tsv else {}
            for mn, m in rolling.items():
                t = targets.get("rolling_day")
                if t: m["day_target"] = t / max(len(rolling), 1)

            ann_tsv = _load_tsv(ann_file, ann_url,
                                 L("Annealing and 2HI SPM MIS.xlsx"),
                                 ann_sheet, "Annealing")
            annealing = parse_anneal_mis(ann_tsv, report_day=int(day)) if ann_tsv else {}

            data = {
                "date":      crs.get("report_date", datetime.date.today().strftime("%d.%m.%Y")),
                "rolling":   rolling,
                "annealing": annealing,
                "crs":       crs,
            }
            alerts = generate_alerts(data, targets)
            notes  = generate_notes(data, targets, alerts)
            html   = build_html(
                data, targets, alerts, notes,
                report_date=data.get("date"),
                report_time=datetime.datetime.now().strftime("%H:%M")
            )
            st.session_state.update(
                data=data, targets=targets, html=html,
                alerts=alerts, notes=notes, png=None, pptx=None
            )
            crit = sum(1 for a in alerts if a["priority"] == "critical")
            st.success(f"✅ Dashboard ready — {len(alerts)} alert(s), "
                       f"{crit} critical, {len(notes)} management note(s)")
        except Exception as e:
            st.error(f"Generation failed: {e}")
            import traceback; st.code(traceback.format_exc())

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 1: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if page == "📊 Dashboard":
    st.title("Daily Operations Dashboard")
    if st.session_state["html"]:
        st.components.v1.html(st.session_state["html"], height=1120, scrolling=True)
        col1, col2 = st.columns(2)
        with col1:
            d = st.session_state["data"]
            st.download_button(
                "⬇ Download HTML",
                data=st.session_state["html"],
                file_name=f"Dashboard_{d.get('date','today')}.html",
                mime="text/html", use_container_width=True
            )
        with col2:
            st.info("Open the HTML → Print → Save as PDF for a clean PDF copy.")
    else:
        st.info("Upload MIS files in the sidebar, then click **⚡ Generate Dashboard**.")
        if st.button("👁 Preview with demo data"):
            from generate_dashboard import demo_data
            dd = demo_data()
            tgt = {
                "rolling_day":200,"ann_day":150,"spm_day":47,
                "tube_gr_day":167,"oem_gr_day":27,"total_gr_day":193,
                "crm06_yield_mtd":99,"crm06_util_mtd":80,
                "crm07_yield_mtd":99,"crm07_util_mtd":80,
            }
            al = generate_alerts(dd, tgt)
            nt = generate_notes(dd, tgt, al)
            st.components.v1.html(build_html(dd, tgt, al, nt), height=1120, scrolling=True)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 2: DATA VALIDATION
# ══════════════════════════════════════════════════════════════════════════════
elif page == "✅ Data Validation":
    st.title("Data Validation — Extracted KPIs")
    data = st.session_state["data"]
    targets = st.session_state["targets"]
    if not data:
        st.info("Generate a dashboard first.")
        st.stop()

    st.subheader("Rolling Mills")
    import pandas as pd
    rows = []
    for mn, m in data.get("rolling", {}).items():
        rows.append({
            "Mill": mn, "Day Total": m.get("day_total"), "Day Target": m.get("day_target"),
            "Day Roll": m.get("day_roll"), "Day R/R": m.get("day_rr"),
            "Yield %": m.get("yield"), "Util %": m.get("day_util"),
            "Avg Gauge (mm)": m.get("avg_gauge"), "Avg Width (mm)": m.get("avg_width"),
            "Cumm Output": m.get("cumm_out"), "Cumm Yield %": m.get("cumm_yield"),
        })
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Annealing")
        ann = data.get("annealing", {}).get("ann02", {})
        if ann: st.json(ann)
        st.subheader("2HI Skin Pass")
        spm = data.get("annealing", {}).get("skin_pass", {})
        if spm: st.json(spm)
    with col2:
        st.subheader("CRS — Inventory & GR")
        crs = data.get("crs", {})
        if crs: st.json(crs)

    st.subheader("Alerts")
    icon = {"critical":"🔴","warning":"🟡","info":"🔵"}
    for a in st.session_state["alerts"]:
        st.write(f"{icon.get(a['priority'],'●')} **{a['title']}** — {a.get('detail','')}")

    st.subheader("Management Notes")
    for n in st.session_state["notes"]:
        st.write(f"✅ {n}")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 3: TARGET MASTER
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🎯 Target Master":
    st.title("Target Master — View & Edit")
    targets = st.session_state["targets"]
    if not targets:
        st.info("Load Target.xlsx first (upload or link in sidebar, then Generate).")
        st.stop()

    st.caption("Values loaded from Target.xlsx. Edit to override for this session only — "
               "changes do NOT save back to the file.")
    new_t = dict(targets)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Day Targets**")
        for k, v in {k: v for k, v in targets.items()
                     if "day" in k and v is not None and v != "NA"}.items():
            new_t[k] = st.number_input(k, value=float(v), step=1.0, key=f"t_{k}")
    with c2:
        st.markdown("**MTD Targets**")
        for k, v in {k: v for k, v in targets.items()
                     if "mtd" in k and v is not None and v != "NA"}.items():
            new_t[k] = st.number_input(k, value=float(v), step=10.0, key=f"t_{k}")
    if st.button("Apply overrides", type="primary"):
        st.session_state["targets"] = new_t
        st.success("Targets updated. Click Generate Dashboard to refresh the dashboard.")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 4: EXPORTS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "⬇ Exports":
    st.title("Exports")
    html    = st.session_state.get("html", "")
    data    = st.session_state.get("data", {})
    date_str = data.get("date", datetime.date.today().strftime("%d-%m-%Y"))

    if not html:
        st.info("Generate a dashboard first.")
        st.stop()

    st.markdown("### HTML")
    st.download_button(
        "⬇ Download HTML dashboard",
        data=html,
        file_name=f"Daily_Dashboard_{date_str}.html",
        mime="text/html", use_container_width=True
    )
    st.caption("Open in any browser → Ctrl+P → Save as PDF")

    st.divider()
    st.markdown("### PNG (1920 × 1080)")
    st.caption("Requires Chromium/Playwright. Works locally; may not be available on all cloud plans.")
    if st.button("Render PNG", use_container_width=True):
        with st.spinner("Rendering via Chromium…"):
            try:
                import tempfile
                from playwright.sync_api import sync_playwright
                with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w",
                                                  encoding="utf-8") as tf:
                    tf.write(html); tf_path = tf.name
                with sync_playwright() as p:
                    b = p.chromium.launch()
                    pg = b.new_page(viewport={"width": 1920, "height": 1080})
                    pg.goto("file://" + tf_path)
                    pg.wait_for_load_state("networkidle")
                    png = pg.screenshot(full_page=False,
                                        clip={"x":0,"y":0,"width":1920,"height":1080})
                    b.close()
                st.session_state["png"] = png
                st.success("PNG ready.")
            except Exception as e:
                st.error(f"PNG failed: {e}")

    if st.session_state.get("png"):
        st.download_button(
            "⬇ Download PNG",
            data=st.session_state["png"],
            file_name=f"Daily_Dashboard_{date_str}.png",
            mime="image/png", use_container_width=True
        )

    st.divider()
    st.markdown("### PowerPoint (16 : 9)")
    if st.button("Build PPTX", use_container_width=True):
        if not st.session_state.get("png"):
            st.warning("Render PNG first — the PPTX embeds the PNG.")
        else:
            with st.spinner("Building PPTX…"):
                try:
                    from pptx import Presentation
                    from pptx.util import Inches
                    import tempfile
                    prs = Presentation()
                    prs.slide_width  = Inches(16)
                    prs.slide_height = Inches(9)
                    sl = prs.slides.add_slide(prs.slide_layouts[6])
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                        tmp.write(st.session_state["png"]); tmp_path = tmp.name
                    sl.shapes.add_picture(tmp_path, Inches(0), Inches(0),
                                          Inches(16), Inches(9))
                    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
                    st.session_state["pptx"] = buf.read()
                    st.success("PPTX ready.")
                except Exception as e:
                    st.error(f"PPTX failed: {e}")

    if st.session_state.get("pptx"):
        st.download_button(
            "⬇ Download PPTX",
            data=st.session_state["pptx"],
            file_name=f"Daily_Dashboard_{date_str}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
