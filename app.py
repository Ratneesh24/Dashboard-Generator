"""
Narrow Complex Day Summary — Streamlit web app.
Read the MIS data from a shared SharePoint Excel LINK, an uploaded .xlsx, or
pasted cells, then render the day-summary dashboard in the browser with a
download. Password-gated via Streamlit secrets.

Run locally:   streamlit run app.py
Deploy:        push to GitHub, then deploy on share.streamlit.io
"""
import datetime
import streamlit as st

import generate_dashboard as gd

st.set_page_config(page_title="Narrow Complex — Day Summary",
                   page_icon="🏭", layout="wide")

# ---------------------------------------------------------------- password gate
def check_password():
    correct = st.secrets.get("app_password", None)
    if not correct:                      # no password set -> open (local dev)
        return True
    if st.session_state.get("auth_ok"):
        return True
    st.title("🔒 Narrow Complex Day Summary")
    pw = st.text_input("Enter access password", type="password")
    if pw:
        if pw == correct:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

check_password()

# ---------------------------------------------------------------- sidebar input
st.sidebar.header("Daily Inputs")
day = st.sidebar.number_input("Report day (date of month)", min_value=1,
                              max_value=31, value=datetime.date.today().day)

mode = st.sidebar.radio(
    "Data source", ["Excel link / upload", "Paste text"],
    help="A shared SharePoint link must be set to 'Anyone with the link'.")

rolling_text = anneal_text = ""
roll_url = roll_file = roll_sheet = ann_url = ann_file = ann_sheet = None

if mode == "Excel link / upload":
    st.sidebar.markdown("**Rolling workbook (CRM06 / CRM07)**")
    roll_url   = st.sidebar.text_input("Rolling Excel link", key="roll_url",
                                       placeholder="https://tslin-my.sharepoint.com/...")
    roll_file  = st.sidebar.file_uploader("…or upload Rolling .xlsx",
                                          type=["xlsx"], key="roll_file")
    roll_sheet = st.sidebar.text_input("Rolling sheet/tab (blank = first)",
                                       key="roll_sheet", placeholder="MAY-26")

    st.sidebar.markdown("**Annealing + 2HI workbook**")
    ann_url   = st.sidebar.text_input("Annealing Excel link", key="ann_url",
                                      placeholder="https://tslin-my.sharepoint.com/...")
    ann_file  = st.sidebar.file_uploader("…or upload Annealing .xlsx",
                                         type=["xlsx"], key="ann_file")
    ann_sheet = st.sidebar.text_input("Annealing sheet/tab (blank = first)",
                                      key="ann_sheet", placeholder="MAY-26")
else:
    st.sidebar.caption("Paste tab-separated cells copied straight from Excel.")
    rolling_text = st.sidebar.text_area("Rolling MIS (CRM06 / CRM07)", height=140)
    anneal_text  = st.sidebar.text_area("Annealing + 2HI / SPM", height=140)

st.sidebar.markdown("**Rolling day targets (MT)**")
t_crm06 = st.sidebar.number_input("CRM06 target", value=260.0, step=10.0)
t_crm07 = st.sidebar.number_input("CRM07 target", value=260.0, step=10.0)

st.sidebar.markdown("**GR headline (T)**")
gr_today  = st.sidebar.number_input("GR today",  value=243.0,  step=1.0)
gr_mtd    = st.sidebar.number_input("GR MTD",    value=6547.0, step=1.0)
gr_target = st.sidebar.number_input("GR target", value=7000.0, step=100.0)

go = st.sidebar.button("Generate dashboard", type="primary")


def load_tsv(url, fileobj, sheet):
    """Return tab-separated text from an upload or a shared link, else ''. """
    src = fileobj.getvalue() if fileobj is not None else ((url or "").strip() or None)
    if not src:
        return ""
    tsv, sheets = gd.xlsx_to_tsv(src, (sheet or "").strip() or None)
    st.sidebar.caption("Tabs found: " + ", ".join(sheets))
    return tsv

# ---------------------------------------------------------------- main
st.title("Narrow Complex — Day Summary Report")

if go:
    if mode == "Excel link / upload":
        try:
            rolling_text = load_tsv(roll_url, roll_file, roll_sheet)
        except Exception as e:
            st.error(f"Could not read the Rolling Excel: {e}")
        try:
            anneal_text = load_tsv(ann_url, ann_file, ann_sheet)
        except Exception as e:
            st.error(f"Could not read the Annealing Excel: {e}")

    data = {
        "date": str(int(day)), "shift": "Day",
        "headline": {"gr_today": gr_today, "gr_mtd": gr_mtd, "gr_target": gr_target},
    }
    if rolling_text.strip():
        rolling = gd.parse_mis(rolling_text, report_day=int(day))
        targets = {"CRM06": t_crm06, "CRM07": t_crm07}
        for mill, m in rolling.items():
            if targets.get(mill):
                m["day_target"] = targets[mill]
        data["rolling"] = rolling
    if anneal_text.strip():
        data["annealing"] = gd.parse_anneal_mis(anneal_text, report_day=int(day))

    if not data.get("rolling") and not data.get("annealing"):
        st.warning("Provide at least one data source (link, upload, or paste), "
                   "then click Generate.")
    else:
        html = gd.build_html(data)
        st.components.v1.html(html, height=1040, scrolling=True)
        st.download_button("⬇️ Download HTML", data=html,
                           file_name=f"day_summary_{int(day)}.html",
                           mime="text/html")
        st.caption("Tip: open the HTML and use the browser's Print → Save as PDF "
                   "for a shareable copy.")
else:
    st.info("Fill in the sidebar and click **Generate dashboard**.")
    if st.button("Show demo dashboard"):
        st.components.v1.html(gd.build_html(gd.demo_data()), height=1040,
                              scrolling=True)
