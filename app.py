"""
Narrow Complex Day Summary — Streamlit web app.
Lets a user paste / upload the MIS sheets and get the dashboard in the browser,
with a PNG/HTML download. Password-gated via Streamlit secrets.

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
    """Returns True once the correct password is entered."""
    correct = st.secrets.get("app_password", None)
    if not correct:          # no password configured -> open (local dev)
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

st.sidebar.caption("Paste the tab-separated MIS data (copy the cells straight "
                   "from Excel). Leave a box empty to skip that section.")

rolling_text = st.sidebar.text_area(
    "Rolling MIS  (CRM06 / CRM07 tables)", height=160,
    placeholder="DATE\tINPUT ROLLING\tROLLING\t...")

anneal_text = st.sidebar.text_area(
    "Annealing + 2HI / SPM sheet", height=160,
    placeholder="5/1/2026\t0.000\t74.680\t...")

st.sidebar.markdown("**Rolling day targets (MT)**")
t_crm06 = st.sidebar.number_input("CRM06 target", value=260.0, step=10.0)
t_crm07 = st.sidebar.number_input("CRM07 target", value=260.0, step=10.0)

st.sidebar.markdown("**GR headline (T)**")
gr_today  = st.sidebar.number_input("GR today",  value=243.0,  step=1.0)
gr_mtd    = st.sidebar.number_input("GR MTD",    value=6547.0, step=1.0)
gr_target = st.sidebar.number_input("GR target", value=7000.0, step=100.0)

go = st.sidebar.button("Generate dashboard", type="primary")

# ---------------------------------------------------------------- main
st.title("Narrow Complex — Day Summary Report")

if go:
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
        st.warning("Paste at least one MIS section in the sidebar, then click "
                   "Generate.")
    else:
        html = gd.build_html(data)
        # live preview (the HTML is a fixed 1500x1000 canvas)
        st.components.v1.html(html, height=1040, scrolling=True)
        st.download_button("⬇️ Download HTML", data=html,
                           file_name=f"day_summary_{int(day)}.html",
                           mime="text/html")
        st.caption("Tip: open the HTML and use your browser's Print → Save as PDF "
                   "for a shareable copy.")
else:
    st.info("Fill in the sidebar and click **Generate dashboard**. "
            "Try the demo to see the layout.")
    if st.button("Show demo dashboard"):
        st.components.v1.html(gd.build_html(gd.demo_data()), height=1040,
                              scrolling=True)
