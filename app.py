"""
Digital Pavement Management System
TCG633 - Bridge & Road Maintenance | Individual Project

Single-section quick-entry workflow (dark UI) on top of the same PCI/IRI/Hybrid
engineering logic used throughout this project. Works as a single standalone
file - no external data folder required.

Run with:  streamlit run app.py
"""

import io
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(page_title="Digital Pavement Management System", page_icon="🚧", layout="wide")

# --------------------------------------------------------------------------------------
# Dark theme (works even without a .streamlit/config.toml - everything below is plain
# CSS injected into this single file, so the look is independent of deployment setup)
# --------------------------------------------------------------------------------------
st.markdown("""
<style>
.stApp { background-color: #0e1117; color: #e6e6e6; }
section[data-testid="stSidebar"] { background-color: #11141b; }
h1,h2,h3,h4,h5,h6, p, span, label, .stMarkdown { color: #e6e6e6 !important; }
div[data-testid="stMetricValue"] { color: #ffffff; }
hr { border-color: #2a2f3a; }

.pms-card {
  background: #161b24; border: 1px solid #262b36; border-radius: 12px;
  padding: 22px 26px; margin-bottom: 18px;
}
.pms-title { font-size: 2.1rem; font-weight: 800; margin-bottom: 0px; }
.pms-subtitle { color: #9aa3b2; font-size: 0.95rem; margin-top: -6px; margin-bottom: 22px; }
.pms-bignum { font-size: 2.6rem; font-weight: 800; line-height: 1.1; }
.pms-label { color: #9aa3b2; font-size: 0.85rem; text-transform: uppercase; letter-spacing: .04em; }
.pms-banner {
  border-radius: 8px; padding: 12px 16px; font-size: 0.98rem; margin-top: 10px;
  background: #16314f; border: 1px solid #1f4068; color: #cfe3ff;
}
.pms-gauge-wrap { margin-top: 18px; }
.pms-gauge-track {
  position: relative; height: 10px; border-radius: 6px; margin-top: 10px;
  background: linear-gradient(to right,
    #C62828 0%, #C62828 55%, #F9A825 55%, #F9A825 70%,
    #9CCC65 70%, #9CCC65 85%, #2E7D32 85%, #2E7D32 100%);
}
.pms-gauge-dot {
  position: absolute; top: -5px; width: 20px; height: 20px; border-radius: 50%;
  background: #ffffff; border: 4px solid var(--dotcolor, #C62828);
  transform: translateX(-50%); box-shadow: 0 0 6px rgba(0,0,0,.5);
}
.pms-pill {
  display:inline-block; padding: 2px 10px; border-radius: 999px; font-size: 0.78rem;
  font-weight: 700; color: #0e1117;
}

/* Native widget dark styling */
button[data-testid^="stBaseButton-secondary"] {
  background-color: #1b2230 !important; color: #e6e6e6 !important;
  border: 1px solid #2f3645 !important;
}
button[data-testid^="stBaseButton-secondary"]:hover {
  border-color: #5DADE2 !important; color: #5DADE2 !important;
}
input[data-testid="stTextInputRootElement"], div[data-testid="stTextInputRootElement"],
.stTextInput input, .stNumberInput input, div[data-testid="stNumberInputContainer"] {
  background-color: #1b2230 !important; color: #e6e6e6 !important;
  border-color: #2f3645 !important;
}
div[data-baseweb="select"] > div {
  background-color: #1b2230 !important; border-color: #2f3645 !important; color: #e6e6e6 !important;
}
div[data-baseweb="popover"] li { background-color: #1b2230 !important; color: #e6e6e6 !important; }
section[data-testid="stFileUploaderDropzone"] {
  background-color: #1b2230 !important; border: 1px dashed #2f3645 !important;
}
section[data-testid="stFileUploaderDropzone"] button {
  background-color: #11141b !important; color: #e6e6e6 !important; border: 1px solid #2f3645 !important;
}
div[data-testid="stDataFrame"] { border: 1px solid #262b36 !important; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------------------
# Engineering constants (same logic as the rest of the project, single-defect-per-
# section "quick entry" model - matches how the dataset was supplied: one representative
# defect record per 100 m section)
# --------------------------------------------------------------------------------------
DEFECT_WEIGHTS = {
    "Longitudinal Crack": 1.0, "Alligator (Fatigue) Crack": 1.6, "Potholes": 2.2,
    "Raveling": 1.2, "Depression/Sag": 1.4, "Patching (Failed)": 1.8,
    "Bleeding/Flushing": 1.0, "Rut/Rutting": 1.6,
}
SEVERITY_FACTORS = {"Low": 0.6, "Medium": 1.0, "High": 1.4}

BANDS = [
    (85, 100, "Very Good", "#2E7D32", "Routine Maintenance"),
    (70, 85, "Good / Satisfactory", "#9CCC65", "Preventive Maintenance"),
    (55, 70, "Fair", "#F9A825", "Overlay / Corrective Maintenance"),
    (0, 55, "Poor", "#C62828", "Major Rehabilitation / Reconstruction"),
]

SEED = [
    {"Section ID": "S1", "Defect Type": "Longitudinal Crack", "Severity": "Low", "Area Affected (%)": 5, "IRI (m/km)": 1.60},
    {"Section ID": "S2", "Defect Type": "Potholes", "Severity": "Medium", "Area Affected (%)": 3, "IRI (m/km)": 1.91},
    {"Section ID": "S3", "Defect Type": "Raveling", "Severity": "Low", "Area Affected (%)": 10, "IRI (m/km)": 2.21},
    {"Section ID": "S4", "Defect Type": "Alligator (Fatigue) Crack", "Severity": "Medium", "Area Affected (%)": 6, "IRI (m/km)": 2.51},
    {"Section ID": "S5", "Defect Type": "Depression/Sag", "Severity": "Low", "Area Affected (%)": 4, "IRI (m/km)": 2.01},
    {"Section ID": "S6", "Defect Type": "Patching (Failed)", "Severity": "High", "Area Affected (%)": 2, "IRI (m/km)": 2.81},
    {"Section ID": "S7", "Defect Type": "Alligator (Fatigue) Crack", "Severity": "High", "Area Affected (%)": 10, "IRI (m/km)": 3.31},
    {"Section ID": "S8", "Defect Type": "Potholes", "Severity": "High", "Area Affected (%)": 11, "IRI (m/km)": 3.81},
    {"Section ID": "S9", "Defect Type": "Patching (Failed)", "Severity": "High", "Area Affected (%)": 20, "IRI (m/km)": 4.51},
    {"Section ID": "S10", "Defect Type": "Rut/Rutting", "Severity": "High", "Area Affected (%)": 30, "IRI (m/km)": 5.51},
]


def classify(score):
    for lo, hi, label, color, action in BANDS:
        if score >= lo:
            return label, color, action
    return BANDS[-1][2], BANDS[-1][3], BANDS[-1][4]


def iri_to_score(iri):
    """Piecewise-linear mapping of IRI (m/km) onto the same 0-100 scale as PCI, so
    that an IRI value classed 'Very Good' always lands in the Very-Good score range
    (>=85), 'Good' in 70-85, etc. Keeps PCI-based and IRI-based classification
    consistent when blended into the Hybrid Index below."""
    if iri <= 2:
        return 100 - (iri / 2) * 15            # 0->100 ... 2->85
    if iri <= 3:
        return 85 - (iri - 2) * 15             # 2->85 ... 3->70
    if iri <= 4:
        return 70 - (iri - 3) * 15             # 3->70 ... 4->55
    return max(0.0, 55 - (iri - 4) * 13.75)    # 4->55 ... 8->0


def compute_row(rec, pci_weight):
    area = float(rec.get("Area Affected (%)", 0) or 0)
    weight = DEFECT_WEIGHTS.get(rec["Defect Type"], 0)
    sev = SEVERITY_FACTORS.get(rec["Severity"], 0)
    deduct = area * weight * sev
    pci = max(0.0, 100 - min(100, deduct))
    iri = float(rec.get("IRI (m/km)", 0) or 0)
    iri_score = iri_to_score(iri)
    hybrid = pci_weight * pci + (1 - pci_weight) * iri_score
    label, color, action = classify(hybrid)
    return {
        **rec,
        "PCI": round(pci, 1),
        "Hybrid Index": round(hybrid, 1),
        "Final Condition": label,
        "Condition Color": color,
        "Maintenance Action": action,
    }


def compute_all(records, pci_weight):
    if not records:
        return pd.DataFrame(columns=["Section ID", "Defect Type", "Severity", "Area Affected (%)",
                                      "IRI (m/km)", "PCI", "Hybrid Index", "Final Condition",
                                      "Condition Color", "Maintenance Action"])
    return pd.DataFrame([compute_row(r, pci_weight) for r in records])


# --------------------------------------------------------------------------------------
# Session state
# --------------------------------------------------------------------------------------
if "records" not in st.session_state:
    st.session_state.records = [dict(r) for r in SEED]
if "last_section" not in st.session_state:
    st.session_state.last_section = "S1"

# --------------------------------------------------------------------------------------
# Header
# --------------------------------------------------------------------------------------
st.markdown(
    '<div class="pms-title">🚧 Digital Pavement Management System</div>'
    '<div class="pms-subtitle">TCG633 — Bridge &amp; Road Maintenance</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    pci_weight = st.slider("PCI weighting in Hybrid Index", 0.0, 1.0, 0.6, 0.05,
                            help="Hybrid Index = (this weight × PCI) + (remainder × IRI-derived score). "
                                 "PCI captures surface distress; IRI captures ride roughness.")
    st.caption(
        "**Limitations:** PCI uses a simplified single-defect deduct value "
        "(area × defect weighting × severity factor) rather than the full ASTM D6433 "
        "corrected-deduct-value curve. IRI is converted onto a comparable 0–100 scale "
        "using a piecewise-linear mapping aligned to JKR roughness bands. Suitable for "
        "network-level screening, not project-level certification."
    )

# --------------------------------------------------------------------------------------
# Upload existing Excel  /  Export current data
# --------------------------------------------------------------------------------------
st.markdown("## 📁 Upload Existing Pavement Excel")
up_col1, up_col2 = st.columns([3, 1])
with up_col1:
    upload = st.file_uploader("Upload Excel exported from this system", type=["xlsx", "csv"],
                               label_visibility="collapsed")
    if upload is not None:
        try:
            new_df = pd.read_excel(upload) if upload.name.endswith("xlsx") else pd.read_csv(upload)
            required = {"Section ID", "Defect Type", "Severity", "Area Affected (%)", "IRI (m/km)"}
            if not required.issubset(new_df.columns):
                st.error(f"File must contain columns: {', '.join(sorted(required))}")
            else:
                by_id = {r["Section ID"]: r for r in st.session_state.records}
                for _, row in new_df.iterrows():
                    rec = {k: row[k] for k in required}
                    by_id[rec["Section ID"]] = rec
                st.session_state.records = list(by_id.values())
                st.session_state.last_section = new_df.iloc[-1]["Section ID"]
                st.success(f"Imported {len(new_df)} section(s).")
                st.rerun()
        except Exception as e:
            st.error(f"Could not read file: {e}")
with up_col2:
    df_export = compute_all(st.session_state.records, pci_weight)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_export.drop(columns=["Condition Color"]).to_excel(writer, sheet_name="Sections", index=False)
    st.write("")
    st.download_button("⬇️ Export Data (Excel)", data=buf.getvalue(),
                        file_name="pavement_sections.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width='stretch')

# --------------------------------------------------------------------------------------
# Add / update a section
# --------------------------------------------------------------------------------------
st.markdown("## 🚧 Add New Pavement Section")
with st.form("add_section_form", clear_on_submit=False):
    c1, c2 = st.columns(2)
    with c1:
        sec_id = st.text_input("Section ID (e.g. S1, S2, S10)", value="S1")
        defect = st.selectbox("Defect Type", list(DEFECT_WEIGHTS.keys()))
        severity = st.selectbox("Severity", list(SEVERITY_FACTORS.keys()))
    with c2:
        area = st.number_input("Area Affected (%)", min_value=0.0, max_value=100.0, value=20.0, step=0.5)
        iri = st.number_input("IRI (m/km)", min_value=0.0, max_value=20.0, value=3.0, step=0.1)
    bcol1, bcol2 = st.columns([1, 1])
    submitted = bcol1.form_submit_button("➕ Add Section", width='stretch')
    reset_clicked = bcol2.form_submit_button("🔄 Reset All Data", width='stretch')

if submitted:
    sid = sec_id.strip() or "S1"
    by_id = {r["Section ID"]: r for r in st.session_state.records}
    by_id[sid] = {"Section ID": sid, "Defect Type": defect, "Severity": severity,
                   "Area Affected (%)": area, "IRI (m/km)": iri}
    st.session_state.records = list(by_id.values())
    st.session_state.last_section = sid
    st.success(f"Section {sid} saved.")
    st.rerun()

if reset_clicked:
    st.session_state.records = [dict(r) for r in SEED]
    st.session_state.last_section = "S1"
    st.success("All data reset to the sample dataset.")
    st.rerun()

# --------------------------------------------------------------------------------------
# Current section result
# --------------------------------------------------------------------------------------
results_df = compute_all(st.session_state.records, pci_weight)

st.markdown("## Current Section Result")
if st.session_state.last_section in results_df["Section ID"].values:
    row = results_df[results_df["Section ID"] == st.session_state.last_section].iloc[0]
else:
    row = results_df.iloc[0]

cA, cB = st.columns(2)
with cA:
    st.markdown(f'<div class="pms-label">PCI — Section {row["Section ID"]}</div>'
                f'<div class="pms-bignum">{row["PCI"]:.1f}</div>', unsafe_allow_html=True)
with cB:
    st.markdown(f'<div class="pms-label">Hybrid Index</div>'
                f'<div class="pms-bignum">{row["Hybrid Index"]:.1f}</div>', unsafe_allow_html=True)

st.markdown(
    f'<div class="pms-banner">Final Condition: <b>{row["Final Condition"]}</b> '
    f'&nbsp;|&nbsp; Maintenance: <b>{row["Maintenance Action"]}</b></div>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<div class="pms-gauge-wrap"><div class="pms-gauge-track">'
    f'<div class="pms-gauge-dot" style="left:{max(0,min(100,row["Hybrid Index"]))}%; '
    f'--dotcolor:{row["Condition Color"]};"></div></div></div>',
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)

# --------------------------------------------------------------------------------------
# Search & filter
# --------------------------------------------------------------------------------------
st.markdown("## 🔍 Search &amp; Filter")
f1, f2, f3 = st.columns([2, 1, 1])
with f1:
    search_text = st.text_input("Search Section", placeholder="e.g. S9")
with f2:
    defect_opt = ["All"] + sorted(results_df["Defect Type"].unique().tolist())
    defect_filter = st.selectbox("Defect", defect_opt)
with f3:
    severity_filter = st.selectbox("Severity", ["All", "Low", "Medium", "High"])

filtered = results_df.copy()
if search_text:
    filtered = filtered[filtered["Section ID"].str.contains(search_text, case=False, na=False)]
if defect_filter != "All":
    filtered = filtered[filtered["Defect Type"] == defect_filter]
if severity_filter != "All":
    filtered = filtered[filtered["Severity"] == severity_filter]

st.dataframe(
    filtered.drop(columns=["Condition Color"]).sort_values("Section ID").reset_index(drop=True),
    width='stretch', hide_index=True,
)

# --------------------------------------------------------------------------------------
# Graphical analysis
# --------------------------------------------------------------------------------------
st.markdown("## 📊 Graphical Analysis")
chart_df = results_df.sort_values("Section ID")
fig = make_subplots(specs=[[{"secondary_y": True}]])
fig.add_trace(go.Bar(x=chart_df["Section ID"], y=chart_df["PCI"], name="PCI",
                      marker_color="#5DADE2"), secondary_y=False)
fig.add_trace(go.Bar(x=chart_df["Section ID"], y=chart_df["Hybrid Index"], name="Hybrid Index",
                      marker_color="#A569BD"), secondary_y=False)
fig.add_trace(go.Scatter(x=chart_df["Section ID"], y=chart_df["IRI (m/km)"], name="IRI (m/km)",
                          mode="lines+markers", marker_color="#F1C40F"), secondary_y=True)
fig.update_layout(template="plotly_dark", barmode="group", height=420,
                   paper_bgcolor="#0e1117", plot_bgcolor="#0e1117",
                   legend=dict(orientation="h", y=1.1))
fig.update_yaxes(title_text="PCI / Hybrid Index (0–100)", secondary_y=False)
fig.update_yaxes(title_text="IRI (m/km)", secondary_y=True)
st.plotly_chart(fig, width='stretch')

# --------------------------------------------------------------------------------------
# Extra: network map, prioritisation, report (kept for engineering-depth / rubric marks)
# --------------------------------------------------------------------------------------
with st.expander("🗺️ Network Condition Map & Maintenance Priority"):
    chainage = 0
    section_len = 100
    fig_map = go.Figure()
    for _, r in chart_df.iterrows():
        fig_map.add_trace(go.Bar(x=[section_len], y=["Road"], base=[chainage], orientation="h",
                                  marker_color=r["Condition Color"], name=r["Section ID"],
                                  hovertext=f"{r['Section ID']} — {r['Final Condition']}", showlegend=False))
        chainage += section_len
    fig_map.update_layout(template="plotly_dark", barmode="stack", height=180,
                           paper_bgcolor="#0e1117", plot_bgcolor="#0e1117",
                           xaxis_title="Chainage (m)", yaxis_visible=False,
                           margin=dict(l=10, r=10, t=10, b=30))
    st.plotly_chart(fig_map, width='stretch')

    priority = chart_df.sort_values("Hybrid Index").reset_index(drop=True)
    priority.index = priority.index + 1
    priority.index.name = "Priority"
    st.caption("Priority 1 = most urgent (lowest Hybrid Index).")
    st.dataframe(priority[["Section ID", "Hybrid Index", "Final Condition", "Maintenance Action"]],
                 width='stretch')

with st.expander("📄 Auto-generated Maintenance Report"):
    lines = [
        "DIGITAL PAVEMENT MANAGEMENT SYSTEM — SUMMARY REPORT",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"Sections: {len(chart_df)}  |  Network avg PCI: {chart_df['PCI'].mean():.1f}  |  "
        f"Network avg Hybrid Index: {chart_df['Hybrid Index'].mean():.1f}",
        "-" * 60,
    ]
    for _, r in chart_df.iterrows():
        lines.append(f"{r['Section ID']}: PCI={r['PCI']:.1f}, IRI={r['IRI (m/km)']:.2f} -> "
                      f"{r['Final Condition']} | Action: {r['Maintenance Action']}")
    report_text = "\n".join(lines)
    st.text_area("Preview", report_text, height=240)
    st.download_button("⬇️ Download report (TXT)", data=report_text.encode("utf-8"),
                        file_name="maintenance_report.txt", mime="text/plain")

st.markdown(
    '<div style="text-align:center; color:#5b6577; padding-top:18px; font-size:0.85rem;">'
    "TCG633 Bridge &amp; Road Maintenance — Individual Project</div>",
    unsafe_allow_html=True,
)
