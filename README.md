[README.md](https://github.com/user-attachments/files/29504084/README.md)
# 🛣️ Digital Pavement Condition Evaluation and Maintenance Decision Tool

**TCG633 Bridge & Road Maintenance — Individual Project (40%)**
Fakulti Kejuruteraan Awam, Universiti Teknologi MARA, Cawangan Sarawak

A Streamlit web application that evaluates road pavement condition from
defect and roughness data, computes **PCI** (Pavement Condition Index) and
**IRI** (International Roughness Index), classifies each section
(*Very Good / Good / Fair / Poor*), and recommends a maintenance action —
satisfying Part A (Digital Tool Development) of the project brief.

---

## 1. What this app does

The app is now a true **multi-page navigation app** (sidebar page picker, not
tabs) — only the page you're viewing actually runs, which keeps it snappy
even with the bonus pages added.

- ✅ Upload pavement data as **CSV or Excel** — including the lecturer's own
  multi-sheet `TCG633_PCI_IRI_Pro` template directly, no reformatting needed
- ✅ Accepts the required columns: `Section`, `Defect Type`, `Severity`, `Area Percentage (%)`, `IRI`
- ✅ Computes **PCI** using a deduct-value method (defect weighting × severity factor × area %)
- ✅ Computes **IRI** as the average roughness per section
- ✅ Classifies each section: **Very Good / Good / Fair / Poor**
- ✅ Produces a **combined (hybrid) PCI + IRI rating** — the worse of the two governs (official, lecturer's rule)
- ✅ Analyses defects **by section** (counts, types, severities)
- ✅ Recommends a maintenance action for **each section** (official, classification-based)
  and a supplementary general action for **each defect**
- ✅ Summary dashboard with KPI cards (sections, average PCI, average IRI, poor-section count, overall condition)
- ✅ Interactive charts: PCI by section, IRI by section, defect distribution, condition distribution, defects-by-section
- ✅ **Download** analysed results as Excel (3 sheets) or CSV
- ✅ Built-in **"How to Use"** and **"Methodology & Assumptions"** pages inside the app

**Bonus features (per the project brief's optional extension list):**

- 🧮 **Hybrid Index page** — blends PCI and IRI into a single adjustable 0-100
  number per section (slider-controlled weighting), in addition to the
  official worse-of rating used everywhere else.
- 🗺️ **GIS Map page** — plots sections on an interactive map using
  Streamlit's native `st.map`. Since the dataset has no real GPS data, this
  uses clearly-labelled **simulated** coordinates (you choose a start point,
  bearing, and spacing) — swap in real survey coordinates for production use.
- 📄 **Report Generator page** — one click produces a self-contained,
  downloadable HTML report (KPIs, charts, tables) with zero extra PDF
  dependencies — open it in any browser and use Print → Save as PDF.

The full computation logic (weighting factors, severity factors, PCI/IRI
condition bands, and the hybrid combination rule) is taken **directly** from
the lecturer-provided workbooks `TCG633_PCI_IRI_Model.xlsx` and
`TCG633_PCI_IRI_Pro_v2.xlsx`. See the in-app **Methodology & Assumptions**
page for the full breakdown and every assumption made explicit, including how
the bonus features relate to (and don't replace) the official rating.

---

## 2. Project files

```
.
├── app.py              # Main Streamlit application (single file)
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

---

## 3. Quick start (run on localhost)

### Option A — using Git + GitHub

```bash
# 1. Clone (or create) your repository
git clone https://github.com/<your-username>/<your-repo-name>.git
cd <your-repo-name>

# 2. Copy app.py and requirements.txt into this folder, then:
git add app.py requirements.txt README.md
git commit -m "Add pavement condition evaluation tool"
git push origin main
```

### Option B — run locally straight away

```bash
# 1. (Recommended) create a virtual environment
python -m venv venv

# Activate it:
#   Windows:
venv\Scripts\activate
#   macOS / Linux:
source venv/bin/activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
streamlit run app.py
```

Streamlit will print a local URL, typically:

```
Local URL: http://localhost:8501
```

Open that link in your browser — the app runs entirely on your machine, no
internet connection or external server required after the initial `pip
install`.

> If you cloned from GitHub, run `pip install -r requirements.txt` then
> `streamlit run app.py` from inside the cloned folder — same steps as
> Option B.

---

## 4. Preparing your data

Build a spreadsheet or CSV with **one row per observed defect** (a section
can repeat across multiple rows) using exactly these column headers (or a
close variant — the app recognises common aliases such as `Section ID`,
`Area (%)`, `IRI (m/km)`, etc.):

| Section | Defect Type | Severity | Area Percentage (%) | IRI |
|---|---|---|---|---|
| S1 | Potholes | High | 6 | 3.8 |
| S1 | Raveling | Low | 10 | 3.8 |
| S2 | Longitudinal Crack | Low | 4 | 1.9 |

- **Severity** must be `Low`, `Medium`, or `High`.
- **Defect Type** should be one of: `Longitudinal Crack`, `Alligator (Fatigue) Crack`,
  `Potholes`, `Raveling`, `Depression/Sag`, `Patching (Failed)`,
  `Bleeding/Flushing`, `Rut/Rutting` (matches the lecturer's Lookup table).
  Unrecognised types are still processed using a flagged default factor.
- **IRI** (m/km) can repeat on every row for a section, or be supplied once —
  the app averages whatever values it finds per section.
- You don't need both PCI and IRI data — the tool will compute whichever
  indicator(s) the data supports (PCI-only, IRI-only, or hybrid).

Don't have data ready? Click **"📊 Load Sample Data"** in the sidebar to try
the tool instantly with a built-in demo dataset spanning all four condition
classes.

---

## 5. How the indicators are calculated (summary)

**PCI (per defect):**
`Deduct Value = Area (%) × Severity Factor × Weighting Factor`
Then per section: `PCI = MAX(0, 100 − MIN(100, ΣDeduct Value))`

**IRI (per section):** simple average of all IRI readings for that section.

**Classification bands** (PCI / IRI) and the **hybrid combined rating**
(worse of PCI vs IRI governs) exactly mirror the lecturer's
`TCG633_PCI_IRI_Model.xlsx` / `TCG633_PCI_IRI_Pro_v2.xlsx` Lookup tables.

Full details, all formulas, and every assumption are documented inside the
running app under **"📐 Methodology & Assumptions"** — use this section as a
direct reference for your technical report.

---

## 6. Suggested use for your TCG633 deliverables

| Deliverable | How this app helps |
|---|---|
| Working Digital Analysis Tool | `app.py` itself (Part A, PO5) |
| Technical Report | Copy the Methodology tab text + exported tables/charts |
| Dataset Used | Export your uploaded data via the Excel download (Raw_Input sheet), or submit the file you uploaded |
| Video Presentation | Screen-record yourself uploading data → Dashboard → Results → Charts → Download (Part B, PO9) |

---

## 7. Notes & limitations

- This is a **teaching simplification** of pavement evaluation practice, not
  the full ASTM D6433 deduct-curve procedure — stated explicitly in-app.
- Per-defect "Suggested Defect Treatment" text is **supplementary general
  guidance** added by this tool, separate from the official section-level,
  classification-based recommendation that follows the lecturer's model.
- GIS mapping, automated PDF report generation, and a PCI/IRI hybrid index
  toggle remain available as **optional bonus** extensions per the project
  brief and are not implemented here.

---

## 8. Troubleshooting

- **"streamlit: command not found"** → activate your virtual environment, or
  run `python -m streamlit run app.py`.
- **Port already in use** → run `streamlit run app.py --server.port 8502`.
- **Excel upload fails** → ensure the file isn't open in another program and
  has a single header row.

### Deploying on Streamlit Community Cloud — `ModuleNotFoundError: plotly` (or similar)

This means Streamlit Cloud built your app **before** it could see
`requirements.txt`, or the file isn't where it expects it. Fix it like this:

1. In your **GitHub repository**, make sure `requirements.txt` sits in the
   **exact same folder** as `app.py` (the repo root, unless you set a
   different "Main file path" when deploying — in that case `requirements.txt`
   must be in that same folder).
2. Open the file on GitHub and confirm it actually contains `plotly` (and the
   other packages) — a failed/incomplete upload is a common cause.
3. On Streamlit Cloud, open your app → bottom-right **"Manage app"** →
   **⋮ menu → Reboot app** (this forces a fresh `pip install -r
   requirements.txt`). A plain page refresh does **not** reinstall packages.
4. Still failing? Open **"Manage app" → logs** and read the actual `pip
   install` error above the redacted message shown on the app page — it
   usually points to the real cause (wrong file path, a typo in a package
   name, or a version that has no pre-built wheel for the platform).

**This app is resilient to this exact failure mode:** if Plotly genuinely
isn't available, `app.py` detects this at startup and automatically falls
back to Streamlit's built-in `st.bar_chart` for every chart instead of
crashing — so a missing/late-installing dependency will degrade the visuals,
not break the whole app. Once `requirements.txt` is fixed and the app is
rebooted, the full interactive Plotly charts return automatically.
