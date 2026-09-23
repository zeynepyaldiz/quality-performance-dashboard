# Quality Performance & Defect Monitoring Dashboard

An interactive Streamlit dashboard for manufacturing quality/defect data, built from the
`TEKNOCEF_GÜNLÜK_GİDİŞAT-TEMMUZ_2026.xlsm` daily quality-control export.

Built with **Streamlit + Python + Pandas + Plotly**.

---

## 1. What it does

The app automatically loads the bundled source file, cleans it, and renders a full
quality-monitoring dashboard:

1. **KPI cards** — Total Inspections, Total Defects, Overall NOK %, number of References,
   number of Defect Types.
2. **Pareto Analysis** — defect-quantity bars + cumulative % line + 80% reference line,
   with critical defects highlighted **dynamically** (whichever defects reach ~80% of the
   total in the *currently filtered* data — not hard-coded).
3. **Reference Analysis** — Reference vs NOK %, Reference × Defect-Type heatmap, and the
   Top 10 Reference–Defect combinations.
4. **Time Analysis** — daily defect trend with a rolling average.
5. **Critical Defects** — table of the defects contributing to ≈80% of total defects and
   their main affected references.
6. **Interactive filters** — date range, reference, defect type, shift (auto-hidden if no
   shift column exists in the source file), and a one-click **Reset filters** button.
7. **Detailed filtered data** — a live table of the filtered records with **CSV** and
   **Excel** download buttons.
8. **Key Insights** — short, automatically generated, data-driven observations (no
   invented causes or conclusions — only what the numbers show).

Nothing is hard-coded: every figure, chart and insight is recomputed from whatever data is
currently loaded and currently filtered. You can also upload a newer/updated export file
from the sidebar and the whole dashboard recalculates from it.

---

## 2. Data cleaning performed automatically on load

The app reads the raw export (wide format: one row per Reference × Date, with one column
per defect type) and, every time it loads a file, performs:

- **Column-header cleanup** — strips stray whitespace (e.g. `"KAPLAMA "` → `"KAPLAMA"`),
  drops empty/`Unnamed:` template columns.
- **Blank/template row removal** — rows with no `REFERANS` and no `TARİH` (these exist at
  the bottom of the sheet as leftover template rows) are dropped.
- **Reference-code normalization** — trims stray spaces, standardizes dash spacing
  (`"456 - B"` → `"456-B"`), unifies int/str variants (`456` vs `"456"`), and upper-cases
  codes, without guessing at genuine typos in the underlying code.
- **Missing/invalid value handling**:
  - Defect-type cells left blank are treated as `0` occurrences.
  - Negative counts (a data-entry error present in the source, e.g. `OK = -35`) are
    clipped to `0`.
  - Missing `KONTROL EDİLEN` (inspected quantity) values are recomputed as `OK + NOK`
    where possible.
- **Automatic column classification** — the app locates the fixed structural columns
  (`REFERANS`, `TARİH`, `HAFTA`, `KONTROL EDİLEN`, `TASHİH YAPILACAK`, `OK`, `NOK`, `FTT`)
  by normalized name matching (so header case/whitespace changes don't break it), and
  treats **every other numeric column as a defect type** automatically — so if the source
  file gains or loses defect columns, the dashboard adapts without code changes.
- **Shift detection** — the app looks for a shift/vardiya-type column; the current source
  file does not contain one, so the Shift filter is automatically hidden with a note
  instead of being shown empty.

A **"Data-cleaning log"** panel in the sidebar lists exactly what was cleaned for full
transparency (useful to show during a presentation).

### Validation

The six known reference totals from the prior manual analysis were used only to *validate*
the cleaning pipeline (never hard-coded into the app) and reproduce exactly from the raw
file:

| Defect Type    | Expected | Reproduced by pipeline |
|----------------|---------:|------------------------:|
| DETAY HATASI   | 6,205    | 6,205 |
| TORK HATASI    | 4,351    | 4,351 |
| DEFORMASYON    | 2,844    | 2,844 |
| KAPLAMA        | 2,534    | 2,534 |
| CİVATA YAMUK   | 1,036    | 1,036 |
| DİŞ HATASI     | 744      | 744   |

---

## 3. Project structure

```
quality_dashboard/
├── app.py                 # the Streamlit application (single-file app)
├── requirements.txt       # Python dependencies
├── README.md              # this file
└── data/
    └── source_data.xlsm   # bundled source dataset (loaded by default)
```

You can replace `data/source_data.xlsm` with an updated export at any time, or simply use
the **"Upload an updated quality export"** control in the sidebar at runtime — no code
changes required either way.

---

## 4. How to run it

### Step 1 — Install dependencies (Python 3.9+ recommended)

```bash
pip install -r requirements.txt
```

### Step 2 — Launch the app

```bash
streamlit run app.py
```

Streamlit will print a local URL (typically `http://localhost:8501`) — open it in your
browser. The dashboard loads the bundled dataset automatically; no further setup is
needed.

### Optional: run on a specific port / make it reachable on your network

```bash
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

---

## 5. Notes for presentation / internship use

- The **Pareto chart** highlights (in red) whichever defects are dynamically found to make
  up ≈80% of total defects for the *currently selected* filters — change the date range or
  reference filter and watch the critical-defect set update live.
- The **Reference × Defect-Type heatmap** is capped to the top 20 references (by inspected
  volume) and top 15 defect types (by quantity) purely for readability; use the filters to
  drill into a smaller set if you need the full detail.
- The **Key Insights** section only states what is directly observable in the data
  (largest defect type, highest/lowest NOK% reference, worst day, etc.) — it deliberately
  does not speculate about root causes, in line with good quality-engineering practice.
- All charts and tables respond to the sidebar filters immediately; use **Reset filters**
  to return to the full dataset view.
