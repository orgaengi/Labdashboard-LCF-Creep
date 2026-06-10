# 🔬 Lab Occupancy Dashboard

A four-tool Streamlit dashboard for lab sample demand, capacity utilization, and vacancy planning.

## Tools
| Tool | Labs | Output |
|------|------|--------|
| 🔵 Tool 1 — LCF & Creep | LCF, Creep | 8-sheet Excel |
| 🟢 Tool 2 — Coating Labs | Cold Spray, HVOF, Plasma | 8-sheet Excel |
| 🟠 Tool 3 — Thermal Lab | Thermal Rig | 8-sheet Excel |
| 🔴 Tool 4 — Comparison & PPT | All 3 groups | 10-sheet Excel + 6-slide PPT |

## Deploy on Streamlit Community Cloud
1. Push these files to a GitHub repo
2. Go to share.streamlit.io → New app → select repo → **app.py**
3. Done — your URL: `https://username-appname.streamlit.app`

## Run Locally
```
pip install -r requirements.txt
streamlit run app.py
```

## Input Formats Supported
- Long: `Year | Type | Value`
- Wide: `Year | LabType1 | LabType2 | …`
- Monthly block: year-named sheets with TOTAL SAMPLES REMOVED blocks
- Messy data: junk rows, mixed case, N/A values, data on any sheet

## Dependencies
`streamlit · pandas · openpyxl · plotly · kaleido==0.2.1 · python-pptx · xlrd`
