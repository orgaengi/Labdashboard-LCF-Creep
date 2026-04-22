# Lab Planning Dashboard — Web App
## Streamlit Edition

---

## Run Locally (Fastest Start)

### Windows
1. Double-click **`RUN_WEB_APP.bat`**
2. Browser opens automatically at **http://localhost:8501**
3. Press `Ctrl+C` in the terminal to stop

### macOS / Linux
```bash
chmod +x run_web_app.sh
./run_web_app.sh
```

### Manual
```bash
pip install streamlit plotly pandas openpyxl xlrd
streamlit run app.py
```

---

## Deploy Free Online (Streamlit Cloud)

### Step 1 — Create accounts (both free)
- GitHub: https://github.com/signup
- Streamlit Cloud: https://share.streamlit.io

### Step 2 — Push to GitHub
```bash
# In this folder:
git init
git add .
git commit -m "Lab Dashboard Web App"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/lab-dashboard.git
git push -u origin main
```

### Step 3 — Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io
2. Click **"New app"**
3. Select your GitHub repo
4. Set **Main file path** = `app.py`
5. Click **"Deploy!"**
6. Get your public URL: `https://YOUR_USERNAME-lab-dashboard.streamlit.app`

---

## Files in This Package

| File | Purpose |
|------|---------|
| `app.py` | Main Streamlit web app |
| `generators.py` | Excel generation (no GUI dependency) |
| `lab_core.py` | Shared data processing engine |
| `tool1_lcf_creep.py` | Tool 1 logic |
| `tool2_coating_labs.py` | Tool 2 logic |
| `tool3_comparison.py` | Tool 3 logic + merge_files |
| `requirements.txt` | Python package dependencies |
| `.streamlit/config.toml` | App theme and settings |
| `RUN_WEB_APP.bat` | Windows launcher |
| `run_web_app.sh` | macOS/Linux launcher |
| `real_lab_data.xlsx` | Sample data from your images |

---

## Input Excel Format

Your Excel file must have these columns (names are flexible):

| Column | Accepted Names | Example |
|--------|---------------|---------|
| Year | Year, Yr, FY | 2024 |
| Type | Type, Process, Lab, Category | LCF |
| Value | Value, Count, Samples, Volume | 59 |

**Lab type names accepted:**
- Mechanical: `LCF`, `Creep`
- Coating: `Cold Spray`, `HVOF`, `Plasma`

---

## Tool Summary

| Tool | Labs | Capacity |
|------|------|---------|
| Tool 1 | LCF + Creep | LCF=50, Creep=22 (editable) |
| Tool 2 | Cold Spray + HVOF + Plasma | 350 combined (editable, with individual split option) |
| Tool 3 | All 5 labs | Both groups compared side-by-side |

---

## Web App Features

- **Interactive Plotly charts** (zoom, hover, export PNG)
- **YoY Pie Charts** — process-wise demand share per year
- **Gantt Heatmap** — weekly occupancy for current year
- **Capacity warning** — alerts when individual caps exceed combined
- **Multi-file upload** in Tool 3 — upload separate files per lab group
- **Excel download** — full dashboard with all sheets including pie charts
- **4 color themes** for utilization coloring
