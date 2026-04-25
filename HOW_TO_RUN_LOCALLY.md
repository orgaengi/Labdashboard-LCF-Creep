# How to Run the Lab Dashboard Locally

## Prerequisites
- Python 3.9 or later  (check: `python --version`)
- pip (comes with Python)

---

## Step 1 — Unzip the project

Unzip `lab_web_app_FIXED.zip` anywhere, e.g. `C:\Users\YourName\lab_web\`

You should see: `app.py`, `lab_core.py`, `generators.py`, `requirements.txt`, `.streamlit/`

---

## Step 2 — Open a terminal in that folder

**Windows:** Command Prompt / PowerShell → `cd C:\Users\YourName\lab_web`
**Mac/Linux:** Terminal → `cd ~/lab_web`

---

## Step 3 — Create a virtual environment (recommended)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Mac / Linux
python3 -m venv venv
source venv/bin/activate
```
You will see `(venv)` at the start of your prompt.

---

## Step 4 — Install dependencies

```bash
pip install -r requirements.txt
```
Installs: streamlit, pandas, openpyxl, plotly, anthropic, xlrd.

---

## Step 5 — Set your Anthropic API key (for AI Chat)

### Option A — secrets.toml (easiest)
```bash
# copy the example
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # Mac/Linux
copy .streamlit\secrets.toml.example .streamlit\secrets.toml  # Windows
```
Open `.streamlit/secrets.toml` and set:
```toml
ANTHROPIC_API_KEY = "sk-ant-your-key-here"
```
Get a key at: https://console.anthropic.com/keys

### Option B — environment variable
```bash
export ANTHROPIC_API_KEY="sk-ant-your-key-here"   # Mac/Linux
$env:ANTHROPIC_API_KEY = "sk-ant-your-key-here"   # Windows PowerShell
```

> Without a key the dashboard works fully — only the AI Chat sidebar will show an error.

---

## Step 6 — Run the app

```bash
streamlit run app.py
```

Browser opens at **http://localhost:8501**

---

## Step 7 — Using the AI Chat

The 🤖 AI Lab Assistant is in the **sidebar** (scroll down).
- Ask about utilization: *"What is Cold Spray utilization in 2024?"*
- Ask for recommendations: *"What capacity do I need for <80% utilization?"*
- Ask about trends: *"Which year had the highest Plasma demand?"*

**What it can do:**
- Answer questions about the data currently loaded
- Suggest capacity adjustments with concrete numbers
- Explain utilization trends and seasonal patterns

**What it cannot do (yet):**
- Automatically change the capacity input fields on your behalf
  (you must apply suggested values yourself in the UI)
- Read the Excel file contents directly — it uses the summary data
  that is populated when you upload a file

---

## Deploying to Streamlit Cloud

1. Push the project folder to a **GitHub repo** (private is fine)
2. Go to https://share.streamlit.io → "New app"
3. Select your repo, branch, set main file = `app.py`
4. **Advanced settings → Secrets** → paste:
   ```
   ANTHROPIC_API_KEY = "sk-ant-your-key-here"
   ```
5. Click **Deploy**

> ⚠️ Never commit `.streamlit/secrets.toml` to GitHub.
> The `.gitignore` already excludes it.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError: streamlit` | `pip install -r requirements.txt` |
| `ModuleNotFoundError: anthropic` | `pip install anthropic>=0.25.0` |
| Port 8501 busy | `streamlit run app.py --server.port 8502` |
| AI chat: "API key not set" | Follow Step 5 above |
| Excel not recognised | Ensure column names contain Year / Type / Value |
