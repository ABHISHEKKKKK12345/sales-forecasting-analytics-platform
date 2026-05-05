<div align="center">

<h1>📈 Sales Forecasting &amp; Automated Reporting System</h1>

<p><strong>Production-grade demand forecasting engine with a dark-themed desktop GUI,<br>
ML ensemble modelling, statistical rigour, and one-click PDF / Excel report generation.</strong></p>

<p>
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/Python-3.9%2B-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.9+"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-22863A?style=for-the-badge" alt="License"></a>
  <a href="https://pyinstaller.org/en/stable/"><img src="https://img.shields.io/badge/Build-PyInstaller-E8730A?style=for-the-badge" alt="PyInstaller"></a>
  <a href="https://docs.python.org/3/library/tkinter.html"><img src="https://img.shields.io/badge/GUI-tkinter%20Dark%20Theme-161B22?style=for-the-badge" alt="tkinter"></a>
  <a href="#building-a-standalone-executable"><img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-6e7681?style=for-the-badge" alt="Platform"></a>
</p>

<p>
  <a href="#overview">Overview</a> &nbsp;•&nbsp;
  <a href="#versions-at-a-glance">Versions</a> &nbsp;•&nbsp;
  <a href="#features">Features</a> &nbsp;•&nbsp;
  <a href="#installation">Install</a> &nbsp;•&nbsp;
  <a href="#running-the-application">Run</a> &nbsp;•&nbsp;
  <a href="#building-a-standalone-executable">Build</a> &nbsp;•&nbsp;
  <a href="#troubleshooting">Troubleshoot</a>
</p>

</div>

---

## 📑 Table of Contents

- [Overview](#overview)
- [Repository Structure](#repository-structure)
- [Versions at a Glance](#versions-at-a-glance)
- [Features](#features)
  - [Forecasting Engine](#forecasting-engine)
  - [Statistical Rigour](#statistical-rigour)
  - [Business Analytics](#business-analytics)
  - [GUI](#gui)
  - [Output](#output)
- [Requirements](#requirements)
  - [Core Python Packages](#core-python-packages)
  - [Additional Packages — Experimental Only](#additional-packages--experimental-only)
  - [tkinter Platform Notes](#tkinter-platform-notes)
- [Installation](#installation)
- [Running the Application](#running-the-application)
  - [Quick Start with Demo Data](#quick-start-with-demo-data)
- [Data Format](#data-format)
  - [Required Columns](#required-columns)
  - [Optional Columns](#optional-columns-auto-filled-with-defaults-if-absent)
  - [Supported File Formats](#supported-file-formats)
- [Forecasting Models](#forecasting-models)
- [Configuration](#configuration)
  - [GUI Settings](#gui-settings-all-versions)
  - [TOML Config — Experimental](#toml-persistent-configuration-experimental-only)
- [Report Output](#report-output)
  - [PDF Report Sections](#pdf-report-sections)
  - [Excel Workbook Sheets](#excel-workbook-sheets)
  - [JSON Summary](#json-summary)
- [Building a Standalone Executable](#building-a-standalone-executable)
  - [Windows (.exe)](#windows-exe)
  - [macOS (.app)](#macos-app)
  - [Linux (binary)](#linux-binary)
- [Version Comparison](#version-comparison)
- [Troubleshooting](#troubleshooting)
- [Author](#author)

---

## Overview

This system delivers a complete **sales forecasting and business analytics pipeline** in self-contained Python files. Load your own data (CSV, Excel, JSON, Parquet) or use the built-in **60-month demo dataset**, run up to **nine forecasting models in parallel**, and generate publication-quality **PDF reports** and **multi-sheet Excel workbooks** — all from a polished desktop GUI with **no internet connection required**.

Three source files are included, each representing a stage of development:

| File | Version | Purpose |
|---|---|---|
| [`sales_forecast_gui_v1.py`](src/sales_forecast_gui_v1.py) | v1.0 | Original full-featured release |
| [`sales_forecast_gui_v2.py`](src/sales_forecast_gui_v2.py) | v2.0 — ✅ **Recommended** | Hardened, production-ready |
| [`forecasting_experimental.py`](src/forecasting_experimental.py) | Experimental 🧪 | Extended model suite + TOML config |

---

## Repository Structure

```
your-repo/
│
├── README.md                          ← you are here
├── requirements.txt                   ← core pip dependencies (v1 / v2)
├── requirements_experimental.txt     ← extended deps (experimental build)
├── .gitignore
├── LICENSE
│
├── experiments/                           ← development & test builds (optional)
│   ├── forecasting_experimental_dev_v2.py
│   ├── sales_forecast_gui_dev_v4.py
│   ├── sales_forecast_gui_v1_dev_v5.py
│   └── sales_forecast_gui_v2_dev_v6.py
│
└── src/
    ├── sales_forecast_gui_v1.py           ← v1 — original release
    ├── sales_forecast_gui_v2.py           ← v2 — hardened (recommended)
    └── forecasting_experimental.py        ← experimental — extended model suite
```
> experiments/ contains development snapshots, intermediate builds, and test variants that are not intended for production use.
---

## Versions at a Glance

### `sales_forecast_gui_v1.py` — Version 1.0

> Original production release. Fully functional with **5 forecasting models**, PDF + Excel + JSON output, walk-forward cross-validation, and a complete dark-themed GUI. Suitable for personal use and exploration.

### `sales_forecast_gui_v2.py` — Version 2.0 &nbsp;✅ Recommended

> All v1 functionality plus hardened error handling, NaN guards, debug mode via `SALES_FORECAST_DEBUG=1`, improved timestamped logging, decomposition length guard, ML feature shape validation, and safer `mainloop` exit. **Use this for any shared or production deployment.**

### `forecasting_experimental.py` — Experimental &nbsp;🧪

> Extends v2 with three optional models — **XGBoost**, **LightGBM**, and **Prophet** — plus a **TOML-based configuration system** that persists all settings between runs. Optional models degrade gracefully if their packages are not installed. Designed for research, benchmarking, and advanced deployments.

---

## Features

### Forecasting Engine

| Model | v1 | v2 | Experimental |
|---|:---:|:---:|:---:|
| ARIMA (MLE grid search, AIC/BIC) | ✅ | ✅ | ✅ |
| Holt-Winters (additive / multiplicative, AIC) | ✅ | ✅ | ✅ |
| Random Forest (200 trees, PACF lags) | ✅ | ✅ | ✅ |
| Gradient Boosting (200 rounds, lr 0.05) | ✅ | ✅ | ✅ |
| Ridge Regression (L2 baseline) | ✅ | ✅ | ✅ |
| Stacked Ensemble (inv-sMAPE weighted) | ✅ | ✅ | ✅ |
| XGBoost | ✗ | ✗ | ✅ optional |
| LightGBM | ✗ | ✗ | ✅ optional |
| Prophet | ✗ | ✗ | ✅ optional |

### Statistical Rigour

- **Metrics** — sMAPE · MASE · RMSE · MAE · R²
- **Confidence Intervals** — ARIMA ψ-weight propagation · Holt-Winters analytical / bootstrap · ML non-parametric bootstrap
- **Diagnostics** — Ljung-Box autocorrelation · Shapiro-Wilk normality · Durbin-Watson serial correlation · heteroscedasticity check
- **Cross-validation** — expanding-window walk-forward CV (zero look-ahead bias)
- **Anomaly detection** — 6-month rolling Z-score, threshold 2.5σ
- **Marketing ROI** — 95% bootstrap confidence interval
- **Seasonal period** — auto-detected via periodogram

### Business Analytics

- CAGR · YoY growth · MoM growth · revenue volatility (CV)
- Peak / trough detection · return rate analysis
- Product-level and region-level breakdowns

### GUI

- Dark GitHub-inspired theme (full `ttk` styling, background `#0D1117`)
- Three-tab layout: **Setup → Log → Results**
- Real-time progress bar and colour-coded activity log
- KPI dashboard panel on Results tab
- Sortable model comparison table
- One-click open for all generated files
- Cross-platform mousewheel scrolling

### Output

- **PDF report** — multi-page A4, branded header/footer, 8 embedded charts, 7 sections
- **Excel workbook** — 9 sheets, frozen panes, Excel Tables, conditional formatting, embedded chart images
- **JSON summary** — machine-readable metrics, ensemble weights, CI method descriptions, platform info

---

## Requirements

- **[Python 3.9+](https://www.python.org/downloads/)**
- **[tkinter](https://docs.python.org/3/library/tkinter.html)** — usually bundled with Python (see [tkinter Platform Notes](#tkinter-platform-notes))

### Core Python Packages

```text
numpy>=1.24
pandas>=1.5
matplotlib>=3.6
seaborn>=0.12
scipy>=1.10
scikit-learn>=1.2
openpyxl>=3.1
reportlab>=4.0
pillow>=9.0
```

All listed in [`requirements.txt`](requirements.txt).

### Additional Packages — Experimental Only

```text
xgboost>=1.7       # optional — XGBoost model
lightgbm>=3.3      # optional — LightGBM model
prophet>=1.1       # optional — Prophet model
tomli>=2.0         # required on Python < 3.11 (TOML config read)
tomli-w>=1.0       # required for TOML config write
```

All listed in [`requirements_experimental.txt`](requirements_experimental.txt).

> **Note:** XGBoost, LightGBM, and Prophet are checked at runtime. If absent, the corresponding model is silently disabled — all other models continue to run normally.

### tkinter Platform Notes

| Platform | Fix if tkinter is missing |
|---|---|
| **Windows** | Reinstall [Python](https://www.python.org/downloads/) and tick **"tcl/tk and IDLE"** during setup |
| **macOS** | `brew install python-tk` &nbsp; ([Homebrew](https://brew.sh/)) |
| **Ubuntu / Debian** | `sudo apt install python3-tk` |
| **Fedora / RHEL** | `sudo dnf install python3-tkinter` |

---

## Installation

### 1 — Clone the repository

```bash
git clone https://github.com/your-username/your-repo.git
cd your-repo
```

### 2 — Create a virtual environment (strongly recommended)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS / Linux
python3 -m venv venv
source venv/bin/activate
```

### 3 — Install dependencies

**For v1 / v2:**

```bash
pip install -r requirements.txt
```

**For the experimental build** (adds XGBoost, LightGBM, Prophet, tomli):

```bash
pip install -r requirements.txt
pip install -r requirements_experimental.txt
```

**Or install optional models individually:**

```bash
pip install xgboost lightgbm prophet tomli tomli-w
```

> On first launch, if any **core** dependency is missing, the app shows a dialog with the exact `pip install` command needed before the main window opens.

---

## Running the Application

```bash
# v2 — recommended
python src/sales_forecast_gui_v2.py

# v1 — original
python src/sales_forecast_gui_v1.py

# Experimental — extended model suite
python src/forecasting_experimental.py

# Experimental — verbose debug logging
SALES_FORECAST_DEBUG=1 python src/forecasting_experimental.py        # macOS / Linux
set SALES_FORECAST_DEBUG=1 && python src/forecasting_experimental.py  # Windows CMD
$env:SALES_FORECAST_DEBUG=1; python src/forecasting_experimental.py   # PowerShell
```

### Quick Start with Demo Data

1. Launch any version of the app
2. Leave **"Use built-in demo data"** checked — 5 products · 60 months pre-generated
3. Enter your company name and desired forecast horizon
4. Click **🚀 Run Analysis & Generate Reports**
5. Choose a save directory when prompted
6. PDF and Excel reports open automatically when complete

---

## Data Format

### Required Columns

| Column | Type | Description |
|---|---|---|
| `date` | date string | Any parseable format: `YYYY-MM-DD`, `MM/DD/YYYY`, etc. |
| `sales` | numeric | Revenue or sales amount (non-negative) |

### Optional Columns (auto-filled with defaults if absent)

| Column | Default | Description |
|---|---|---|
| `product` | `"Default"` | Product or SKU name |
| `region` | `"Default"` | Sales region |
| `units` | `sales / 50` | Units sold |
| `returns` | `sales × 2%` | Returns value |
| `marketing_spend` | `sales × 10%` | Marketing expenditure |

### Supported File Formats

| Format | Extensions |
|---|---|
| CSV / TSV | `.csv` &nbsp; `.tsv` |
| Excel | `.xlsx` &nbsp; `.xlsm` &nbsp; `.xls` |
| JSON | `.json` — records, split, index, values, table orientations |
| Parquet | `.parquet` |

The loader automatically tries multiple encodings: `utf-8` → `latin-1` → `cp1252` → `iso-8859-15`.

**Example CSV:**

```csv
date,product,region,sales,units,returns,marketing_spend
2022-01-01,Widget A,North,125000,2100,1800,14000
2022-02-01,Widget A,North,118000,1980,1650,13500
2022-03-01,Widget B,South,89000,3100,890,9800
```

---

## Forecasting Models

**ARIMA** — Full Gaussian MLE via innovations form. Grid search over `p ∈ [0,3]`, `d ∈ [0,1]`, `q ∈ [0,2]`; winner selected by AIC. Confidence intervals via ψ-weight variance propagation (Box-Jenkins §5.4). Multi-start L-BFGS-B optimisation for reliable convergence.

**Holt-Winters** — Triple exponential smoothing (level + trend + seasonal). Both additive and multiplicative variants fitted; AIC selects the winner. Analytical CI for additive; bootstrap CI for multiplicative.

**Random Forest** — 200 trees, max depth 8, min samples/leaf 2. Features: PACF-guided lags, rolling mean / std / min / max (3/6/12 windows), harmonic seasonality encoding, polynomial trend terms. Recursive multi-step forecast strategy.

**Gradient Boosting** — 200 rounds, learning rate 0.05, max depth 4, subsample 0.8. Same feature engineering pipeline as Random Forest.

**Ridge Regression** — L2 regularisation (α = 1.0). Fast linear baseline using the same feature matrix.

**Stacked Ensemble** — Weights = inverse of walk-forward CV sMAPE per model. Falls back to uniform weights if CV scores are degenerate. NaN-guarded — always produces a finite, valid forecast.

**XGBoost** *(Experimental — optional)* — Gradient boosting via the [xgboost](https://xgboost.readthedocs.io/) library. Same feature matrix as RF / GB. Disabled gracefully if package is absent.

**LightGBM** *(Experimental — optional)* — Gradient boosting via the [lightgbm](https://lightgbm.readthedocs.io/) library. Same feature matrix. Disabled gracefully if package is absent.

**Prophet** *(Experimental — optional)* — [Meta's Prophet](https://facebook.github.io/prophet/) with configurable changepoint scale and seasonality mode (additive / multiplicative). Full TOML config support. Disabled gracefully if package is absent.

---

## Configuration

### GUI Settings (all versions)

All settings are accessible from the **Setup** tab in the GUI:

| Setting | Default | Range / Options | Description |
|---|---|---|---|
| Company Name | `My Company` | ≤ 120 chars | Printed in all report headers |
| Analyst Name | `Automated Analytics Engine` | ≤ 120 chars | Printed in report footers |
| Currency Symbol | `$` | Any symbol | Used in all monetary displays |
| Forecast Horizon | `12` months | 1 – 60 | How many months ahead to forecast |
| Confidence Level | `95%` | 0 – 100% | CI width for all models |
| Seasonality Period | `12` (auto if `0`) | 0 or 2 – 60 | Override periodogram-detected period |
| Bootstrap Samples | `300` | 50 – 5,000 | Bootstrap iterations for ML CIs |
| CV Folds | `3` | ≥ 2 | Walk-forward cross-validation splits |
| Models | All enabled | Toggle per model | Enable / disable individual models |
| Output Formats | PDF + Excel | PDF / Excel / Both | Report types to generate |

### TOML Persistent Configuration (Experimental only)

The experimental build saves all settings to a TOML file so they survive between runs.

**Config file locations** (checked in this order):

1. `./sf_config.toml` — portable, next to the script *(takes precedence)*
2. `~/.config/sales_forecast/config.toml` — user-level config

**Example `sf_config.toml`:**

```toml
[forecast]
horizon            = 12
confidence_level   = 0.95
seasonality_period = 12   # set to 0 for auto-detect
n_bootstrap        = 300
cv_folds           = 3

[models]
arima             = true
exp_smoothing     = true
random_forest     = true
gradient_boosting = true
ridge             = true
xgboost           = true   # requires: pip install xgboost
lightgbm          = true   # requires: pip install lightgbm
prophet           = true   # requires: pip install prophet

[prophet]
changepoint_prior_scale = 0.05
seasonality_mode        = "additive"   # or "multiplicative"

[report]
company_name   = "My Company"
analyst_name   = "Automated Analytics Engine"
currency       = "$"
formats        = ["pdf", "excel"]
```

> Settings edited in the GUI are automatically written back to the TOML file on each run.

---

## Report Output

Three files are produced on every run:

```
sales_forecast_20260115_143022.pdf
sales_forecast_20260115_143022.xlsx
sales_forecast_20260115_143022.json
```

Application logs: `~/SalesForecastReports/sales_forecast.log` — 5 MB rotating, 3 backups.

### PDF Report Sections

1. Executive KPI Summary
2. Historical Sales Performance — trend, anomalies, YoY growth, heatmap
3. Time Series Decomposition — trend · seasonal · residual
4. Product-Level Analysis
5. Demand Forecasting Results — all models vs ensemble + residual diagnostics
6. Monthly Forecast Detail with Confidence Intervals
7. Strategic Insights & Recommendations

### Excel Workbook Sheets

| Sheet | Contents |
|---|---|
| 📊 Summary | 23-row KPI table |
| 🔮 Forecast | Monthly point forecast + CI + MoM change + cumulative total |
| 🤖 Models | All model metrics, AIC, diagnostics, CV sMAPE |
| 📈 History | Monthly sales with rolling averages and growth rates |
| 🏷️ Products | Per-product revenue, units, returns, market share |
| 🔬 Decomposition | Trend / seasonal / residual values per month |
| ⚠️ Anomalies | Detected anomalies with Z-score and severity label |
| 📁 Raw Data | Up to 5,000 raw input records |
| 📷 Charts | Embedded PNG chart images |

### JSON Summary

Machine-readable output containing all model metrics, ensemble weights, CI method descriptions, detected seasonality, diagnostics, and full platform info — suitable for automated pipelines and audit trails.

---

## Building a Standalone Executable

A standalone build runs on any machine **without Python installed**.

**Pre-requisite — all platforms:**

```bash
pip install pyinstaller
```

Refer to the [PyInstaller documentation](https://pyinstaller.org/en/stable/) for advanced options.

> **Expected size:** 150–250 MB. NumPy, SciPy, scikit-learn, and Matplotlib are bundled — this is normal for PyInstaller + scientific Python.

---

### Windows (.exe)

```bash
cd your-repo

# ── v2 build (recommended) ────────────────────────────────────────────────────
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "SalesForecastingSystem" ^
  --icon "assets/icon.ico" ^
  src/sales_forecast_gui_v2.py

# ── v1 build ──────────────────────────────────────────────────────────────────
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "SalesForecastingSystem_v1" ^
  src/sales_forecast_gui_v1.py

# ── Experimental build ────────────────────────────────────────────────────────
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "SalesForecastingSystem_Experimental" ^
  --icon "assets/icon.ico" ^
  src/forecasting_experimental.py
```

> Remove `--icon` if you don't have an `.ico` file.

**Output:** `dist\SalesForecastingSystem.exe`

**PyInstaller flag reference:**

| Flag | Effect |
|---|---|
| `--onefile` | Bundles everything into a single `.exe` |
| `--windowed` | Suppresses console window (GUI-only mode) |
| `--name` | Sets the output filename |
| `--icon` | Sets the taskbar / desktop icon (`.ico` format) |

**If hidden sklearn modules cause runtime errors:**

```bash
pyinstaller --onefile --windowed ^
  --hidden-import="sklearn.utils._cython_blas" ^
  --hidden-import="sklearn.neighbors.typedefs" ^
  --hidden-import="sklearn.neighbors._partition_nodes" ^
  src/sales_forecast_gui_v2.py
```

**Experimental build — include optional model packages:**

```bash
pyinstaller --onefile --windowed ^
  --hidden-import="xgboost" ^
  --hidden-import="lightgbm" ^
  --hidden-import="prophet" ^
  --hidden-import="sklearn.utils._cython_blas" ^
  src/forecasting_experimental.py
```

---

### macOS (.app)

```bash
# ── v2 build (recommended) ────────────────────────────────────────────────────
pyinstaller \
  --onefile \
  --windowed \
  --name "SalesForecastingSystem" \
  src/sales_forecast_gui_v2.py

# ── v1 build ──────────────────────────────────────────────────────────────────
pyinstaller \
  --onefile \
  --windowed \
  --name "SalesForecastingSystem_v1" \
  src/sales_forecast_gui_v1.py

# ── Experimental build ────────────────────────────────────────────────────────
pyinstaller \
  --onefile \
  --windowed \
  --name "SalesForecastingSystem_Experimental" \
  src/forecasting_experimental.py
```

**Output:** `dist/SalesForecastingSystem.app`

> **Apple Silicon (M1 / M2 / M3 / M4):** Use a native ARM Python install for best performance. Rosetta builds work but are slower. Verify:
> ```bash
> python3 -c "import platform; print(platform.machine())"
> # Expected output: arm64
> ```

---

### Linux (binary)

```bash
# ── v2 build (recommended) ────────────────────────────────────────────────────
pyinstaller \
  --onefile \
  --name "sales_forecast_v2" \
  src/sales_forecast_gui_v2.py

# ── v1 build ──────────────────────────────────────────────────────────────────
pyinstaller \
  --onefile \
  --name "sales_forecast_v1" \
  src/sales_forecast_gui_v1.py

# ── Experimental build ────────────────────────────────────────────────────────
pyinstaller \
  --onefile \
  --name "sales_forecast_experimental" \
  src/forecasting_experimental.py
```

**Output:** `dist/sales_forecast_v2`

Make executable and launch:

```bash
chmod +x dist/sales_forecast_v2
./dist/sales_forecast_v2
```

---

## Version Comparison

| Capability | v1 | v2 | Experimental |
|---|:---:|:---:|:---:|
| 5 core forecasting models | ✅ | ✅ | ✅ |
| XGBoost / LightGBM / Prophet | ✗ | ✗ | ✅ optional |
| Stacked Ensemble | ✅ | ✅ | ✅ |
| PDF + Excel + JSON output | ✅ | ✅ | ✅ |
| Walk-forward CV | ✅ | ✅ | ✅ |
| TOML config persistence | ✗ | ✗ | ✅ |
| Named logger `SalesForecast` | ✅ | ✅ | ✅ |
| Global crash handler | ✅ | ✅ | ✅ |
| Debug mode via env var | ✗ | ✅ | ✅ |
| NaN mask in metrics | ✗ | ✅ | ✅ |
| Finite fallback on empty metrics | Zeros → breaks ensemble | ✅ Safe tuple | ✅ |
| Decompose length guard | `assert` — crashes | ✅ `np.resize` + log | ✅ |
| ML feature shape validation | Silent wrong results | ✅ + `nan_to_num` | ✅ |
| Button re-enable on cancel | After thread only | ✅ Immediate | ✅ |
| Safe `mainloop` exit | Hangs on Ctrl-C | ✅ `try/except` | ✅ |
| Prophet TOML config | ✗ | ✗ | ✅ |
| Full class docstrings | Partial | ✅ | ✅ |

> **Recommendation:** **v2** for stable / shared deployments · **Experimental** for extended model suite, research, or benchmarking.

---

## Troubleshooting

**App opens then immediately closes (standalone `.exe`)**
Remove `--windowed` from your PyInstaller command, rebuild, and re-run — the console will reveal the error. Fix it, then add `--windowed` back.

**`.exe` is 150–250 MB**
Expected. NumPy, SciPy, scikit-learn, and Matplotlib are all bundled. Standard behaviour for [PyInstaller](https://pyinstaller.org/en/stable/) + scientific Python.

**Antivirus flags the `.exe`**
Common false positive with PyInstaller. Sign with a code-signing certificate or submit to your AV vendor for whitelisting.

**`ModuleNotFoundError: No module named 'tkinter'`**
See [tkinter Platform Notes](#tkinter-platform-notes).

**`prophet` install fails on Python 3.12+**
```bash
pip install prophet --no-build-isolation
```
If that also fails, see the [Prophet installation guide](https://facebook.github.io/prophet/docs/installation.html).

**`lightgbm` not available on Apple Silicon**
```bash
pip install lightgbm --prefer-binary
```
See the [LightGBM installation guide](https://lightgbm.readthedocs.io/en/latest/Installation-Guide.html) for platform-specific wheels.

**TOML config not loading (experimental, Python < 3.11)**
Python 3.11+ ships `tomllib` built in. Older versions need:
```bash
pip install tomli tomli-w
```

**Charts missing from PDF / Excel reports**
```bash
pip install pillow
```
See [Pillow docs](https://pillow.readthedocs.io/en/stable/installation.html).

**XGBoost / LightGBM / Prophet not detected after installing**
Confirm you installed into the active virtual environment:
```bash
pip show xgboost lightgbm prophet
```

---

## Author

<div align="center">

| | |
|---|---|
| **Name** | [Abhishek Srivastava](https://www.linkedin.com/in/abhishek-srivastava-1538461b1/) |
| **Role** | Business Analyst, EY |
| **Connect** | [![LinkedIn](https://img.shields.io/badge/LinkedIn-Abhishek_Srivastava-0A66C2?style=flat-square&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/abhishek-srivastava-1538461b1/) |

</div>

---

<div align="center">

*Forecasts carry inherent uncertainty.*  
*Always combine model output with domain expertise and business context before making decisions.*

</div>
