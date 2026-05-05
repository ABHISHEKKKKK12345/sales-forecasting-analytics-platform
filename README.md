<div align="center">

# 📈 Sales Forecasting & Automated Reporting System

**Production-grade demand forecasting engine with a dark-themed desktop GUI,  
ML ensemble modelling, statistical rigour, and one-click PDF / Excel report generation.**

[![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-See_LICENSE-green?style=for-the-badge)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=for-the-badge)]()
[![GUI](https://img.shields.io/badge/GUI-tkinter%20Dark%20Theme-161B22?style=for-the-badge)]()

</div>

---

## 📑 Table of Contents

- [Overview](#-overview)
- [Repository Structure](#-repository-structure)
- [Versions at a Glance](#-versions-at-a-glance)
- [Features](#-features)
- [Requirements](#-requirements)
- [Installation](#-installation)
- [Running the Application](#-running-the-application)
- [Data Format](#-data-format)
- [Forecasting Models](#-forecasting-models)
- [Configuration](#-configuration)
- [Report Output](#-report-output)
- [Building a Standalone Executable](#-building-a-standalone-executable)
  - [Windows (.exe)](#windows-exe)
  - [macOS (.app)](#macos-app)
  - [Linux (binary)](#linux-binary)
- [Version Comparison](#-version-comparison)
- [Troubleshooting](#-troubleshooting)
- [Author](#-author)

---

## 🔭 Overview

This system delivers a complete **sales forecasting and business analytics pipeline** in self-contained Python files. Load your own data (CSV, Excel, JSON, Parquet) or use the built-in 60-month demo dataset, run up to **nine forecasting models in parallel**, and generate publication-quality **PDF reports** and **multi-sheet Excel workbooks** — all from a polished desktop GUI with zero internet connection required.

Three files are included, each representing a stage of development:

| File | Version | Purpose |
|---|---|---|
| `sales_forecast_gui_v1.py` | v1.0 | Original full-featured release |
| `sales_forecast_gui_v2.py` | v2.0 | Hardened & production-ready |
| `forecasting_experimental.py` | Experimental | Extended model suite + TOML config |

---

## 🗂 Repository Structure

```
your-repo/
│
├── README.md                        ← you are here
├── requirements.txt                 ← pip dependencies
├── requirements_experimental.txt   ← extended deps for experimental build
├── .gitignore
├── LICENSE
│
└── src/
    ├── sales_forecast_gui_v1.py         ← v1 — original release
    ├── sales_forecast_gui_v2.py         ← v2 — hardened (recommended)
    └── forecasting_experimental.py      ← experimental — extended model suite
```

---

## 🔖 Versions at a Glance

### `sales_forecast_gui_v1.py` — Version 1.0
> The original production release. Fully functional with 5 forecasting models, PDF + Excel + JSON output, walk-forward cross-validation, and a complete dark-themed GUI. Suitable for personal use and testing.

### `sales_forecast_gui_v2.py` — Version 2.0 ✅ Recommended
> All v1 functionality plus hardened error handling, NaN guards, debug mode via environment variable (`SALES_FORECAST_DEBUG=1`), improved logging, decomposition length guard, ML feature shape validation, and safer `mainloop` exit. **Use this for any shared or production deployment.**

### `forecasting_experimental.py` — Experimental 🧪
> Extends v2 with three additional optional models — **XGBoost**, **LightGBM**, and **Prophet** — along with a full **TOML-based configuration system** that persists settings between runs. Optional models degrade gracefully if their packages are not installed. Intended for research, benchmarking, and advanced deployments.

---

## ✨ Features

### Forecasting Engine

| Model | v1 | v2 | Experimental |
|---|:---:|:---:|:---:|
| ARIMA (MLE grid, AIC/BIC) | ✅ | ✅ | ✅ |
| Holt-Winters (add/mul, AIC) | ✅ | ✅ | ✅ |
| Random Forest (200 trees) | ✅ | ✅ | ✅ |
| Gradient Boosting (200 rounds) | ✅ | ✅ | ✅ |
| Ridge Regression (L2 baseline) | ✅ | ✅ | ✅ |
| Stacked Ensemble | ✅ | ✅ | ✅ |
| XGBoost | ✗ | ✗ | ✅ (optional) |
| LightGBM | ✗ | ✗ | ✅ (optional) |
| Prophet | ✗ | ✗ | ✅ (optional) |

### Statistical Rigour
- **Metrics** — sMAPE · MASE · RMSE · MAE · R²
- **Confidence Intervals** — ARIMA ψ-weight propagation · Holt-Winters analytical/bootstrap · ML non-parametric bootstrap
- **Diagnostics** — Ljung-Box autocorrelation · Shapiro-Wilk normality · Durbin-Watson serial correlation · heteroscedasticity check
- **Cross-validation** — expanding-window walk-forward CV (no look-ahead bias)
- **Anomaly detection** — 6-month rolling Z-score (threshold 2.5σ)
- **Marketing ROI** — bootstrap confidence interval (95%)
- **Seasonal period auto-detection** via periodogram

### Business Analytics
- CAGR, YoY growth, MoM growth, revenue volatility (CV)
- Peak/trough detection, return rate analysis
- Product-level and region-level breakdowns

### GUI
- Dark GitHub-inspired theme (full `ttk` styling, hex `#0D1117` background)
- Three-tab layout: **Setup → Log → Results**
- Real-time progress bar and colour-coded activity log
- KPI dashboard on results tab
- Sortable model comparison table
- One-click open for generated files
- Cross-platform mousewheel scrolling

### Output
- **PDF report** — multi-page A4, branded header/footer, 8 embedded charts, 7 sections
- **Excel workbook** — 9 sheets with frozen panes, Excel Tables, conditional formatting, embedded chart images
- **JSON summary** — machine-readable metrics, model scores, CI method descriptions, platform info

---

## 📦 Requirements

- **Python** 3.9 or higher
- **tkinter** — usually bundled with Python (see platform notes below)

### Core Python Packages

```
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

### Additional Packages — Experimental Only

```
xgboost>=1.7          # optional — XGBoost model
lightgbm>=3.3         # optional — LightGBM model
prophet>=1.1          # optional — Prophet model
tomli>=2.0            # required on Python < 3.11 (TOML config read)
tomli-w>=1.0          # required for TOML config write
```

> **Note:** The experimental file checks for XGBoost, LightGBM, and Prophet at runtime and simply disables the corresponding models if the packages are not present. You do not need to install all three.

### tkinter Platform Notes

| Platform | Action if tkinter is missing |
|---|---|
| **Windows** | Reinstall Python and check **"tcl/tk and IDLE"** during setup wizard |
| **macOS** | `brew install python-tk` |
| **Ubuntu / Debian** | `sudo apt install python3-tk` |
| **Fedora / RHEL** | `sudo dnf install python3-tkinter` |

---

## 🛠 Installation

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

**For the experimental build (adds XGBoost, LightGBM, Prophet, tomli):**
```bash
pip install -r requirements.txt
pip install -r requirements_experimental.txt
```

Or install optional models individually:
```bash
pip install xgboost lightgbm prophet tomli tomli-w
```

> On first launch, if any **core** dependency is missing, the app will display a dialog listing the exact `pip install` command needed before the main window opens.

---

## ▶️ Running the Application

```bash
# v2 — recommended
python src/sales_forecast_gui_v2.py

# v1 — original
python src/sales_forecast_gui_v1.py

# Experimental — extended model suite
python src/forecasting_experimental.py

# Experimental — with verbose debug logging
SALES_FORECAST_DEBUG=1 python src/forecasting_experimental.py   # macOS / Linux
set SALES_FORECAST_DEBUG=1 && python src/forecasting_experimental.py  # Windows CMD
```

### Quick Start with Demo Data

1. Launch any version of the app
2. Leave **"Use built-in demo data"** checked (5 products · 60 months pre-generated)
3. Set your company name and desired forecast horizon
4. Click **🚀 Run Analysis & Generate Reports**
5. Choose a save directory when prompted
6. PDF and Excel reports open automatically when complete

---

## 📂 Data Format

### Required Columns

| Column | Type | Description |
|---|---|---|
| `date` | date string | Any parseable format — `YYYY-MM-DD`, `MM/DD/YYYY`, etc. |
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
| CSV / TSV | `.csv` `.tsv` |
| Excel | `.xlsx` `.xlsm` `.xls` |
| JSON | `.json` (records, split, index, values, table orientations) |
| Parquet | `.parquet` |

The loader automatically tries multiple encodings: `utf-8` → `latin-1` → `cp1252` → `iso-8859-15`.

### Example CSV

```csv
date,product,region,sales,units,returns,marketing_spend
2022-01-01,Widget A,North,125000,2100,1800,14000
2022-02-01,Widget A,North,118000,1980,1650,13500
2022-03-01,Widget B,South,89000,3100,890,9800
```

---

## 🤖 Forecasting Models

### ARIMA
- Full Gaussian MLE via innovations form
- Grid search over `p ∈ [0,3]`, `d ∈ [0,1]`, `q ∈ [0,2]` — winner selected by AIC
- Confidence intervals via ψ-weight variance propagation (Box-Jenkins §5.4)
- Multi-start L-BFGS-B optimisation for reliable convergence

### Holt-Winters
- Triple exponential smoothing (level + trend + seasonal component)
- Both additive and multiplicative variants fitted; AIC selects the winner
- Analytical CI for additive; bootstrap CI for multiplicative

### Random Forest
- 200 trees, max depth 8, min samples per leaf 2
- Features: PACF-guided lags, rolling mean/std/min/max (3/6/12 window), harmonic seasonality encoding, polynomial trend terms
- Recursive multi-step forecast strategy

### Gradient Boosting
- 200 rounds, learning rate 0.05, max depth 4, subsample 0.8
- Same feature engineering pipeline as Random Forest

### Ridge Regression
- L2 regularisation (α = 1.0)
- Fast linear baseline using the same feature matrix

### Stacked Ensemble
- Weights = inverse of walk-forward CV sMAPE per model
- Falls back to uniform weights if CV scores are degenerate
- NaN-guarded — always produces a finite, valid forecast

### XGBoost *(Experimental only — optional)*
- Gradient boosting via `xgboost` library; same feature matrix as RF/GB
- Gracefully disabled if `xgboost` is not installed

### LightGBM *(Experimental only — optional)*
- Gradient boosting via `lightgbm` library; same feature matrix
- Gracefully disabled if `lightgbm` is not installed

### Prophet *(Experimental only — optional)*
- Facebook/Meta Prophet with configurable changepoint scale and seasonality mode
- Configurable via TOML: `prophet.changepoint_prior_scale` and `prophet.seasonality_mode`
- Gracefully disabled if `prophet` is not installed

---

## ⚙️ Configuration

### GUI Settings (all versions)

| Setting | Default | Description |
|---|---|---|
| Company Name | `My Company` | Appears in report headers |
| Analyst Name | `Automated Analytics Engine` | Appears in report footers |
| Currency Symbol | `$` | Used in all monetary displays |
| Forecast Horizon | `12` months | 1–60 months |
| Confidence Level | `95%` | CI width for all models |
| Seasonality Period | `12` (auto if `0`) | Override the periodogram-detected period |
| Models | All enabled | Toggle individual models on/off |
| Output Formats | PDF + Excel | Choose PDF, Excel, or both |
| Bootstrap samples | `300` | Number of bootstrap iterations for ML CIs |
| CV folds | `3` | Walk-forward cross-validation splits |

### TOML Persistent Configuration *(Experimental only)*

The experimental build persists all settings to a TOML file so they survive between runs.

**Config file locations (in priority order):**

1. `./sf_config.toml` — portable, next to the script (takes precedence)
2. `~/.config/sales_forecast/config.toml` — user-level config

**Example `sf_config.toml`:**

```toml
[forecast]
horizon            = 12
confidence_level   = 0.95
seasonality_period = 12
n_bootstrap        = 300
cv_folds           = 3

[models]
arima             = true
exp_smoothing     = true
random_forest     = true
gradient_boosting = true
ridge             = true
xgboost           = true
lightgbm          = true
prophet           = true

[prophet]
changepoint_prior_scale = 0.05
seasonality_mode        = "additive"   # or "multiplicative"

[report]
company_name   = "My Company"
analyst_name   = "Automated Analytics Engine"
currency       = "$"
formats        = ["pdf", "excel"]
```

Settings edited in the GUI are automatically saved back to the TOML file on each run.

---

## 📄 Report Output

Three files are always produced per run:

```
sales_forecast_20240115_143022.pdf
sales_forecast_20240115_143022.xlsx
sales_forecast_20240115_143022.json
```

Logs are written to `~/SalesForecastReports/sales_forecast.log` (5 MB rotating, 3 backups).

### PDF Report Sections

1. Executive KPI Summary
2. Historical Sales Performance (trend, anomalies, YoY growth, heatmap)
3. Time Series Decomposition (trend · seasonal · residual)
4. Product-Level Analysis
5. Demand Forecasting Results (all models vs ensemble + residual diagnostics)
6. Monthly Forecast Detail with Confidence Intervals
7. Strategic Insights & Recommendations

### Excel Workbook Sheets

| Sheet | Contents |
|---|---|
| 📊 Summary | 23-row KPI table |
| 🔮 Forecast | Monthly point forecast + CI + MoM change + cumulative |
| 🤖 Models | All model metrics, AIC, diagnostics, CV sMAPE |
| 📈 History | Monthly sales with rolling averages and growth rates |
| 🏷️ Products | Per-product revenue, units, returns, market share |
| 🔬 Decomposition | Trend / seasonal / residual values per month |
| ⚠️ Anomalies | Detected anomalies with Z-score and severity |
| 📁 Raw Data | Up to 5,000 raw records |
| 📷 Charts | Embedded PNG charts |

### JSON Summary

Machine-readable output containing all model metrics, ensemble weights, confidence interval method descriptions, detected statistics, and full platform info — useful for automated pipelines and audit trails.

---

## 📦 Building a Standalone Executable

A standalone build runs on any machine **without a Python installation**.

### Pre-requisite (all platforms)

```bash
pip install pyinstaller
```

---

### Windows (.exe)

```bash
cd your-repo

# Build v2 (recommended)
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "SalesForecastingSystem" ^
  --icon "assets/icon.ico" ^
  src/sales_forecast_gui_v2.py

# Build experimental
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "SalesForecastingSystem_Experimental" ^
  --icon "assets/icon.ico" ^
  src/forecasting_experimental.py
```

> Remove `--icon` if you don't have an `.ico` file.

**Output:** `dist\SalesForecastingSystem.exe`

**Flag reference:**

| Flag | Effect |
|---|---|
| `--onefile` | Bundles everything into a single `.exe` |
| `--windowed` | Suppresses the console window (GUI-only mode) |
| `--name` | Sets the output file name |
| `--icon` | Sets the taskbar / desktop icon |

**If missing modules at runtime (common with sklearn):**

```bash
pyinstaller --onefile --windowed ^
  --hidden-import="sklearn.utils._cython_blas" ^
  --hidden-import="sklearn.neighbors.typedefs" ^
  --hidden-import="sklearn.neighbors._partition_nodes" ^
  src/sales_forecast_gui_v2.py
```

**For the experimental build with optional models:**

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
# v2
pyinstaller \
  --onefile \
  --windowed \
  --name "SalesForecastingSystem" \
  src/sales_forecast_gui_v2.py

# Experimental
pyinstaller \
  --onefile \
  --windowed \
  --name "SalesForecastingSystem_Experimental" \
  src/forecasting_experimental.py
```

**Output:** `dist/SalesForecastingSystem.app`

> On **Apple Silicon (M1/M2/M3/M4)**, use a native ARM Python install for best performance. Rosetta builds work but are slower. Verify with:
> ```bash
> python3 -c "import platform; print(platform.machine())"
> # Should print: arm64
> ```

---

### Linux (binary)

```bash
# v2
pyinstaller \
  --onefile \
  --name "sales_forecast" \
  src/sales_forecast_gui_v2.py

# Experimental
pyinstaller \
  --onefile \
  --name "sales_forecast_experimental" \
  src/forecasting_experimental.py
```

**Output:** `dist/sales_forecast`

Make it executable and run:
```bash
chmod +x dist/sales_forecast
./dist/sales_forecast
```

---

## 🔄 Version Comparison

| Capability | v1 `sales_forecast_gui_v1.py` | v2 `sales_forecast_gui_v2.py` | Experimental `forecasting_experimental.py` |
|---|:---:|:---:|:---:|
| All 5 core forecasting models | ✅ | ✅ | ✅ |
| XGBoost / LightGBM / Prophet | ✗ | ✗ | ✅ (optional) |
| PDF + Excel + JSON output | ✅ | ✅ | ✅ |
| Walk-forward CV | ✅ | ✅ | ✅ |
| TOML config persistence | ✗ | ✗ | ✅ |
| Named logger `SalesForecast` | ✅ | ✅ | ✅ |
| Global crash handler | ✅ | ✅ | ✅ |
| Debug mode via env var | ✗ | ✅ `SALES_FORECAST_DEBUG=1` | ✅ `SALES_FORECAST_DEBUG=1` |
| NaN mask in metrics | ✗ | ✅ | ✅ |
| Finite fallback on empty metrics | Returns zeros → breaks ensemble | ✅ Returns `(0,0,100,1,100,0)` | ✅ |
| Decompose length guard | `assert` — crashes on edge cases | ✅ `np.resize` + warning log | ✅ |
| ML feature shape validation | Silently wrong possible | ✅ Checked + `nan_to_num` guard | ✅ |
| Button re-enable on cancel | Only after thread completes | ✅ Immediately on validation fail | ✅ |
| Safe `mainloop` exit | Bare call — hangs on Ctrl-C | ✅ `try/except KeyboardInterrupt` | ✅ |
| Prophet changepoint config | ✗ | ✗ | ✅ `prophet.*` in TOML |
| Class docstrings | Partial | ✅ All classes documented | ✅ |

**Recommendation:**
- **v2** for any stable, shared, or production deployment
- **Experimental** for extended model suite, research, or benchmarking

---

## 🔧 Troubleshooting

**App opens then immediately closes (built .exe)**
Remove `--windowed` to see the console output, debug the error, then re-add it.

**Large .exe file size (150–250 MB)**
Expected — numpy, scipy, sklearn, matplotlib are bundled. This is standard for PyInstaller + scientific Python.

**Antivirus flags the .exe**
Common false positive with PyInstaller. Sign the executable with a code-signing certificate or submit to your AV vendor for whitelisting.

**`ModuleNotFoundError: No module named 'tkinter'`**
See the [tkinter platform notes](#requirements) above for your OS.

**`prophet` install fails on Python 3.12+**
Prophet currently requires `pystan<3` on some platforms. Try:
```bash
pip install prophet --no-build-isolation
```

**`lightgbm` not available on Apple Silicon**
Install the pre-built wheel:
```bash
pip install lightgbm --prefer-binary
```

**TOML config not loading (experimental)**
Ensure `tomli` is installed on Python < 3.11. Python 3.11+ has `tomllib` built in:
```bash
pip install tomli tomli-w
```

**Charts not rendering in reports**
Ensure `pillow` is installed: `pip install pillow`

---

## 👨‍💼 Author

<div align="center">

| Field | Detail |
|---|---|
| **Name** | [Abhishek](https://www.linkedin.com/in/abhishek-srivastava-1538461b1/) |
| **Role** | Senior Consultant, EY |

</div>

---

<div align="center">

*Forecasts carry inherent uncertainty. Always combine model output with domain expertise and business context before making decisions.*

</div>
