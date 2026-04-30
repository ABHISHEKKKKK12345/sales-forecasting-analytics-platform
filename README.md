# 📈 Sales Forecasting & Automated Reporting System

> Production-grade demand forecasting engine with a dark-themed GUI, statistical rigour, ML ensemble modelling, and one-click PDF/Excel report generation.

---

## Table of Contents

- [Overview](#overview)
- [Versions](#versions)
- [Features](#features)
- [Screenshots & Output](#screenshots--output)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Installation](#installation)
- [Running the Application](#running-the-application)
- [Building a Standalone .exe (Windows)](#building-a-standalone-exe-windows)
- [Building a Standalone App (macOS & Linux)](#building-a-standalone-app-macos--linux)
- [Data Format](#data-format)
- [Forecasting Models](#forecasting-models)
- [Report Output](#report-output)
- [Configuration](#configuration)
- [Version Comparison](#version-comparison)
- [License](#license)

---

## Overview

This system provides a complete sales forecasting and business analytics pipeline in a single Python file. Load your own sales data (or use the built-in demo dataset), run five forecasting models in parallel, and generate publication-quality PDF reports and multi-sheet Excel workbooks — all from a polished desktop GUI with no internet connection required.

---

## Versions

Two versions of the application are included in `src/`:

| File | Description |
|---|---|
| `sales_forecast_gui.py` | **v1 — Original release.** Full-featured forecasting engine. |
| `sales_forecast_gui_v2.py` | **v2 — Hardened & refined.** All bug fixes, global error handling, improved logging, NaN guards, and edge-case protection applied. **Recommended for production use.** |

See [Version Comparison](#version-comparison) for a full diff of what changed.

---

## Features

### Forecasting Engine
- **ARIMA** — full MLE via innovations form, AIC/BIC grid search (auto order selection), ψ-weight confidence interval propagation
- **Holt-Winters** — triple exponential smoothing with automatic additive vs multiplicative selection by AIC, analytical and bootstrap CIs
- **Random Forest** — 200 trees, PACF-guided lag selection, bootstrap CI
- **Gradient Boosting** — 200 rounds, learning rate 0.05, bootstrap CI
- **Ridge Regression** — L2-regularised linear baseline, bootstrap CI
- **Stacked Ensemble** — inverse CV-sMAPE weighted combination of all models (non-leaky, walk-forward)

### Statistical Rigour
- **Metrics**: sMAPE · MASE · RMSE · MAE · R²
- **Confidence Intervals**: ARIMA ψ-weight propagation · HW analytical/bootstrap · ML non-parametric bootstrap
- **Diagnostics**: Ljung-Box autocorrelation test · Shapiro-Wilk normality · Durbin-Watson serial correlation · heteroscedasticity check
- **Cross-validation**: Expanding-window walk-forward CV (no look-ahead bias)
- **Anomaly detection**: 6-month rolling Z-score (threshold 2.5σ)
- **Marketing ROI**: bootstrap confidence interval (95%)

### Business Analytics
- CAGR, YoY growth, MoM growth, revenue volatility (CV)
- Peak/trough detection, return rate analysis
- Product-level and region-level breakdowns
- Seasonal period auto-detection via periodogram

### GUI
- Dark GitHub-inspired theme (fully themed with ttk)
- Three-tab layout: Setup → Log → Results
- Real-time progress bar and colour-coded activity log
- KPI dashboard on results tab
- Sortable model comparison table
- One-click open for generated files
- Cross-platform mousewheel scrolling

### Output
- **PDF report** — multi-page A4 with branded header/footer, 8 embedded charts, 6 sections
- **Excel workbook** — 9 sheets with frozen panes, Excel Tables, conditional formatting, embedded chart images
- **JSON summary** — machine-readable metrics, model scores, platform info

---

## Screenshots & Output

### PDF Report Sections
1. Executive KPI Summary
2. Historical Sales Performance (trend, anomalies, YoY growth, heatmap)
3. Time Series Decomposition (trend · seasonal · residual)
4. Product-Level Analysis
5. Demand Forecasting Results (all models vs ensemble, residual diagnostics)
6. Monthly Forecast Detail with CI
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
| 📷 Charts | Embedded PNG charts (one sheet per chart) |

---

## Project Structure

```
your-repo/
├── README.md                  ← you are here
├── requirements.txt
├── .gitignore
├── LICENSE
└── src/
    ├── sales_forecast_gui.py      ← v1 original
    └── sales_forecast_gui_v2.py   ← v2 hardened (recommended)
```

---

## Requirements

- **Python** 3.9 or higher
- **tkinter** (usually bundled with Python; see platform notes below)

### Python Packages

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

All packages are listed in `requirements.txt` at the repo root.

### Platform Notes for tkinter

| Platform | Action if tkinter is missing |
|---|---|
| **Windows** | Reinstall Python and check "tcl/tk and IDLE" during setup |
| **macOS** | `brew install python-tk` |
| **Ubuntu / Debian** | `sudo apt install python3-tk` |
| **Fedora / RHEL** | `sudo dnf install python3-tkinter` |

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/your-username/your-repo.git
cd your-repo
```

### 2. Create a virtual environment (recommended)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS / Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

---

## Running the Application

```bash
# Run v2 (recommended)
python src/sales_forecast_gui_v2.py

# Run v1 (original)
python src/sales_forecast_gui.py
```

On first launch, if any dependency is missing, a dialog will appear listing the exact `pip install` command needed before the main window opens.

### Quick Start with Demo Data

1. Launch the app
2. Leave **"Use built-in demo data"** checked (5 products · 60 months pre-generated)
3. Set your company name and forecast horizon
4. Click **🚀 Run Analysis & Generate Reports**
5. Choose a save location when prompted
6. Reports open automatically when complete

---

## Building a Standalone .exe (Windows)

A standalone `.exe` requires no Python installation on the target machine.

### Step 1 — Install PyInstaller

```bash
pip install pyinstaller
```

### Step 2 — Build

```bash
cd your-repo

pyinstaller ^
  --onefile ^
  --windowed ^
  --name "SalesForecastingSystem" ^
  --icon "assets/icon.ico" ^
  src/sales_forecast_gui_v2.py
```

> Remove `--icon` if you don't have an `.ico` file.

| Flag | Effect |
|---|---|
| `--onefile` | Bundles everything into a single `.exe` |
| `--windowed` | Suppresses the console window (GUI only) |
| `--name` | Sets the output filename |
| `--icon` | Sets the taskbar/desktop icon |

### Step 3 — Find your executable

```
dist/
└── SalesForecastingSystem.exe   ← distribute this file
```

### Troubleshooting the .exe

**App opens then immediately closes:**
Remove `--windowed` temporarily and re-build to see console errors.

**Missing module errors at runtime:**
Add hidden imports:
```bash
pyinstaller --onefile --windowed \
  --hidden-import="sklearn.utils._cython_blas" \
  --hidden-import="sklearn.neighbors.typedefs" \
  --hidden-import="sklearn.neighbors._partition_nodes" \
  src/sales_forecast_gui_v2.py
```

**Antivirus flags the .exe:**
This is a false positive common with PyInstaller. Sign the executable with a code-signing certificate to resolve it, or submit it to your AV vendor for whitelisting.

**Large file size:**
The `.exe` will be approximately 150–250 MB due to bundled scientific libraries (numpy, scipy, sklearn, matplotlib). This is expected.

---

## Building a Standalone App (macOS & Linux)

### macOS — .app bundle

```bash
pip install pyinstaller

pyinstaller \
  --onefile \
  --windowed \
  --name "SalesForecastingSystem" \
  src/sales_forecast_gui_v2.py
```

Output: `dist/SalesForecastingSystem.app`

> On Apple Silicon (M1/M2/M3), ensure you are using a native ARM Python install for best performance. Rosetta builds work but are slower.

### Linux — standalone binary

```bash
pip install pyinstaller

pyinstaller \
  --onefile \
  --name "sales_forecast" \
  src/sales_forecast_gui_v2.py
```

Output: `dist/sales_forecast` (executable binary)

Make it executable and run:
```bash
chmod +x dist/sales_forecast
./dist/sales_forecast
```

---

## Data Format

### Required Columns

| Column | Type | Description |
|---|---|---|
| `date` | date string | Any parseable date format (YYYY-MM-DD, MM/DD/YYYY, etc.) |
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

The loader tries multiple encodings automatically (`utf-8`, `latin-1`, `cp1252`, `iso-8859-15`).

### Example CSV

```csv
date,product,region,sales,units,returns,marketing_spend
2022-01-01,Widget A,North,125000,2100,1800,14000
2022-02-01,Widget A,North,118000,1980,1650,13500
2022-03-01,Widget B,South,89000,3100,890,9800
```

---

## Forecasting Models

### ARIMA
- Full Gaussian MLE via innovations form
- Grid search over `p ∈ [0,3]`, `d ∈ [0,1]`, `q ∈ [0,2]` — best selected by AIC
- Confidence intervals via ψ-weight variance propagation (Box-Jenkins §5.4)
- Multi-start L-BFGS-B optimisation for reliable convergence

### Holt-Winters
- Triple exponential smoothing (level + trend + seasonal)
- Additive and multiplicative variants both fitted; AIC selects the winner
- Analytical CI for additive, bootstrap CI for multiplicative

### Random Forest
- 200 trees, max depth 8, min samples per leaf 2
- Features: PACF-guided lags, rolling mean/std/min/max (3/6/12 window), harmonic seasonality encoding, polynomial trend terms
- Recursive multi-step forecast strategy

### Gradient Boosting
- 200 rounds, learning rate 0.05, max depth 4, subsample 0.8
- Same feature engineering as Random Forest

### Ridge Regression
- L2 regularisation (α = 1.0)
- Fast linear baseline with same feature matrix

### Stacked Ensemble
- Weights = inverse of walk-forward CV sMAPE per model
- Falls back to uniform weights if CV scores are degenerate
- NaN-guarded — always produces a finite, valid forecast

---

## Report Output

Reports are saved to a location you choose at runtime. Three files are always produced:

```
sales_forecast_20240115_143022.pdf
sales_forecast_20240115_143022.xlsx
sales_forecast_20240115_143022.json
```

The JSON summary contains all model metrics, ensemble weights, statistics, CI method descriptions, and platform info — useful for automated pipelines or audit trails.

---

## Configuration

The following settings are available in the GUI Setup tab:

| Setting | Default | Description |
|---|---|---|
| Company Name | `My Company` | Appears in report headers |
| Currency Symbol | `$` | Used in all monetary displays |
| Forecast Horizon | `12` months | 1–36 months |
| Confidence Level | `95%` | CI width for all models |
| Seasonality Period | `12` (auto-detect if 0) | Override the detected seasonal period |
| Models | All enabled | Toggle individual models on/off |
| Output Formats | PDF + Excel | Choose PDF, Excel, or both |

---

## Version Comparison

| Capability | v1 `sales_forecast_gui.py` | v2 `sales_forecast_gui_v2.py` |
|---|---|---|
| All forecasting models | ✅ | ✅ |
| PDF + Excel + JSON output | ✅ | ✅ |
| Walk-forward CV | ✅ | ✅ |
| Logging level | WARNING | **INFO + timestamp format** |
| Named logger | ✗ | ✅ `SalesForecast` |
| Global crash handler | ✗ | ✅ `sys.excepthook` |
| NaN mask in metrics | ✗ | ✅ filters before computing |
| Finite fallback on empty metrics | Returns zeros (breaks ensemble) | ✅ Returns `(0,0,100,1,100,0)` |
| Decompose length guard | `assert` — crashes on edge cases | ✅ `np.resize` + warning log |
| ML feature shape mismatch | Silently wrong results possible | ✅ Checked + `nan_to_num` guard |
| Button re-enable on cancel | Only after thread completes | ✅ Immediately on validation fail |
| Safe `mainloop` exit | Bare call — hangs on Ctrl-C | ✅ `try/except KeyboardInterrupt` |
| Module docstring | 40-line ASCII box (with typo) | Clean 3-line docstring |
| Class docstrings | Missing on 3 engine classes | ✅ All classes documented |

**Recommendation: use v2 for any production or shared deployment.**

---

## License

See `LICENSE` in the repository root.

---

## 👨‍💼 Author

<div align="center">

| Field | Detail |
|---|---|
| **Name** | [Abhishek](https://www.linkedin.com/in/abhishek-srivastava-1538461b1/) |
| **Project** | Customer Churn Intelligence Platform |

</div>

---

*Forecasts carry inherent uncertainty. Always combine model output with domain expertise and business context before making decisions.*
