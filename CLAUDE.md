# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

The **P-Fin 8 Data Exploration Tool** is a single-file Streamlit web app for exploring TIAA Institute-GFLEC Personal Finance Index (P-Fin 8) survey data. It visualizes financial literacy across topics, demographics, and financial well-being dimensions for years 2017–2026.

## Running the App

```bash
pip install -r requirements.txt
streamlit run pfin8_data_tool.py
# Available at http://localhost:8501
```

A devcontainer config (`.devcontainer/devcontainer.json`) is available for VS Code — it auto-installs dependencies and launches on port 8501.

## Architecture

Everything lives in `pfin8_data_tool.py` (~1,700 lines). The flow is:

1. **Sidebar** (`render_sidebar`) — user selects exploration type, analysis type, view mode, topic, demographic filters, year/age ranges
2. **Data loading** — Excel files loaded and cached with `@st.cache_data`
   - `allYearsPFin8.xlsx` — historical data 2017–2026
   - `PFin2026_GenPop.xlsx` — 2026 general population data
3. **Analysis** (`run_analysis`) — applies filters, computes weighted statistics, prepares chart data
4. **Visualization** (`create_chart`) — Plotly charts (bar, grouped bar, stacked, pie, line, table) with export to PNG/HTML/CSV/Excel
5. **Validation** (`run_sanity_checks`) — automated data integrity checks; visible in the debug panel when `DEBUG_MODE = True`

### Key Functions

| Function | Purpose |
|---|---|
| `load_all_years()` | Load & cache multi-year Excel data |
| `load_genpop()` | Load & cache 2026 general population data |
| `render_sidebar()` | All sidebar UI and filter state |
| `run_analysis()` | Core analysis dispatch |
| `prepare_topic_binary_data()` | % Correct / Not Correct per topic |
| `prepare_topic_cat3_data()` | % Correct / Incorrect / Don't Know per topic |
| `prepare_total_correct_data()` | Score distribution (0–8) by group |
| `weighted_percentage_binary()` | Survey-weighted binary calc |
| `weighted_percentage_cat3()` | Survey-weighted 3-category calc |
| `weighted_total_correct_distribution()` | Survey-weighted score distribution |
| `create_chart()` | Plotly chart generation |
| `run_sanity_checks()` | Validation: weight sums, percentage bounds, sample size |

### Exploration Modes

- **Over the Years** — trend lines 2017–2026
- **Demographics** — breakdown by age, gender, education, income, generation, race/ethnicity, employment, marital status, etc.
- **Financial Well-Being** — breakdown by debt constraint, savings fragility, financial stress, etc.

### Analysis Types

- **Topic Bucket** — per-topic performance; binary (correct/not) or 3-category (correct/incorrect/don't know)
- **Total Correct** — distribution of scores 0–8

## Configuration

- `config.toml` and `.streamlit/config.toml` — Streamlit theme (primary color `#1f4e79`) and server settings
- `DEBUG_MODE = True` at the top of `pfin8_data_tool.py` — enables the debug validation panel
- `MIN_SAMPLE_SIZE = 30` — threshold for sample size warnings in sanity checks

## Data Notes

- All statistics are **survey-weighted** to represent the U.S. population
- Data corrections are applied on load: `"Corrcet"` → `"Correct"`, `"None"` → `"No hours"`
- Custom age grouping supports up to 8 user-defined ranges with overlap validation
- Export uses base64 encoding for client-side file downloads (no server-side file writes)
