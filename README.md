# 🏛️ Autonomous Institutional Equity Research & Valuation Agent

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Code Style: Institutional](https://img.shields.io/badge/Standard-Tier--1%20Institutional-1A365D.svg)]()

An autonomous, production-grade financial valuation and equity research agent. Automatically fetches exchange market data, computes comprehensive financial statement models, performs multi-method valuation, generates high-resolution Matplotlib charts, and compiles:

1. **A 16-Page Publication-Grade PDF Initiation Report** (Motilal Oswal / Goldman Sachs / Morgan Stanley institutional standard).
2. **A 10-Tab Executive Interactive Financial Model (`.xlsx`)** with pre-configured dynamic Excel Data Validation dropdowns, 7 KPI stat cards, and 6 dynamic charts.

---

## 🤖 4-Agent Parallel Orchestration Pipeline

Whenever a research job is triggered, the system coordinates four specialized sub-agent roles:

```
                      ┌────────────────────────────────────────┐
                      │            ORCHESTRATOR                │
                      │      (Master Task Coordinator)         │
                      └──────────────────┬─────────────────────┘
                                         │
       ┌──────────────────┬──────────────┴─────┬──────────────────┐
       ▼                  ▼                    ▼                  ▼
┌──────────────┐   ┌──────────────┐    ┌──────────────┐    ┌──────────────┐
│   AGENT 1:   │   │   AGENT 2:   │    │   AGENT 3:   │    │   AGENT 4:   │
│  Live Market │   │ Dynamic P&L  │    │ 16-Page PDF  │    │  Reviewer &  │
│  Data Scout  │   │  & Financial │    │ Vector Chart │    │ Multi-Channel│
│ (Scrapes NSE │   │Model Builder │    │ Engine (7 HD │    │  Dispatcher  │
│  & yfinance) │   │ (Tesla-Style)│    │  Matplotlib) │    │(Email+WA Bot)│
└──────────────┘   └──────────────┘    └──────────────┘    └──────────────┘
```

1. **Live Market Data Scout:** Extracts real-time NSE closing ticks, CMP, 52-week High/Low, Market Capitalization, Shares Outstanding, Trailing P/E, and Audited Annual Turnover.
2. **Dynamic Financial Modeler:** Resolves Sector DNA, compiles the 10-tab articulated model, applies OpenXML `<c:manualLayout>` injections, and formats numbers strictly to standard OpenXML specifications (zero Excel repair warnings).
3. **Publication-Grade PDF Compiler:** Produces 7 vector charts and assembles a 16-page ReportLab initiation report with dynamic headers and zero font rendering glitches.
4. **QA & Multi-Channel Dispatcher:** Verifies Balance Sheet zero-audit checks (`Assets - (Liabilities + Equity) = 0.00`), builds mobile executive digests, and triggers SMTP delivery.

---

## 🧠 Continuous Learning Loop (`memory/`)

The agent includes an autonomous, persistent learning system that starts completely clean from **0 (tabula rasa)**:

- **`memory/learnings.json`**: A local, persistent JSON memory bank. Automatically persists sector-specific accounting rules, custom working capital baselines, and past research history across runs.
- **`memory/memory_manager.py`**: Built-in memory manager that reads historical calibrations before generating models and logs completed runs.
- **Teaching the Agent via CLI**:
  ```bash
  python generate_equity_report.py --ticker TCS.NS --learn "sector:IT:dio=0"
  ```

---

## 🏗️ Dual-Engine Sector Architecture

The agent automatically resolves the target company's business archetype:

### 1. IT Services & Digital Tech Suite (e.g. TCS, Infosys, HCL Tech, Wipro)
- **Zero Physical Inventory**: Working capital correctly sets `Inventory Days (DIO)` to `0.0 Days (N/A - Services)` and models `Unbilled Revenue Days` (~28 days).
- **Service P&L Cost Breakdown**: Replaces material COGS with `Employee Benefit Expenses (Personnel)` (~56% of revenue) and `Subcontracting & SG&A Overheads` (~20%).
- **Asset-Light DuPont ROE**: Reflects authentic tech capital efficiency (~45%–50% ROE) driven by zero debt, high asset turnover, and high dividend payouts.

### 2. Corporate, Industrials & Consumer DCF Suite (e.g. Reliance, Titan, ITC, Tata Motors)
- **3-Statement Articulated Model**: Articulated 8-year Income Statement, Balance Sheet, and Cash Flow Statement with automated balance checks.
- **Fixed Asset & PP&E Schedule**: Capex roll-forward, gross block, and depreciation engine.
- **Working Capital & Cash Conversion Cycle (CCC)**: DSO, DIO, DPO days driving trade cycle.
- **Valuation Engines**: 10-Year Explicit Unlevered Discounted Cash Flow (DCF), Reverse DCF Implied Expectations, Sum-of-the-Parts (SOTP) Valuation, and 5-Method Football Field Matrix.

### 3. Banking & BFSI Valuation Suite (e.g. HDFC Bank, ICICI Bank, SBI, Kotak)
- **Loan Portfolio Roll-Forward**: Retail, Wholesale, SME, Agriculture breakdown.
- **NII & NIM Margin Engine**: Net Interest Income, Yield on Advances, Cost of Funds.
- **Asset Quality (NPA) Module**: Gross NPA %, Net NPA %, Provision Coverage Ratio (PCR).
- **Valuation Engines**: 5-Year Explicit Dividend Discount Model (DDM) & Justified Price-to-Book (P/BV) 2-Way Sensitivity Matrix.

---

## 📊 10-Tab Executive Interactive Financial Model (`.xlsx`)

The generated Excel workbook is modeled directly after executive institutional models:
- **`Dashboard` Tab**: 
  - Top Company Banner with bold company identity.
  - **Cell `E6`**: Interactive Excel Data Validation Dropdown (`"FY23 (A), FY24 (A), FY25 (A), FY26E, FY27E, FY28E, FY29E, FY30E"`).
  - **Top 7 KPI Stat Cards**: Revenue, COGS, OPEX, Gross Profit, Net Profit, ROA %, and ROE % with active `% VS Pre Year` growth badges.
  - **6 Embedded Visual Charts**: All charts dynamically re-orient when the `E6` dropdown year changes.
- **OpenXML Compliance**: Clean numeric formatting (`#,##0`, `#,##0.0`, `#,##0.00`, `0.0%`, `0.0000`). Zero unquoted strings in `styles.xml`, guaranteeing **zero recovery or corruption popups**.
- **Full Statement Tabs**: Cover Page, Drivers, Segment Breakdown, Income Statement, PP&E Schedule, Working Capital, Cash Flow, Balance Sheet, CAPM & WACC, DCF Valuation.

---

## 🚀 Installation & Quick Start

### 1. Clone Repository
```bash
git clone https://github.com/siddheshumrigar-tech/equity-research-agent.git
cd equity-research-agent
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Generate Complete Institutional Package
```bash
# Research an IT Services Company (e.g. TCS)
python generate_equity_report.py --ticker TCS.NS --name "Tata Consultancy Services Limited" --sector IT_SERVICES

# Research a Conglomerate (e.g. Reliance)
python generate_equity_report.py --ticker RELIANCE.NS --name "Reliance Industries Limited" --sector Energy
```

### 4. Optional Email Dispatch
```bash
python generate_equity_report.py --ticker RELIANCE.NS --email your-email@example.com
```

---

## ⚖️ License

Distributed under the **MIT License**. See `LICENSE` for more information.

---

## 🛡️ Disclaimer

*This software is intended strictly for financial modeling, educational research, and quantitative analysis. It does not constitute investment advice. Consult a certified financial advisor before executing investment transactions.*
