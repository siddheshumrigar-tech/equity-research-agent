# 🏛️ Autonomous Institutional Equity Research & Valuation Agent

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Code Style: Institutional](https://img.shields.io/badge/Standard-Tier--1%20Institutional-1A365D.svg)]()

An autonomous, production-grade financial valuation and equity research agent. Automatically fetches exchange market data, computes comprehensive financial statement models, performs multi-method valuation, generates high-resolution Matplotlib charts, and compiles:

1. **A 16-Page Publication-Grade PDF Initiation Report** (Motilal Oswal / Goldman Sachs / Morgan Stanley institutional standard).
2. **A 10-Tab Executive Interactive Financial Model (`.xlsx`)** with pre-configured dynamic Excel Data Validation dropdowns, 7 KPI stat cards, and 6 dynamic charts.

---

## 🏗️ Dual-Engine Sector Architecture

The agent automatically resolves the target company's business archetype:

### 1. Banking & BFSI Valuation Suite
*Designed for Commercial Banks, NBFCs, and Financial Institutions (e.g. HDFC Bank, ICICI Bank, SBI, Kotak, Axis, Bajaj Finance).*
- **Loan Portfolio Roll-Forward**: Retail, Wholesale, SME, Agriculture breakdown.
- **NII & NIM Margin Engine**: Net Interest Income, Yield on Advances, Cost of Funds.
- **Asset Quality (NPA) Module**: Gross NPA %, Net NPA %, Provision Coverage Ratio (PCR).
- **Capital Adequacy**: Tier-1 CRAR capital roll-forward and DuPont ROE decomposition.
- **Valuation Engines**: 5-Year Explicit Dividend Discount Model (DDM) & Justified Price-to-Book (P/BV) 2-Way Sensitivity Matrix.

### 2. Corporate & Industrials DCF Suite
*Designed for Non-Financial Corporates (e.g. Reliance Industries, Titan, ITC, TCS, Tata Motors, Sun Pharma).*
- **3-Statement Articulated Model**: Articulated 8-year Income Statement, Balance Sheet, and Cash Flow Statement with automated balance checks.
- **Fixed Asset & PP&E Schedule**: Capex roll-forward, gross block, and depreciation engine.
- **Working Capital & Cash Conversion Cycle (CCC)**: DSO, DIO, DPO days.
- **Valuation Engines**: 10-Year Explicit Unlevered Discounted Cash Flow (DCF), Reverse DCF Implied Expectations, Sum-of-the-Parts (SOTP) Valuation, and 5-Method Football Field Matrix.

---

## 📑 16-Page Master Institutional PDF Report Layout

Every report is compiled with high text density, structured financial tables, zero wasted whitespace, running headers, and `Page X of 16` footers:

- **Page 1**: Cover & Institutional Executive Dashboard (CMP, Target Price, MOS %, 1-Yr Stock vs. NIFTY 50 Chart).
- **Page 2**: Executive Summary & Core Investment Thesis (3 Strategic Pillars & Catalyst Matrix Table).
- **Page 3**: Corporate Architecture, Business Flywheel & Strategic Milestone Timeline.
- **Page 4**: Segment Deep-Dive #1: Primary Technology/Digital Services (Jio Subscriber & ARPU Curve Chart).
- **Page 5**: Segment Deep-Dive #2: Consumer & Retail Footprint (Segment EBITDA Donut Chart).
- **Page 6**: Segment Deep-Dive #3: Core Industrial / Petrochemical Operations (Revenue & EBITDA Trajectory Chart).
- **Page 7**: Segment Deep-Dive #4: Future Growth Engines & Cleantech Gigafactories.
- **Page 8**: Competitive Moats & Porter's Five Forces Deep-Dive Matrix.
- **Page 9**: Macro Landscape, Demographic Inflection & Government PLI Incentive Schemes.
- **Page 10**: 5-Year Historical & Projected Common-Size Financial Statements.
- **Page 11**: DuPont 5-Stage ROE Decomposition & Capital Efficiency (DuPont Driver Chart).
- **Page 12**: Working Capital, Cash Conversion Cycle & Capex Peak vs. Free Cash Flow Inflection Chart.
- **Page 13**: Sum-of-the-Parts (SOTP) Valuation & Multi-Model Football Field Chart.
- **Page 14**: 10-Year Explicit DCF, Reverse DCF Expectations & 2-Way WACC vs. Growth Sensitivity Matrix.
- **Page 15**: Risk Governance Matrix, Bear/Base/Bull Scenarios & Staggered Accumulation Tranches.
- **Page 16**: Floor Trader Technical Pivot Filters (R2, R1, P, S1, S2) & SEBI Statutory Disclosures.

---

## 📊 10-Tab Executive Interactive Financial Model (`.xlsx`)

The generated Excel workbook is modeled directly after executive institutional models:
- **`Dashboard` Tab**: 
  - Top Company Banner with bold company identity.
  - **Cell `E6`**: Interactive Excel Data Validation Dropdown (`"FY23 (A), FY24 (A), FY25 (A), FY26E, FY27E, FY28E, FY29E, FY30E"`).
  - **Top 7 KPI Stat Cards**: Revenue, COGS, OPEX, Gross Profit, Net Profit, ROA %, and ROE % with active `% VS Pre Year` growth badges.
  - **6 Embedded Visual Charts**: All charts dynamically re-orient when the `E6` dropdown year changes.
- **`Dashboard_Engine` Tab**: Underlying index-match lookup formulas dynamically powering cards and charts without VBA or external macros.
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
python generate_equity_report.py --ticker RELIANCE.NS --name "Reliance Industries Limited" --cmp 1302.50 --sector Energy
```

### 4. Optional Email Dispatch
To automatically dispatch the output `.xlsx` and `.pdf` to your email:
```bash
python generate_equity_report.py --ticker RELIANCE.NS --email your-email@example.com
```

---

## ⚖️ License

Distributed under the **MIT License**. See `LICENSE` for more information.

---

## 🛡️ Disclaimer

*This software is intended strictly for financial modeling, educational research, and quantitative analysis. It does not constitute investment advice. Consult a certified financial advisor before executing investment transactions.*
