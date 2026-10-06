# Antigravity Foreign Asset Skill (Schedule FA, Form 67 & RSU Capital Gains)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Antigravity Skill](https://img.shields.io/badge/Antigravity-Agent%20Skill-blue.svg)](https://github.com/tbhavar/antigravity-foreign-asset-skill)
[![Compliance: Indian Income Tax](https://img.shields.io/badge/Compliance-Rule%20115%20%7C%20DTAA%20Sec%2090-green.svg)](https://incometaxindia.gov.in)

An enterprise-grade, multi-agent AI skill for **Antigravity (Gemini 3.8 Flash / Pro)** that automates:
1. **Schedule FA (Table A3)**: Foreign equity asset declarations, initial values, intra-year peak valuations, and period-end closing values.
2. **Form 67 (DTAA Relief u/s 90)**: Foreign tax credit computation on US/foreign dividends under Article 10 of Double Taxation Avoidance Agreements.
3. **RSU / Share Capital Gains**: Short-term and Long-term capital gains calculation complying with Rule 115 SBI TT Buying Rates, preserving live Excel formulas.

---

## 🌟 Key Features

- **Multi-Agent Collaborative Review**:
  - **Agent 1: Ingestion & Forex Normalizer**: Extracts vesting tranches, 1042-S dividends, and resolves Rule 115 SBI TTBR exchange rates.
  - **Agent 2: Statutory Tax Computation Specialist**: Drafts `RSU Gain.xlsx`, `Form-67.xlsx`, and `Schedule FA.xlsx` while strictly keeping Excel formulas intact.
  - **Agent 3: Senior CA Auditor**: Cross-reconciles schedules, verifies peak valuation logic ($\text{Peak} \ge \text{Initial}, \text{Closing}$), confirms formula integrity, and issues audit sign-off.
- **Configurable Period Modes**: Toggle between **Indian Financial Year (01-Apr to 31-Mar)** and **Foreign Calendar Year (01-Jan to 31-Dec)**.
- **Rule 115 Forex Engine**: Automatically fetches SBI TT Buying Rate on the last day of the month preceding the transaction month, with built-in holiday fallback.
- **Live Formula Preservation**: Never overwrites Excel calculations with static values (`=MIN()`, `=ROUND()`, `=SUM()` preserved).
- **Non-Coder Friendly**: Simple parameter tweaks and intuitive instructions.

---

## 📁 Repository Structure

```
antigravity-foreign-asset-skill/
├── SKILL.md                                 # Main Antigravity Skill Definition & Instructions
├── README.md                                # Repository Documentation
├── LICENSE                                  # MIT License
├── resources/                               # Prescribed Templates & Data
│   ├── Form-67_template.xlsx               # Blank Form 67 Template with Formulas
│   ├── RSU-CG-Calculations_template.xlsx   # Capital Gains Computation Workbook
│   ├── Schedule_FA_template.xlsx           # Schedule FA Table A3 Workbook
│   └── SBI_REFERENCE_RATES_USD.csv         # Historical SBI TT Buying & Selling Rates
└── scripts/                                 # Modular Python Automation Engines
    ├── sbi_forex_rates.py                  # Rule 115 Forex Resolver
    ├── compute_rsu_capital_gains.py        # RSU Capital Gains Engine
    ├── compute_form_67.py                  # Form 67 DTAA Relief Engine
    ├── populate_schedule_fa.py             # Schedule FA Table A3 Engine
    └── audit_and_reconcile_all.py          # Senior CA Audit & Quality Review Engine
```

---

## 🚀 Quick Start

### 1. Installation into Antigravity
Clone or copy this repository into your Antigravity skills directory:
```bash
# Global skills directory
cp -r antigravity-foreign-asset-skill ~/.gemini/config/skills/foreign-assets-itr-reporting
```

### 2. Client Folder Preparation
Place client tax documents in their working folder:
- Broker Statement (e.g. `Custom-transaction-summary-Fidelity-NetBenefits.pdf`)
- US Tax Statement (`Form-1042S.pdf`)
- Exchange rates file (`SBI_REFERENCE_RATES_USD.csv`)
- Target Excel templates (`Form-67.xlsx`, `RSU Gain.xlsx`, `Schedule FA.xlsx`)

### 3. Trigger Prompt
Execute the skill with the prompt:
```text
Please act as an expert chartered accountant, having deep knowledge in income tax,
specifically regarding foreign assets, holdings, declarations, and double taxation
avoidance agreements, like claiming the benefit of DTAA by filing Form 67. And also
expert in calculating any gains arising on transfer of these assets. Since you have
this knowledge, I want you to compute all the Schedule FA related disclosure for this client.
Please directly execute your plan and don't wait for my approval.
```

---

## ⚙️ Input Parameters

You can customize execution directly in your prompt:

| Parameter | Default | Description |
|---|---|---|
| `FINANCIAL_YEAR` | `2025-26` | Target Indian Financial Year |
| `ASSESSMENT_YEAR` | `2026-27` | Target Assessment Year |
| `SCHEDULE_FA_PERIOD` | `FINANCIAL_YEAR` | Toggle: `FINANCIAL_YEAR` (01-Apr to 31-Mar) or `CALENDAR_YEAR` (01-Jan to 31-Dec) |
| `FOREIGN_CURRENCY` | `USD` | Foreign asset currency (USD, EUR, GBP, CAD) |
| `INDIAN_TAX_SLAB_RATE` | `31.2%` | Client marginal rate (30% + 4% cess under New Regime) |
| `PRESERVE_FORMULAS` | `True` | Enforce dynamic Excel formulas |

---

## ⚖️ Legal & Statutory Compliance

- **Rule 115, Income-tax Rules, 1962**: Conversion of foreign dividend income and capital gains using SBI TT Buying Rates on the last day of the preceding month.
- **Section 90 read with Rule 128**: Double taxation relief capped at lower of Indian tax payable and foreign tax paid.
- **Section 45 & 48**: RSU capital gains with 24-month threshold for foreign unlisted shares. Transfers post 23-Jul-2024 taxed at flat 12.5% for LTCG.

---

## 📄 License
This project is licensed under the [MIT License](LICENSE).
