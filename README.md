# 🔍 Financial Fraud Detector

> AI-powered financial statement analysis using Claude AI and logistic regression — no server, no installation, runs entirely in your browser.

---

## 📌 Overview

Upload any financial statement and this tool will:

1. Use **Claude AI** to extract 50 forensic financial metrics from the document
2. Run a **logistic regression model** trained on 650 real companies
3. Return a **fraud probability score** with flagged anomalies sorted by severity

---

## 🌐 Live Web Application

Access the deployed web application here:

[Open Financial Fraud Detector](https://basil2005assaf-hue.github.io/Fraud-Detector-AI/)

### Deploying via GitHub Pages

To host the application using GitHub Pages:

1. Upload the project files to a GitHub repository
2. Ensure the main application file is named `index.html`
3. Navigate to:
   - **Repository → Settings → Pages**
4. Under **Build and deployment**:
   - Set **Source** to `Deploy from a branch`
   - Select:
     - Branch: `main`
     - Folder: `/ (root)`
5. Click **Save**

GitHub will automatically generate a public website URL for the application.

---

## ✨ Features

- **Drag & drop upload** — supports PDF, DOCX, PNG, JPEG, TXT
- **50 extracted metrics** — Beneish M-Score, Altman Z-Score, Benford deviation, accrual ratio, cash flow ratios, governance score, NLP sentiment, and more
- **Severity classification** — Critical / High / Medium problem flags with plain-English explanations
- **Feature contributions** — see exactly which metrics drove the model's verdict
- **API key saved in browser** — enter once, never re-enter unless you switch devices
- **Zero dependencies** — single HTML file, open it and go

---

## 🚀 Quick Start

### 1. Get an Anthropic API Key

- Go to https://console.anthropic.com
- Sign up or log in → navigate to **API Keys** → **Create Key**
- Copy your key (starts with `sk-ant-`)
- Add billing credit at **Settings → Billing** (minimum $5)

> ⚠️ Keep your API key private. Never share it or post it publicly.

### 2. Open the App

**Option A — GitHub Pages (recommended):**

[Launch Web Application](https://basil2005assaf-hue.github.io/Fraud-Detector-AI/)

**Option B — Local:**

- Download `index.html`
- Open it in any browser (Chrome, Firefox, Edge, Safari)

### 3. Enter Your API Key

- Paste your `sk-ant-...` key into the **API Key** field
- Click **Save** — stored in your browser only, never sent to any server

### 4. Analyze

- Upload a financial statement
- Click **▶ Extract & Predict**
- Review the verdict, flagged issues, and metric breakdown

---

## 💰 Cost

| Model | Input | Output |
|---|---|---|
| Claude Sonnet 4.6 | $3.00 / 1M tokens | $15.00 / 1M tokens |

Each analysis uses approximately 3,000–6,000 input tokens and 1,000–1,500 output tokens.

| Credit | Estimated Analyses |
|---|---|
| $5 | ~125–250 |
| $10 | ~250–500 |
| $20 | ~500–1,000 |

---

## 📊 Model Details

| Property | Value |
|---|---|
| Algorithm | Logistic Regression |
| Dataset | 650 companies |
| Train / Test Split | 60% / 40% (389 / 261 samples) |
| Accuracy | 100% |
| ROC-AUC | 1.00 |
| Features | 50 |
| Extraction Model | Claude Sonnet 4.6 (Anthropic) |

---

## 🧮 Metrics Extracted

<details>
<summary>Click to expand all 50 metrics</summary>

| # | Metric | Description |
|---|---|---|
| 1 | Revenue | Total revenue (thousands USD) |
| 2 | Net_Income | Net income, can be negative |
| 3 | Operating_Cash_Flow | Cash from operations |
| 4 | Total_Assets | Total asset base |
| 5 | Total_Debt | Total debt obligations |
| 6 | Current_Ratio | Current assets / current liabilities |
| 7 | Debt_to_Equity | Leverage ratio |
| 8 | Gross_Margin | Gross profit margin (0–1) |
| 9 | ROA | Return on assets |
| 10 | ROE | Return on equity |
| 11 | Asset_Turnover | Revenue efficiency |
| 12 | Free_Cash_Flow | Operating CF minus capex |
| 13 | Benford_Deviation | Digit distribution anomaly score (0–1) |
| 14 | Audit_Opinion_Risk | Qualified/adverse opinion flag (0/1) |
| 15 | Insider_Transactions | Insider trading flag (0/1) |
| 16 | Related_Party_Transactions | Related-party flag (0/1) |
| 17 | Governance_Score | Corporate governance score (0–100) |
| 18 | Internal_Control_Weakness | Material weakness flag (0/1) |
| 19 | NLP_Sentiment_Score | Filing sentiment score (−1 to 1) |
| 20 | Fraud_Risk_Score | Composite fraud risk (0–1) |
| 21 | Operating_Margin | Operating profit margin |
| 22 | EBITDA | Earnings before interest, tax, depreciation |
| 23 | EPS | Earnings per share |
| 24 | SGA_Expense | Selling, general & admin expense |
| 25 | RD_Expense | Research & development expense |
| 26 | Interest_Coverage_Ratio | EBIT / interest expense |
| 27 | Current_Assets | Short-term assets |
| 28 | Current_Liabilities | Short-term liabilities |
| 29 | Inventory | Inventory balance |
| 30 | Accounts_Receivable | AR balance |
| 31 | Goodwill | Goodwill on balance sheet |
| 32 | Retained_Earnings | Accumulated retained earnings |
| 33 | Shareholder_Equity | Total equity |
| 34 | Capital_Expenditures | Capex spending |
| 35 | CashFlow_to_Debt | Cash flow coverage of debt |
| 36 | CFO_to_NetIncome_Ratio | Cash earnings quality ratio |
| 37 | Financing_Dependence | Reliance on external financing (0–1) |
| 38 | Beneish_MScore | Earnings manipulation score |
| 39 | Altman_ZScore | Bankruptcy distress score |
| 40 | Days_Sales_Outstanding | AR collection period (days) |
| 41 | Inventory_Growth_Rate | YoY inventory growth |
| 42 | Revenue_Growth_Rate | YoY revenue growth |
| 43 | Accrual_Ratio | Accrual-based earnings ratio |
| 44 | Auditor_Change_Frequency | Frequent auditor changes flag (0/1) |
| 45 | Restatement_History | Prior restatement flag (0/1) |
| 46 | Executive_Turnover | Leadership turnover flag (0/1) |
| 47 | Uncertainty_Word_Count | Count of uncertainty language in filings |
| 48 | Readability_Score | Flesch readability score (0–100) |
| 49 | Risk_Factor_Count | Number of disclosed risk factors |
| 50 | MDA_Sentiment | MD&A section sentiment (−1 to 1) |

</details>

---

## ⚠️ Troubleshooting

| Error | Fix |
|---|---|
| `Invalid API key` | Check your key starts with `sk-ant-` and is saved correctly |
| `Rate limit exceeded` | Wait 60 seconds and try again, or add more credit to move to a higher tier |
| `Model not found` | Make sure the model ID in the code is `claude-sonnet-4-6` |
| File won't upload | Only PDF, DOCX, PNG, JPEG, TXT are supported |

---

## 🔒 Privacy

- Your API key is stored in **your browser only** (localStorage)
- Documents are sent directly from your browser to Anthropic's API — no third-party server involved
- Nothing is stored or logged outside of Anthropic's standard API logging

---

## 📁 Repository Structure

```text
fraud-detector/
├── index.html        # Complete app — single file, open in any browser
└── README.md         # This file
```

---

## 📄 License

MIT — free to use, modify, and distribute.

---

## ⚖️ Disclaimer

This tool is for educational and research purposes only. It is not a substitute for professional forensic accounting, certified audit services, or legal review. All results should be interpreted alongside qualified professional judgment.
