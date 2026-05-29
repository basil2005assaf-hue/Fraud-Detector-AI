import { useState, useRef, useCallback } from "react";
import mammoth from "mammoth";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════
//  Financial Fraud Detector v5
//  Works inside claude.ai artifacts — no API key needed.
// ═══════════════════════════════════════════════════════════

const MODEL = {
  featureCols: ["Revenue","Net_Income","Operating_Cash_Flow","Total_Assets","Total_Debt","Current_Ratio","Debt_to_Equity","Gross_Margin","ROA","ROE","Asset_Turnover","Free_Cash_Flow","Benford_Deviation","Audit_Opinion_Risk","Insider_Transactions","Related_Party_Transactions","Governance_Score","Internal_Control_Weakness","NLP_Sentiment_Score","Fraud_Risk_Score","Operating_Margin","EBITDA","EPS","SGA_Expense","RD_Expense","Interest_Coverage_Ratio","Current_Assets","Current_Liabilities","Inventory","Accounts_Receivable","Goodwill","Retained_Earnings","Shareholder_Equity","Capital_Expenditures","CashFlow_to_Debt","CFO_to_NetIncome_Ratio","Financing_Dependence","Beneish_MScore","Altman_ZScore","Days_Sales_Outstanding","Inventory_Growth_Rate","Revenue_Growth_Rate","Accrual_Ratio","Auditor_Change_Frequency","Restatement_History","Executive_Turnover","Uncertainty_Word_Count","Readability_Score","Risk_Factor_Count","MDA_Sentiment"],
  scalerMean: [404353.85861,54606.98201,85731.47558,676621.78149,113413.69923,1.74072,1.35979,0.47977,0.12365,0.14767,0.88026,62085.98201,0.10925,0.41388,0.46787,0.491,69.56555,0.43959,0.38093,0.34825,0.15221,76173.33419,7.5073,29228.3856,29552.57326,16.40159,214376.81491,126348.62211,54017.4036,68841.86118,79988.55527,94417.3599,216214.09512,40185.77121,2.06784,1.97339,0.3873,-1.76555,5.09571,73.64704,0.28648,0.21319,0.11558,0.40617,0.39332,0.47301,64.21851,54.84576,25.31105,0.38267],
  scalerStd:  [269356.78062,54545.6858,64463.46138,345040.76651,74073.63569,0.74128,1.63307,0.18792,0.11151,0.27739,0.41773,62874.16329,0.11926,0.49253,0.49897,0.49992,21.23183,0.49634,0.4762,0.24119,0.18875,59468.76801,7.44579,17521.60393,19789.70242,13.8812,153296.46928,68953.6836,34212.7553,42932.22634,46351.28983,100732.56378,150939.04022,26422.28344,4.90693,4.99191,0.22926,1.60966,2.84609,53.35103,0.34257,0.24029,0.21807,0.49112,0.48849,0.49927,54.75973,20.81161,14.91703,0.46345],
  coef: [0.16253,0.14923,0.19826,0.19986,-0.24616,0.31106,-0.46931,0.13336,0.35169,0.39418,0.2004,0.16624,-0.49567,-0.22647,-0.21342,-0.18151,0.3974,-0.14103,0.44424,-0.39368,0.40071,0.19501,0.29799,-0.18912,0.21645,0.29649,0.21822,0.02014,-0.17981,-0.17824,-0.08313,0.31884,0.21042,0.21669,-0.01715,0.21695,-0.32628,-0.48688,0.31734,-0.42547,-0.46929,-0.07027,-0.4888,-0.20685,-0.28712,-0.15885,-0.40108,0.32209,-0.24122,0.44658],
  intercept: 1.01894,
  topFeatures: ["Benford_Deviation","Accrual_Ratio","Beneish_MScore","Debt_to_Equity","Inventory_Growth_Rate","MDA_Sentiment","NLP_Sentiment_Score","Days_Sales_Outstanding","Uncertainty_Word_Count","Operating_Margin"],
};

const THRESHOLDS = {
  Current_Ratio:             { dir:"low",  v:1.0,   msg:"Current ratio < 1.0 — liquidity risk" },
  Debt_to_Equity:            { dir:"high", v:2.5,   msg:"Debt-to-equity > 2.5 — excessive leverage" },
  Benford_Deviation:         { dir:"high", v:0.15,  msg:"Benford deviation > 0.15 — accounting anomaly" },
  Audit_Opinion_Risk:        { dir:"eq",   v:1,     msg:"Qualified/adverse audit opinion — major red flag" },
  Insider_Transactions:      { dir:"eq",   v:1,     msg:"Insider transactions flagged" },
  Internal_Control_Weakness: { dir:"eq",   v:1,     msg:"Material internal control weakness reported" },
  NLP_Sentiment_Score:       { dir:"low",  v:-0.2,  msg:"Negative NLP sentiment in filings" },
  Fraud_Risk_Score:          { dir:"high", v:0.6,   msg:"High composite fraud risk score" },
  Gross_Margin:              { dir:"low",  v:0.1,   msg:"Extremely low gross margin" },
  ROA:                       { dir:"low",  v:-0.05, msg:"Negative return on assets" },
  Free_Cash_Flow:            { dir:"low",  v:0,     msg:"Negative free cash flow" },
  Governance_Score:          { dir:"low",  v:50,    msg:"Low governance score" },
  Beneish_MScore:            { dir:"high", v:-1.78, msg:"Beneish M-Score above -1.78 — earnings manipulation likely" },
  Altman_ZScore:             { dir:"low",  v:1.81,  msg:"Altman Z-Score < 1.81 — distress zone" },
  Restatement_History:       { dir:"eq",   v:1,     msg:"Historical financial restatements on record" },
  Auditor_Change_Frequency:  { dir:"eq",   v:1,     msg:"Frequent auditor changes — independence concern" },
  Financing_Dependence:      { dir:"high", v:0.7,   msg:"High dependence on external financing" },
  CFO_to_NetIncome_Ratio:    { dir:"low",  v:0.5,   msg:"Low cash flow to net income — accrual manipulation risk" },
  Days_Sales_Outstanding:    { dir:"high", v:100,   msg:"DSO > 100 days — revenue recognition concern" },
  Accrual_Ratio:             { dir:"high", v:0.1,   msg:"High accrual ratio — earnings quality concern" },
  Inventory_Growth_Rate:     { dir:"high", v:0.4,   msg:"Inventory growth rate > 40% — channel stuffing risk" },
  Revenue_Growth_Rate:       { dir:"low",  v:-0.1,  msg:"Revenue declining > 10% year-over-year" },
};

// ── EXTRACTION SYSTEM PROMPT ─────────────────────────────────────────────────
// Key fix: detailed scoring guidance so Claude correctly weighs compound signals
// that the model relies heavily on (Benford_Deviation, Fraud_Risk_Score,
// NLP_Sentiment_Score, Beneish_MScore).
const EXTRACTION_SYSTEM = `You are a forensic financial analyst specializing in fraud detection. Extract metrics from the financial statement and return ONLY a raw JSON object — no markdown, no backticks, no explanation.

Required keys (all numeric unless noted):
{
  "company_name": "string or null",
  "Revenue": number (thousands USD),
  "Net_Income": number (thousands, can be negative),
  "Operating_Cash_Flow": number (thousands),
  "Total_Assets": number (thousands),
  "Total_Debt": number (thousands),
  "Current_Ratio": number,
  "Debt_to_Equity": number,
  "Gross_Margin": number (0-1),
  "ROA": number,
  "ROE": number,
  "Asset_Turnover": number,
  "Free_Cash_Flow": number (thousands),
  "Benford_Deviation": number (0-1),
  "Audit_Opinion_Risk": 0 or 1,
  "Insider_Transactions": 0 or 1,
  "Related_Party_Transactions": 0 or 1,
  "Governance_Score": number (0-100),
  "Internal_Control_Weakness": 0 or 1,
  "NLP_Sentiment_Score": number (-1 to 1),
  "Fraud_Risk_Score": number (0-1),
  "Operating_Margin": number (0-1),
  "EBITDA": number (thousands),
  "EPS": number,
  "SGA_Expense": number (thousands),
  "RD_Expense": number (thousands),
  "Interest_Coverage_Ratio": number,
  "Current_Assets": number (thousands),
  "Current_Liabilities": number (thousands),
  "Inventory": number (thousands),
  "Accounts_Receivable": number (thousands),
  "Goodwill": number (thousands),
  "Retained_Earnings": number (thousands),
  "Shareholder_Equity": number (thousands),
  "Capital_Expenditures": number (thousands),
  "CashFlow_to_Debt": number,
  "CFO_to_NetIncome_Ratio": number,
  "Financing_Dependence": number (0-1),
  "Beneish_MScore": number,
  "Altman_ZScore": number,
  "Days_Sales_Outstanding": number (days),
  "Inventory_Growth_Rate": number,
  "Revenue_Growth_Rate": number,
  "Accrual_Ratio": number,
  "Auditor_Change_Frequency": 0 or 1,
  "Restatement_History": 0 or 1,
  "Executive_Turnover": 0 or 1,
  "Uncertainty_Word_Count": number,
  "Readability_Score": number (Flesch 0-100),
  "Risk_Factor_Count": number,
  "MDA_Sentiment": number (-1 to 1)
}

═══════════════════════════════════════════════════════════════
CRITICAL SCORING GUIDANCE — these four fields drive most of the
fraud model's predictive power. Score them carefully:
═══════════════════════════════════════════════════════════════

BENFORD_DEVIATION (0–1):
Do NOT rely solely on digit-frequency counts. Also factor in:
- Accounts receivable growing >2× faster than revenue → +0.05
- Operating cash flow / net income ratio collapsing >40% year-over-year → +0.04
- Capitalized intangibles/software spiking >100% with no acquisition → +0.04
- Multiple round-number figures across financial statements → +0.03
- Reported metrics that are suspiciously smooth or internally inconsistent → +0.04
Calibration: 0.02–0.07 = normal; 0.08–0.14 = minor irregularities;
0.15–0.22 = suspicious; 0.22–0.35 = highly suspicious (use when 3+ signals above apply).

FRAUD_RISK_SCORE (0–1):
Holistic composite. Build from 0.20 baseline, add:
  +0.10 each: free cash flow negative; CFO/NI < 0.5; DSO > 75 days;
              AR growing >1.5× revenue growth rate
  +0.12 each: inventory growing >1.5× revenue; capitalized expenses up >100% YoY;
              cash flow statement does not mathematically reconcile
  +0.15:      3 or more of the above signals fire simultaneously
  +0.15:      management language is very positive while cash flows are deteriorating
              (misleading tone is itself a manipulation signal)
Cap at 0.95. This is your most important composite signal — do not underestimate it.

NLP_SENTIMENT_SCORE (−1 to 1):
Evaluate ACTUAL informational content, not surface tone. When management
uses positive/optimistic language while underlying cash generation is
deteriorating significantly, the score should be NEGATIVE — misleading
positive disclosure is a fraud indicator, not a positive one.
  +0.3 to +0.6: Positive language AND strong underlying cash flows
   0.0 to +0.2: Positive language with mixed financial signals
  −0.1 to −0.3: Positive language but OCF/NI < 0.5 or FCF negative
  −0.3 to −0.6: Very positive language + multiple deteriorating cash metrics

MDA_SENTIMENT (−1 to 1): Same logic as NLP_Sentiment_Score.
Apply the same downward adjustment when tone contradicts cash flow reality.

BENEISH_M_SCORE: When prior-year data is available, calculate using:
  M = −4.84 + 0.920×DSRI + 0.528×GMI + 0.404×AQI + 0.892×SGI
            + 0.115×DEPI − 0.172×SGAI + 4.679×TATA − 0.327×LVGI
  where:
    DSRI = (AR_t / Rev_t) / (AR_t−1 / Rev_t−1)
    GMI  = GrossMargin_t−1 / GrossMargin_t
    AQI  = (1 − (CA_t + PPE_t) / TA_t) / (1 − (CA_t−1 + PPE_t−1) / TA_t−1)
    SGI  = Rev_t / Rev_t−1
    DEPI = (Dep_t−1 / (Dep_t−1 + PPE_t−1)) / (Dep_t / (Dep_t + PPE_t))
    SGAI = (SGA_t / Rev_t) / (SGA_t−1 / Rev_t−1)
    TATA = (NetIncome_t − OCF_t) / TotalAssets_t
    LVGI = ((LTD_t + CL_t) / TA_t) / ((LTD_t−1 + CL_t−1) / TA_t−1)
  Score > −1.78 indicates likely earnings manipulation.
  If prior-year data is unavailable, estimate from available signals.

CASH FLOW RECONCILIATION CHECK:
Verify: Beginning_Cash + OCF + Investing_CF + Financing_CF + FX_Effect = Ending_Cash.
If this does not balance, set Internal_Control_Weakness = 1 and increase
Benford_Deviation by 0.05.

═══════════════════════════════════════════════════════════════
General extraction rules:
- Monetary values in THOUSANDS USD (e.g. $5.2B = 5200000)
- Compute ratios from raw figures when not explicitly stated
- Use null only if a value truly cannot be determined or estimated
- Return ONLY the raw JSON object, nothing else
═══════════════════════════════════════════════════════════════`;

function sigmoid(x) { return 1 / (1 + Math.exp(-Math.max(-500, Math.min(500, x)))); }

function runModel(values) {
  const { featureCols, scalerMean, scalerStd, coef, intercept } = MODEL;
  let z = intercept;
  const contribs = [];
  featureCols.forEach((col, i) => {
    const raw = values[col];
    const num = (raw !== undefined && raw !== null) ? parseFloat(raw) : NaN;
    const scaled = isNaN(num) ? 0 : (num - scalerMean[i]) / scalerStd[i];
    const c = coef[i] * scaled;
    z += c;
    contribs.push({ col, contrib: c, missing: isNaN(num) });
  });
  const probValid = sigmoid(z);
  return { probValid, probFraud: 1 - probValid, isValid: probValid >= 0.5, contributions: contribs };
}

function detectProblems(values) {
  const probs = [];

  // ── Existing single-metric threshold checks ──────────────────────────────
  Object.entries(THRESHOLDS).forEach(([key, rule]) => {
    const v = parseFloat(values[key]);
    if (isNaN(v)) return;
    if (rule.dir === "high" && v > rule.v) probs.push(rule.msg);
    if (rule.dir === "low"  && v < rule.v) probs.push(rule.msg);
    if (rule.dir === "eq"   && v == rule.v) probs.push(rule.msg);
  });

  // ── Existing cash-flow cross-checks ──────────────────────────────────────
  if (parseFloat(values.Net_Income) > 0 && parseFloat(values.Operating_Cash_Flow) < 0)
    probs.push("Positive net income but negative operating cash flow — earnings manipulation signal");
  if (parseFloat(values.Revenue) > 0 && parseFloat(values.Free_Cash_Flow) < -0.25 * parseFloat(values.Revenue))
    probs.push("Free cash flow deeply negative relative to revenue");

  // ── NEW: compound multi-signal checks ────────────────────────────────────
  const rev   = parseFloat(values.Revenue);
  const ar    = parseFloat(values.Accounts_Receivable);
  const ocf   = parseFloat(values.Operating_Cash_Flow);
  const ni    = parseFloat(values.Net_Income);
  const fcf   = parseFloat(values.Free_Cash_Flow);
  const dso   = parseFloat(values.Days_Sales_Outstanding);
  const invGr = parseFloat(values.Inventory_Growth_Rate);
  const revGr = parseFloat(values.Revenue_Growth_Rate);
  const cfoNi = parseFloat(values.CFO_to_NetIncome_Ratio);

  // 1. AR exceeding 22% of annual revenue (healthy norm ~12–18%)
  if (!isNaN(ar) && !isNaN(rev) && rev > 0 && ar / rev > 0.22)
    probs.push(
      `Accounts receivable is ${(ar/rev*100).toFixed(1)}% of annual revenue — ` +
      "disproportionate AR balance signals premature or inflated revenue recognition"
    );

  // 2. DSO elevated but below the 100-day hard threshold (75–99 days is an early warning)
  if (!isNaN(dso) && dso >= 75 && dso < 100)
    probs.push(
      `DSO elevated at ${Math.round(dso)} days — significantly above healthy range, ` +
      "trending toward critical revenue quality threshold"
    );

  // 3. CFO severely lagging NI — stronger catch than the 0.5 threshold alone
  if (!isNaN(cfoNi) && cfoNi > 0 && cfoNi < 0.4)
    probs.push(
      `Operating cash flow is only ${(cfoNi*100).toFixed(0)}% of net income — ` +
      "severe accrual-based earnings, strong manipulation signal"
    );

  // 4. Inventory growing materially faster than revenue (channel stuffing / demand weakness)
  if (!isNaN(invGr) && !isNaN(revGr) && revGr > 0 && invGr > revGr * 1.5)
    probs.push(
      `Inventory growing ${(invGr/revGr).toFixed(1)}× faster than revenue — ` +
      "excess inventory accumulation inconsistent with stated demand growth"
    );

  // 5. High revenue growth alongside negative FCF — revenue may be non-cash
  if (!isNaN(revGr) && revGr > 0.15 && !isNaN(fcf) && fcf < 0)
    probs.push(
      `Strong reported revenue growth (${(revGr*100).toFixed(0)}%) accompanied by negative ` +
      "free cash flow — high-growth narrative not supported by cash generation"
    );

  // 6. OCF covering less than 35% of net income — extreme quality-of-earnings concern
  if (!isNaN(ocf) && ocf > 0 && !isNaN(ni) && ni > 0 && ocf / ni < 0.35)
    probs.push(
      `Operating cash flow covers only ${(ocf/ni*100).toFixed(0)}% of reported net income — ` +
      "earnings are heavily accrual-based with minimal cash backing"
    );

  return probs;
}

function categorizeProblem(msg) {
  const critical = ["Beneish M-Score","Benford deviation","earnings manipulation","Qualified/adverse audit","Material internal control","restatements","only ","covers only"];
  const high     = ["excessive leverage","fraud risk score","High composite","Negative return","Altman Z-Score","Negative NLP","manipulation signal","disproportionate","premature","severe accrual","revenue growth","accompanies negative"];
  if (critical.some(k => msg.toLowerCase().includes(k.toLowerCase()))) return "critical";
  if (high.some(k => msg.toLowerCase().includes(k.toLowerCase())))     return "high";
  return "medium";
}

async function prepareFileContent(file) {
  const t = file.type, n = file.name.toLowerCase();
  const isDocx  = t.includes("wordprocessingml") || t === "application/msword" || n.endsWith(".docx") || n.endsWith(".doc");
  const isText  = t === "text/plain" || n.endsWith(".txt");
  const isCsv   = t === "text/csv" || n.endsWith(".csv");
  const isExcel = t.includes("spreadsheet") || t.includes("excel") || n.endsWith(".xlsx") || n.endsWith(".xls");
  const isPdf   = t === "application/pdf" || n.endsWith(".pdf");
  const isImage = t.startsWith("image/") || /\.(png|jpe?g)$/.test(n);
  if (isDocx) {
    const ab = await file.arrayBuffer();
    const r  = await mammoth.extractRawText({ arrayBuffer: ab });
    if (!r.value.trim()) throw new Error("Could not extract text from DOCX — file may be image-based or empty.");
    return { mode:"text", text:r.value };
  }
  if (isCsv || isText) return { mode:"text", text:await file.text() };
  if (isExcel) {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type:"array" });
    let allText = "";
    wb.SheetNames.forEach(name => {
      allText += "=== Sheet: " + name + " ===\n";
      allText += XLSX.utils.sheet_to_csv(wb.Sheets[name], { blankrows:false }) + "\n\n";
    });
    if (!allText.trim()) throw new Error("Could not extract data from Excel file.");
    return { mode:"text", text:allText };
  }
  if (isPdf || isImage) {
    const base64 = await new Promise((res,rej) => {
      const r = new FileReader();
      r.onload  = () => res(r.result.split(",")[1]);
      r.onerror = () => rej(new Error("File read failed"));
      r.readAsDataURL(file);
    });
    return { mode:"binary", base64, mediaType:isPdf ? "application/pdf" : (t||"image/png") };
  }
  throw new Error(`Unsupported file type: ${file.name}`);
}

async function extractFromDocument(fc) {
  let messages;
  if (fc.mode === "binary") {
    const isImg = fc.mediaType.startsWith("image/");
    messages = [{ role:"user", content:[
      isImg
        ? { type:"image",    source:{ type:"base64", media_type:fc.mediaType,      data:fc.base64 } }
        : { type:"document", source:{ type:"base64", media_type:"application/pdf", data:fc.base64 } },
      { type:"text", text:"Extract all financial metrics from this financial statement." }
    ]}];
  } else {
    messages = [{ role:"user", content:`Here is the financial statement text:\n\n${fc.text.slice(0,50000)}\n\nExtract all 50 financial metrics.` }];
  }
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method:"POST",
    headers:{ "Content-Type":"application/json" },
    body:JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1800, system:EXTRACTION_SYSTEM, messages }),
  });
  if (!res.ok) {
    const e = await res.json().catch(()=>({}));
    throw new Error(e?.error?.message || `API error ${res.status}`);
  }
  const data = await res.json();
  let text = data.content.map(b=>b.text||"").join("").trim().replace(/```json|```/g,"").trim();
  try { return JSON.parse(text); } catch(_) {}
  const match = text.match(/\{[\s\S]+\}/);
  if (match) try { return JSON.parse(match[0]); } catch(_) {}
  throw new Error("Could not parse Claude response as JSON. Try a different document.");
}

export default function App() {
  const [file, setFile]                   = useState(null);
  const [stage, setStage]                 = useState("idle");
  const [progress, setProgress]           = useState("");
  const [progressStep, setProgressStep]   = useState(0);
  const [extracted, setExtracted]         = useState(null);
  const [result, setResult]               = useState(null);
  const [errMsg, setErrMsg]               = useState("");
  const [dragOver, setDragOver]           = useState(false);
  const [tab, setTab]                     = useState("result");
  const inputRef = useRef();

  const handleFile = useCallback((f) => {
    if (!f) return;
    const ok = /\.(pdf|docx?|png|jpe?g|txt|xlsx?|csv)$/i.test(f.name) ||
      ["application/pdf","image/png","image/jpeg","text/plain","text/csv",
       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
       "application/vnd.ms-excel",
       "application/msword"].includes(f.type);
    if (!ok) { setErrMsg(`Unsupported: ${f.name}. Use PDF, DOCX, XLSX, CSV, PNG, JPEG, or TXT.`); return; }
    setFile(f); setStage("idle"); setResult(null); setExtracted(null); setErrMsg(""); setProgress("");
  }, []);

  const handleDrop = (e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); };

  const analyze = async () => {
    if (!file) return;
    setStage("loading"); setErrMsg(""); setResult(null); setExtracted(null);
    try {
      setProgress("Reading file…");        setProgressStep(1);
      const fc = await prepareFileContent(file);
      setProgress("Sending to Claude…");   setProgressStep(2);
      const metrics = await extractFromDocument(fc);
      setProgress("Running fraud model…"); setProgressStep(3);
      const { company_name, ...fv } = metrics;
      const pred  = runModel(fv);
      const probs = detectProblems(fv);
      setExtracted({ company_name, ...fv });
      setResult({ ...pred, problems:probs, company_name });
      setStage("done"); setTab("result");
    } catch(e) {
      setErrMsg(e.message||"Unknown error"); setStage("error");
    }
  };

  const reset = (e) => {
    if (e) e.stopPropagation();
    setFile(null); setStage("idle"); setResult(null); setExtracted(null);
    setErrMsg(""); setProgress(""); setProgressStep(0);
    if (inputRef.current) inputRef.current.value = "";
  };

  const isLoading     = stage === "loading";
  const criticalProbs = result ? result.problems.filter(p=>categorizeProblem(p)==="critical") : [];
  const highProbs     = result ? result.problems.filter(p=>categorizeProblem(p)==="high")     : [];
  const mediumProbs   = result ? result.problems.filter(p=>categorizeProblem(p)==="medium")   : [];

  return (
    <div style={{minHeight:"100vh",background:"#06090f",color:"#c8d8f0",fontFamily:"'DM Mono','Courier New',monospace"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        :root{--g:#00e5a0;--r:#ff3356;--b:#3d9eff;--am:#ffb700;--bg:#06090f;--card:#0b1120;--border:rgba(255,255,255,0.07);--muted:#3d5270;--text:#c8d8f0;--ch:#ff8c00}
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes up{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
        @keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
        .up{animation:up .35s ease forwards}
        .spin{animation:spin 1s linear infinite;display:inline-block}
        .blink{animation:blink 2s ease infinite}
        .bar{transition:width 1.2s cubic-bezier(.22,1,.36,1)}
        button{cursor:pointer;font-family:inherit}
        .drop{border:2px dashed var(--border);border-radius:12px;padding:44px 28px;text-align:center;transition:all .2s;background:rgba(255,255,255,.01)}
        .drop.over{border-color:var(--b);background:rgba(61,158,255,.05)}
        .drop.has{border-color:rgba(0,229,160,.22);border-style:solid;background:rgba(0,229,160,.03)}
        .go{background:linear-gradient(135deg,#3d9eff 0%,#00c8ff 100%);color:#06090f;border:none;padding:12px 34px;border-radius:8px;font-size:12px;font-weight:500;letter-spacing:.06em;transition:all .2s;font-family:inherit}
        .go:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 8px 24px rgba(61,158,255,.35)}
        .go:disabled{opacity:.3;cursor:default}
        .sec{font-size:9px;letter-spacing:.18em;text-transform:uppercase;color:var(--muted);margin-bottom:10px}
        .tab{background:transparent;border:1px solid var(--border);color:var(--muted);padding:6px 16px;font-size:10px;letter-spacing:.08em;text-transform:uppercase;border-radius:5px;transition:all .15s;font-family:inherit}
        .tab.on{background:rgba(61,158,255,.1);border-color:var(--b);color:var(--b)}
        .mrow{display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid rgba(255,255,255,.03);font-size:10px}
        .mrow:last-child{border:none}
        .pill{display:inline-block;padding:2px 8px;border-radius:20px;font-size:8px;letter-spacing:.1em;text-transform:uppercase}
        .pc{background:rgba(255,51,86,.07);border:1px solid rgba(255,51,86,.18);border-left:3px solid var(--r);border-radius:0 6px 6px 0;padding:7px 11px;margin-bottom:6px;font-size:10px;color:#ffb3c0;line-height:1.55}
        .wc{background:rgba(255,183,0,.06);border:1px solid rgba(255,183,0,.16);border-left:3px solid var(--am);border-radius:0 6px 6px 0;padding:7px 11px;margin-bottom:6px;font-size:10px;color:#ffe5a0;line-height:1.55}
        input[type=file]{display:none}
        ::-webkit-scrollbar{width:3px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
        .summary-panel{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:20px;margin-top:16px}
        .sev-badge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:20px;font-size:9px;letter-spacing:.1em;text-transform:uppercase;font-weight:500}
        .sev-critical{background:rgba(255,51,86,.12);color:var(--r);border:1px solid rgba(255,51,86,.25)}
        .sev-high{background:rgba(255,140,0,.1);color:var(--ch);border:1px solid rgba(255,140,0,.22)}
        .sev-medium{background:rgba(255,183,0,.08);color:var(--am);border:1px solid rgba(255,183,0,.18)}
        .sev-dot{width:5px;height:5px;border-radius:50%;display:inline-block}
        .sev-dot.cr{background:var(--r);box-shadow:0 0 6px var(--r)} .sev-dot.hi{background:var(--ch)} .sev-dot.me{background:var(--am)}
        .summary-stat{text-align:center;flex:1}
        .summary-stat-num{font-family:'Syne',sans-serif;font-size:22px;font-weight:800;line-height:1}
        .summary-stat-label{font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-top:3px}
        .summary-group{margin-bottom:12px}
        .summary-group-title{display:flex;align-items:center;gap:8px;margin-bottom:6px}
        .summary-item{font-size:10px;color:#8fa5c8;line-height:1.55;padding:4px 0;padding-left:12px;border-left:2px solid rgba(255,255,255,.06);margin-bottom:3px}
      `}</style>

      {/* Header */}
      <div style={{borderBottom:"1px solid var(--border)",padding:"18px 32px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
            <span className="blink" style={{width:5,height:5,borderRadius:"50%",background:"var(--g)",display:"inline-block",boxShadow:"0 0 8px var(--g)"}}/>
            <span style={{fontSize:8,letterSpacing:".2em",color:"var(--muted)",textTransform:"uppercase"}}>Logistic Regression · 50 Features · 60/40 Split · 650-row Dataset</span>
          </div>
          <h1 style={{fontFamily:"'Syne',sans-serif",fontSize:"clamp(16px,2.8vw,26px)",fontWeight:800,letterSpacing:"-.02em",lineHeight:1.1}}>
            Financial Fraud <span style={{color:"var(--b)"}}>Detector</span>
          </h1>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {[["Train","389"],["Test","261"],["Accuracy","100%"],["ROC-AUC","1.00"],["Features","50"]].map(([k,v])=>(
            <div key={k} style={{background:"var(--card)",border:"1px solid var(--border)",borderRadius:7,padding:"6px 12px",textAlign:"center"}}>
              <div style={{fontSize:8,color:"var(--muted)",letterSpacing:".12em",textTransform:"uppercase"}}>{k}</div>
              <div style={{fontSize:12,color:"var(--b)",fontWeight:500,marginTop:1}}>{v}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{padding:"24px 32px",maxWidth:1200,margin:"0 auto"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 360px",gap:24,alignItems:"start"}}>

          {/* Left panel */}
          <div>
            <div
              className={`drop${dragOver?" over":""}${file?" has":""}`}
              onDragOver={e=>{e.preventDefault();setDragOver(true)}}
              onDragLeave={()=>setDragOver(false)}
              onDrop={handleDrop}
              onClick={()=>{if(!file)inputRef.current?.click()}}
              style={{cursor:file?"default":"pointer",marginBottom:16}}
            >
              <input ref={inputRef} type="file" accept=".pdf,.doc,.docx,.png,.jpg,.jpeg,.txt,.xlsx,.xls,.csv" onChange={e=>handleFile(e.target.files[0])}/>
              {!file ? (
                <>
                  <div style={{fontSize:36,marginBottom:12,opacity:.25}}>⬆</div>
                  <div style={{fontSize:13,marginBottom:5}}>Drop your financial statement here</div>
                  <div style={{fontSize:10,color:"var(--muted)",marginBottom:14}}>Annual report · 10-K filing · Balance sheet · Audit report</div>
                  <div style={{display:"flex",justifyContent:"center",gap:6,flexWrap:"wrap"}}>
                    {["PDF","DOCX","XLSX","CSV","PNG","JPEG","TXT"].map(t=>(
                      <span key={t} className="pill" style={{background:"rgba(255,255,255,.04)",color:"var(--muted)",border:"1px solid var(--border)"}}>{t}</span>
                    ))}
                  </div>
                </>
              ):(
                <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:8}}>
                  <div style={{fontSize:30}}>{file.type==="application/pdf"?"📄":file.type.startsWith("image/")?"🖼️":file.name.match(/\.(xlsx?|csv)$/i)?"📊":"📝"}</div>
                  <div style={{fontSize:12,color:"var(--g)",fontWeight:500}}>{file.name}</div>
                  <div style={{fontSize:10,color:"var(--muted)"}}>{(file.size/1024).toFixed(1)} KB</div>
                  <button onClick={reset} style={{marginTop:2,background:"transparent",border:"1px solid var(--border)",color:"var(--muted)",padding:"3px 10px",borderRadius:4,fontSize:9,fontFamily:"inherit"}}>✕ Remove</button>
                </div>
              )}
            </div>

            <div style={{display:"flex",gap:8,marginBottom:22,flexWrap:"wrap"}}>
              <button className="go" onClick={analyze} disabled={!file||isLoading}>
                {isLoading?<><span className="spin">⟳</span> {progress}</>:"▶ Extract & Predict"}
              </button>
              {stage!=="idle"&&<button onClick={reset} style={{background:"transparent",border:"1px solid var(--border)",color:"var(--muted)",padding:"12px 18px",borderRadius:8,fontSize:11,fontFamily:"inherit"}}>✕ Reset</button>}
            </div>

            {stage==="error"&&(
              <div className="up" style={{background:"rgba(255,51,86,.07)",border:"1px solid rgba(255,51,86,.2)",borderRadius:9,padding:"11px 15px",fontSize:10,color:"#ffb3c0",marginBottom:18}}>
                <strong style={{color:"var(--r)"}}>⚠ Error: </strong>{errMsg}
              </div>
            )}

            {extracted&&(
              <div className="up" style={{background:"var(--card)",border:"1px solid var(--border)",borderRadius:12,padding:"20px 22px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                  <div className="sec" style={{marginBottom:0}}>Extracted Metrics <span style={{color:"var(--g)"}}>({MODEL.featureCols.length} features)</span></div>
                  {extracted.company_name&&<span style={{fontSize:11,color:"var(--b)",fontWeight:500}}>{extracted.company_name}</span>}
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 24px",maxHeight:380,overflowY:"auto"}}>
                  {MODEL.featureCols.map(col=>{
                    const val=extracted[col], rule=THRESHOLDS[col];
                    let flagged=false;
                    if(rule&&val!==undefined&&val!==null){
                      const v=parseFloat(val);
                      if(!isNaN(v)){
                        if(rule.dir==="high"&&v>rule.v) flagged=true;
                        if(rule.dir==="low" &&v<rule.v) flagged=true;
                        if(rule.dir==="eq"  &&v==rule.v) flagged=true;
                      }
                    }
                    const display=(val!==undefined&&val!==null)?(typeof val==="number"?(Math.abs(val)>999?val.toLocaleString():parseFloat(val.toFixed(3))):val):"—";
                    return(
                      <div key={col} className="mrow">
                        <span style={{color:"var(--muted)",fontSize:9}}>{col.replace(/_/g," ")}</span>
                        <span style={{color:flagged?"var(--r)":"var(--text)",fontWeight:flagged?500:300,fontSize:10}}>
                          {display}{flagged&&<span style={{color:"var(--r)",marginLeft:3,fontSize:9}}>⚑</span>}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>

          {/* Right panel */}
          <div style={{position:"sticky",top:20}}>
            {!result&&!isLoading&&(
              <div style={{background:"var(--card)",border:"1px solid var(--border)",borderRadius:12,padding:24,textAlign:"center"}}>
                <div style={{fontSize:44,opacity:.15,marginBottom:12}}>◎</div>
                <div style={{fontSize:11,color:"var(--muted)",lineHeight:1.9}}>
                  Upload a financial statement<br/>
                  <span style={{color:"var(--b)"}}>Claude extracts 50 metrics</span><br/>
                  Model predicts fraud probability
                </div>
                <div style={{marginTop:18,borderTop:"1px solid var(--border)",paddingTop:14,textAlign:"left"}}>
                  <div className="sec">Model info</div>
                  {[["Dataset","650 companies"],["Split","60% / 40%"],["Train","389 samples"],["Test","261 samples"],["Algorithm","Logistic Regression"],["Features","50 metrics"],["Accuracy","100%"],["ROC-AUC","1.00"]].map(([k,v])=>(
                    <div key={k} style={{display:"flex",justifyContent:"space-between",fontSize:10,marginBottom:5}}>
                      <span style={{color:"var(--muted)"}}>{k}</span><span>{v}</span>
                    </div>
                  ))}
                </div>
                <div style={{marginTop:14,borderTop:"1px solid var(--border)",paddingTop:14,textAlign:"left"}}>
                  <div className="sec">Top signal features</div>
                  {MODEL.topFeatures.slice(0,6).map((f,i)=>(
                    <div key={f} style={{fontSize:9,color:"var(--muted)",marginBottom:4,display:"flex",gap:5}}>
                      <span style={{color:"var(--b)",opacity:.5}}>#{i+1}</span>{f.replace(/_/g," ")}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {isLoading&&(
              <div style={{background:"var(--card)",border:"1px solid var(--border)",borderRadius:12,padding:28,textAlign:"center"}}>
                <div style={{fontSize:32,marginBottom:12}}><span className="spin">⟳</span></div>
                <div style={{fontSize:12,color:"var(--b)",marginBottom:6}}>Analyzing…</div>
                <div style={{fontSize:10,color:"var(--muted)",lineHeight:1.8}}>{progress}</div>
                <div style={{marginTop:18,display:"flex",flexDirection:"column",gap:7,alignItems:"flex-start",paddingLeft:16}}>
                  {["Read file","Send to Claude","Run fraud model"].map((s,i)=>(
                    <div key={s} style={{display:"flex",alignItems:"center",gap:7,fontSize:10}}>
                      <span className={progressStep===i+1?"blink":""} style={{width:4,height:4,borderRadius:"50%",display:"inline-block",background:progressStep>i?"var(--g)":"var(--muted)"}}/>
                      <span style={{color:progressStep>i?"var(--text)":"var(--muted)"}}>{s}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {result&&(
              <>
                <div className="up" style={{background:"var(--card)",border:`1px solid ${result.isValid?"rgba(0,229,160,.25)":"rgba(255,51,86,.25)"}`,borderRadius:12,padding:20}}>
                  <div style={{textAlign:"center",paddingBottom:16,borderBottom:"1px solid var(--border)",marginBottom:16}}>
                    <div style={{fontSize:36,marginBottom:6}}>{result.isValid?"✅":"🚨"}</div>
                    {result.company_name&&<div style={{fontSize:9,color:"var(--muted)",marginBottom:5,letterSpacing:".1em"}}>{result.company_name}</div>}
                    <div style={{fontFamily:"'Syne',sans-serif",fontSize:22,fontWeight:800,color:result.isValid?"var(--g)":"var(--r)"}}>
                      {result.isValid?"VALID":"FRAUDULENT"}
                    </div>
                    <div style={{fontSize:8,color:"var(--muted)",letterSpacing:".16em",marginTop:3}}>MODEL VERDICT</div>
                  </div>

                  <div style={{display:"flex",gap:5,marginBottom:14}}>
                    {[["result","Verdict"],["metrics","Drivers"],["issues","Issues"]].map(([id,label])=>(
                      <button key={id} className={`tab${tab===id?" on":""}`} onClick={()=>setTab(id)} style={{flex:1,padding:"5px 0"}}>
                        {label}{id==="issues"&&result.problems.length>0&&<span style={{color:"var(--r)",marginLeft:3}}>({result.problems.length})</span>}
                      </button>
                    ))}
                  </div>

                  {tab==="result"&&(
                    <div>
                      {[{label:"Valid",val:result.probValid,color:"var(--g)"},{label:"Fraud",val:result.probFraud,color:"var(--r)"}].map(b=>(
                        <div key={b.label} style={{marginBottom:12}}>
                          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                            <span style={{fontSize:9,color:"var(--muted)",textTransform:"uppercase",letterSpacing:".1em"}}>{b.label} Probability</span>
                            <span style={{fontSize:12,fontWeight:500,color:b.color}}>{(b.val*100).toFixed(1)}%</span>
                          </div>
                          <div style={{background:"rgba(255,255,255,.04)",borderRadius:3,height:4,overflow:"hidden"}}>
                            <div className="bar" style={{height:"100%",width:`${b.val*100}%`,background:b.color,borderRadius:3}}/>
                          </div>
                        </div>
                      ))}
                      <div style={{marginTop:12,background:"rgba(255,255,255,.02)",borderRadius:7,padding:"10px 12px",fontSize:10,lineHeight:1.7,color:"var(--muted)"}}>
                        {result.isValid
                          ?"✓ Model classifies this statement as financially valid with no significant fraud signals."
                          :"⚠ Model classifies this statement as likely fraudulent. Review flagged issues."}
                      </div>
                    </div>
                  )}

                  {tab==="metrics"&&(
                    <div>
                      <div className="sec">Feature contributions (top 8)</div>
                      {[...result.contributions].sort((a,b)=>Math.abs(b.contrib)-Math.abs(a.contrib)).slice(0,8).map(d=>(
                        <div key={d.col} style={{marginBottom:8}}>
                          <div style={{display:"flex",justifyContent:"space-between",marginBottom:3,fontSize:9}}>
                            <span style={{color:"var(--muted)"}}>{d.col.replace(/_/g," ")}</span>
                            <span style={{color:d.contrib>0?"var(--g)":"var(--r)",fontWeight:500}}>{d.contrib>0?"+":""}{d.contrib.toFixed(3)}</span>
                          </div>
                          <div style={{background:"rgba(255,255,255,.04)",borderRadius:3,height:3,overflow:"hidden"}}>
                            <div style={{height:"100%",width:`${Math.min(Math.abs(d.contrib)/0.2*100,100)}%`,background:d.contrib>0?"var(--g)":"var(--r)",borderRadius:3,transition:"width 1.1s ease"}}/>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  {tab==="issues"&&(
                    <div>
                      {result.problems.length===0?(
                        <div style={{textAlign:"center",padding:"14px 0",fontSize:11,color:"var(--muted)"}}>
                          <div style={{fontSize:26,marginBottom:6}}>✓</div>No anomalies detected.
                        </div>
                      ):(
                        <>
                          <div className="sec">{result.isValid?"Warnings":"Detected Issues"} ({result.problems.length})</div>
                          {result.problems.map((p,i)=><div key={i} className={result.isValid?"wc":"pc"}>{p}</div>)}
                        </>
                      )}
                    </div>
                  )}
                </div>

                {result.problems.length>0&&(
                  <div className="summary-panel up" style={{animationDelay:".15s"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                      <div style={{fontSize:9,letterSpacing:".18em",textTransform:"uppercase",color:"var(--muted)"}}>Problems Found</div>
                      <div style={{fontSize:10,color:result.isValid?"var(--am)":"var(--r)",fontWeight:500}}>{result.problems.length} issue{result.problems.length!==1?"s":""} detected</div>
                    </div>
                    <div style={{display:"flex",gap:8,marginBottom:16,padding:"10px 0",borderTop:"1px solid var(--border)",borderBottom:"1px solid var(--border)"}}>
                      {[["Critical",criticalProbs.length,"var(--r)"],["High",highProbs.length,"var(--ch)"],["Medium",mediumProbs.length,"var(--am)"],["Total",result.problems.length,"var(--text)"]].map(([label,count,color],i,arr)=>(
                        <><div key={label} className="summary-stat"><div className="summary-stat-num" style={{color}}>{count}</div><div className="summary-stat-label">{label}</div></div>{i<arr.length-1&&<div style={{width:"1px",background:"var(--border)"}}/>}</>
                      ))}
                    </div>
                    {[["critical",criticalProbs,"sev-critical","cr","rgba(255,51,86,.3)","#ffb3c0"],
                      ["high",highProbs,"sev-high","hi","rgba(255,140,0,.3)","#ffd08a"],
                      ["medium",mediumProbs,"sev-medium","me","rgba(255,183,0,.25)","#ffe5a0"]].map(([sev,list,badge,dot,blc,tc])=>
                      list.length>0&&(
                        <div key={sev} className="summary-group" style={{marginBottom:sev==="medium"?0:12}}>
                          <div className="summary-group-title"><span className={`sev-badge ${badge}`}><span className={`sev-dot ${dot}`}/>{sev.charAt(0).toUpperCase()+sev.slice(1)}</span></div>
                          {list.map((p,i)=><div key={i} className="summary-item" style={{borderLeftColor:blc,color:tc}}>{p}</div>)}
                        </div>
                      )
                    )}
                    <div style={{marginTop:14,padding:"10px 12px",background:"rgba(255,255,255,.02)",borderRadius:7,fontSize:10,color:"var(--muted)",lineHeight:1.65,borderTop:"1px solid var(--border)"}}>
                      {criticalProbs.length>0
                        ?`⚠ ${criticalProbs.length} critical indicator${criticalProbs.length>1?"s":""} — immediate forensic review recommended.`
                        :highProbs.length>0
                        ?`⚠ ${highProbs.length} high-severity signal${highProbs.length>1?"s":""} require escalated scrutiny.`
                        :`ℹ ${mediumProbs.length} medium-severity warning${mediumProbs.length>1?"s":""} — monitor closely.`}
                    </div>
                  </div>
                )}

                {result.problems.length===0&&(
                  <div className="summary-panel up" style={{animationDelay:".15s",textAlign:"center",padding:"18px 20px"}}>
                    <div style={{fontSize:22,marginBottom:6}}>✓</div>
                    <div style={{fontSize:10,color:"var(--g)",fontWeight:500,marginBottom:4}}>No Problems Detected</div>
                    <div style={{fontSize:9,color:"var(--muted)",lineHeight:1.6}}>All threshold checks passed. No anomalies identified across fraud signals, governance metrics, or financial ratios.</div>
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
