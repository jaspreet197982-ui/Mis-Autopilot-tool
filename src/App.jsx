import { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ── helpers ──────────────────────────────────────────────────────────────────
const fmt = (n) =>
  typeof n === "number" && !isNaN(n)
    ? n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : "—";

const pctColor = (val, isRevenue = true) => {
  const fav = isRevenue ? val >= 0 : val <= 0;
  return fav ? "#16a34a" : "#dc2626";
};

const pctBg = (val, isRevenue = true) => {
  const fav = isRevenue ? val >= 0 : val <= 0;
  return fav ? "#dcfce7" : "#fee2e2";
};

function calcPct(actual, budget) {
  if (!budget || budget === 0) return 0;
  return ((actual - budget) / Math.abs(budget)) * 100;
}

// ── parse uploaded Excel ──────────────────────────────────────────────────────
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ── detect columns automatically ─────────────────────────────────────────────
function detectColumns(rows) {
  if (!rows.length) return {};
  const keys = Object.keys(rows[0]);
  const find = (...hints) =>
    keys.find((k) => hints.some((h) => k.toLowerCase().includes(h.toLowerCase()))) || null;
  return {
    label: find("narration", "particular", "account", "head", "description", "item", "category"),
    actual: find("actual", "actuals", "current", "achieved", "real"),
    budget: find("budget", "budgeted", "target", "plan", "planned"),
    lastYear: find("last year", "ly", "previous", "prior", "py", "last"),
  };
}

// ── main component ────────────────────────────────────────────────────────────
export default function MISAutopilot() {
  const [stage, setStage] = useState("landing"); // landing | setup | generating | report
  const [files, setFiles] = useState({ actuals: null, budget: null });
  const [fileNames, setFileNames] = useState({ actuals: "", budget: "" });
  const [companyName, setCompanyName] = useState("");
  const [industry, setIndustry] = useState("");
  const [period, setPeriod] = useState("");
  const [currency, setCurrency] = useState("₹ Cr");
  const [rows, setRows] = useState([]);
  const [colMap, setColMap] = useState({});
  const [progress, setProgress] = useState(0);
  const [progressMsg, setProgressMsg] = useState("");
  const [commentary, setCommentary] = useState("");
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(null);
  const actualsRef = useRef();
  const budgetRef = useRef();

  // ── file handlers ────────────────────────────────────────────────────────
  const handleFile = async (type, file) => {
    if (!file) return;
    setFileNames((p) => ({ ...p, [type]: file.name }));
    setFiles((p) => ({ ...p, [type]: file }));
  };

  // ── generate report ──────────────────────────────────────────────────────
  const generate = async () => {
    setError("");
    if (!files.actuals) { setError("Please upload your Actuals file."); return; }
    if (!companyName.trim()) { setError("Please enter your company name."); return; }
    if (!period.trim()) { setError("Please enter the report period."); return; }

    setStage("generating");
    setProgress(10);
    setProgressMsg("Reading your data files...");

    try {
      // Parse actuals
      const actualsRows = await parseExcel(files.actuals);
      let budgetRows = [];
      if (files.budget) budgetRows = await parseExcel(files.budget);

      setProgress(30);
      setProgressMsg("Detecting columns and mapping data...");

      // Detect columns
      const aCols = detectColumns(actualsRows);
      const bCols = detectColumns(budgetRows.length ? budgetRows : actualsRows);

      // Build unified rows
      let unified = [];

      if (files.budget && budgetRows.length) {
        // Merge by label
        const budgetMap = {};
        budgetRows.forEach((r) => {
          const lbl = r[bCols.label] || r[Object.keys(r)[0]];
          const val = parseFloat(r[bCols.actual] || r[bCols.budget] || 0) || 0;
          if (lbl) budgetMap[String(lbl).trim()] = val;
        });

        unified = actualsRows
          .filter((r) => r[aCols.label])
          .map((r) => {
            const label = String(r[aCols.label]).trim();
            const actual = parseFloat(r[aCols.actual] || 0) || 0;
            const ly = parseFloat(r[aCols.lastYear] || 0) || 0;
            const budget = budgetMap[label] ?? parseFloat(r[aCols.budget] || 0) || 0;
            return { label, actual, budget, ly };
          })
          .filter((r) => r.label && (r.actual !== 0 || r.budget !== 0));
      } else {
        // Single file — use budget column if present
        unified = actualsRows
          .filter((r) => r[aCols.label])
          .map((r) => {
            const label = String(r[aCols.label]).trim();
            const actual = parseFloat(r[aCols.actual] || 0) || 0;
            const budget = parseFloat(r[aCols.budget] || 0) || 0;
            const ly = parseFloat(r[aCols.lastYear] || 0) || 0;
            return { label, actual, budget, ly };
          })
          .filter((r) => r.label && r.actual !== 0);
      }

      setRows(unified);
      setColMap(aCols);
      setProgress(55);
      setProgressMsg("Sending data to AI for analysis...");

      // Build prompt
      const tableText = unified
        .slice(0, 20)
        .map((r) => {
          const v = r.budget ? calcPct(r.actual, r.budget).toFixed(1) : "N/A";
          return `${r.label}: Actual ${currency} ${r.actual.toFixed(2)} | Budget ${currency} ${r.budget.toFixed(2)} | Variance ${v}%`;
        })
        .join("\n");

      const prompt = `You are a Senior Finance Manager. Write a professional MIS management commentary for the CFO and Board.

COMPANY: ${companyName}
INDUSTRY: ${industry || "Corporate"}
PERIOD: ${period}
CURRENCY: ${currency}

FINANCIAL DATA:
${tableText}

Write the following in professional corporate finance language:

1. EXECUTIVE SUMMARY (3 bullet points — key highlights for CFO in 60 seconds)
2. REVENUE / TOP-LINE COMMENTARY (2-3 sentences on performance vs budget)
3. COST COMMENTARY (2-3 sentences on major cost variances and drivers)
4. PROFITABILITY COMMENTARY (1-2 sentences on overall margin and bottom line)
5. TOP 3 VARIANCE DRIVERS (specific items with numbers and reasons)
6. RECOMMENDED ACTIONS (3 specific actions management should take)

TONE: Professional, direct, factual. Suitable for board presentation. Use specific numbers from the data provided.`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          messages: [{ role: "user", content: prompt }],
        }),
      });

      setProgress(85);
      setProgressMsg("Formatting your report...");

      const data = await response.json();
      const text = data.content?.map((c) => c.text || "").join("\n") || "";
      setCommentary(text || "Commentary could not be generated. Please check your API connection.");
      setProgress(100);
      setTimeout(() => setStage("report"), 500);
    } catch (err) {
      setError("Something went wrong: " + err.message);
      setStage("setup");
    }
  };

  const reset = () => {
    setStage("landing");
    setFiles({ actuals: null, budget: null });
    setFileNames({ actuals: "", budget: "" });
    setCompanyName(""); setIndustry(""); setPeriod(""); setRows([]);
    setCommentary(""); setError(""); setProgress(0);
  };

  // ── render commentary ─────────────────────────────────────────────────────
  const renderCommentary = (text) =>
    text.split("\n").map((line, i) => {
      if (!line.trim()) return <div key={i} style={{ height: 8 }} />;
      if (line.match(/^#+\s/) || (line.match(/^[A-Z\s\/&]+:$/) || (line.endsWith(":") && line.length < 60 && line === line.toUpperCase()))) {
        return <div key={i} style={{ fontWeight: 700, color: "#0f4c81", fontSize: 13, marginTop: 14, marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>{line.replace(/^#+\s/, "")}</div>;
      }
      if (line.match(/^\d\.|^•|^-\s/)) {
        return <div key={i} style={{ display: "flex", gap: 8, marginBottom: 4 }}>
          <span style={{ color: "#0f4c81", fontWeight: 700, minWidth: 16 }}>•</span>
          <span style={{ color: "#334155", fontSize: 13, lineHeight: 1.6 }}>{line.replace(/^\d\.|^•|^-\s/, "").trim()}</span>
        </div>;
      }
      return <p key={i} style={{ color: "#334155", fontSize: 13, lineHeight: 1.7, margin: "3px 0" }}>{line}</p>;
    });

  // ── KPI summary from rows ─────────────────────────────────────────────────
  const topRows = rows.slice(0, 6);
  const totalActual = rows.reduce((s, r) => s + r.actual, 0);
  const totalBudget = rows.reduce((s, r) => s + r.budget, 0);
  const overBudgetCount = rows.filter((r) => r.budget && r.actual > r.budget).length;
  const underBudgetCount = rows.filter((r) => r.budget && r.actual < r.budget).length;

  // ─────────────────────────────────────────────────────────────────────────
  // RENDER
  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div style={{ fontFamily: "'Georgia', 'Times New Roman', serif", background: "#f0f4f8", minHeight: "100vh" }}>

      {/* ── HEADER ── */}
      <div style={{ background: "linear-gradient(135deg, #0a3d6b 0%, #1260a8 100%)", color: "white", padding: "20px 28px", display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ background: "rgba(255,255,255,0.15)", borderRadius: 10, padding: "8px 12px", fontSize: 20 }}>📊</div>
        <div>
          <div style={{ fontSize: 20, fontWeight: 700, letterSpacing: 1 }}>MIS AUTOPILOT</div>
          <div style={{ fontSize: 11, opacity: 0.75, letterSpacing: 2 }}>AI-POWERED FINANCIAL REPORTING • ANY COMPANY • ANY INDUSTRY</div>
        </div>
        {stage === "report" && (
          <button onClick={reset} style={{ marginLeft: "auto", background: "rgba(255,255,255,0.15)", border: "1px solid rgba(255,255,255,0.3)", color: "white", borderRadius: 8, padding: "7px 16px", cursor: "pointer", fontSize: 13 }}>
            ← New Report
          </button>
        )}
      </div>

      <div style={{ maxWidth: 900, margin: "0 auto", padding: "24px 16px" }}>

        {/* ══ LANDING ══════════════════════════════════════════════════════ */}
        {stage === "landing" && (
          <div>
            {/* Hero */}
            <div style={{ background: "white", borderRadius: 16, padding: "40px 36px", textAlign: "center", boxShadow: "0 4px 24px rgba(10,61,107,0.10)", marginBottom: 20 }}>
              <div style={{ fontSize: 52, marginBottom: 12 }}>⚡</div>
              <h1 style={{ color: "#0a3d6b", fontSize: 26, fontWeight: 700, margin: "0 0 10px" }}>
                6 Hours of MIS Work → 2 Minutes
              </h1>
              <p style={{ color: "#64748b", fontSize: 15, maxWidth: 520, margin: "0 auto 28px", lineHeight: 1.6 }}>
                Upload your financial data. Get a CFO-ready MIS report with variance analysis and AI management commentary — instantly.
              </p>
              <button onClick={() => setStage("setup")}
                style={{ background: "linear-gradient(135deg, #0a3d6b, #1260a8)", color: "white", border: "none", borderRadius: 12, padding: "15px 44px", fontSize: 16, fontWeight: 700, cursor: "pointer", boxShadow: "0 4px 18px rgba(10,61,107,0.3)", letterSpacing: 0.5 }}>
                Get Started →
              </button>
            </div>

            {/* Feature cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14 }}>
              {[
                { icon: "📁", title: "Upload Any Format", desc: "Excel, CSV from Tally, SAP, ERP or any system" },
                { icon: "🤖", title: "AI Commentary", desc: "Professional management narrative written in seconds" },
                { icon: "📄", title: "CFO-Ready Report", desc: "Variance analysis, KPIs, cost centres — all formatted" },
              ].map((f, i) => (
                <div key={i} style={{ background: "white", borderRadius: 12, padding: "22px 18px", boxShadow: "0 2px 10px rgba(0,0,0,0.06)", textAlign: "center" }}>
                  <div style={{ fontSize: 32, marginBottom: 10 }}>{f.icon}</div>
                  <div style={{ fontWeight: 700, color: "#0a3d6b", marginBottom: 6, fontSize: 14 }}>{f.title}</div>
                  <div style={{ color: "#64748b", fontSize: 13, lineHeight: 1.5 }}>{f.desc}</div>
                </div>
              ))}
            </div>

            {/* Works for */}
            <div style={{ background: "white", borderRadius: 12, padding: "18px 24px", marginTop: 14, boxShadow: "0 2px 10px rgba(0,0,0,0.06)" }}>
              <div style={{ color: "#94a3b8", fontSize: 11, textTransform: "uppercase", letterSpacing: 1, marginBottom: 10 }}>Works for any industry</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {["Manufacturing", "FMCG & Beverages", "IT Services", "Pharma", "Infrastructure", "Retail", "Real Estate", "Automotive", "Chemicals", "Logistics"].map((ind) => (
                  <span key={ind} style={{ background: "#eff6ff", color: "#1e40af", borderRadius: 20, padding: "4px 12px", fontSize: 12 }}>{ind}</span>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ══ SETUP ════════════════════════════════════════════════════════ */}
        {stage === "setup" && (
          <div style={{ background: "white", borderRadius: 16, padding: "32px 28px", boxShadow: "0 4px 24px rgba(10,61,107,0.10)" }}>
            <h2 style={{ color: "#0a3d6b", fontSize: 20, marginBottom: 4 }}>Set Up Your MIS Report</h2>
            <p style={{ color: "#64748b", fontSize: 13, marginBottom: 28 }}>Upload your data and fill in your company details below.</p>

            {error && (
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 8, padding: "10px 16px", marginBottom: 20, color: "#dc2626", fontSize: 13 }}>
                ⚠️ {error}
              </div>
            )}

            {/* Company details */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
              <div>
                <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>Company Name *</label>
                <input value={companyName} onChange={e => setCompanyName(e.target.value)}
                  placeholder="e.g. Varun Beverages Ltd"
                  style={{ width: "100%", padding: "10px 12px", border: "1.5px solid #e2e8f0", borderRadius: 8, fontSize: 14, color: "#1e293b", outline: "none", boxSizing: "border-box" }} />
              </div>
              <div>
                <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>Industry</label>
                <input value={industry} onChange={e => setIndustry(e.target.value)}
                  placeholder="e.g. FMCG Beverages"
                  style={{ width: "100%", padding: "10px 12px", border: "1.5px solid #e2e8f0", borderRadius: 8, fontSize: 14, color: "#1e293b", outline: "none", boxSizing: "border-box" }} />
              </div>
              <div>
                <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>Report Period *</label>
                <input value={period} onChange={e => setPeriod(e.target.value)}
                  placeholder="e.g. April 2026"
                  style={{ width: "100%", padding: "10px 12px", border: "1.5px solid #e2e8f0", borderRadius: 8, fontSize: 14, color: "#1e293b", outline: "none", boxSizing: "border-box" }} />
              </div>
              <div>
                <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>Currency / Unit</label>
                <select value={currency} onChange={e => setCurrency(e.target.value)}
                  style={{ width: "100%", padding: "10px 12px", border: "1.5px solid #e2e8f0", borderRadius: 8, fontSize: 14, color: "#1e293b", outline: "none", background: "white", boxSizing: "border-box" }}>
                  <option>₹ Cr</option>
                  <option>₹ Lakhs</option>
                  <option>₹</option>
                  <option>USD Mn</option>
                  <option>USD</option>
                </select>
              </div>
            </div>

            {/* File uploads */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 28 }}>
              {/* Actuals */}
              <div>
                <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>
                  Actuals File * <span style={{ color: "#94a3b8", fontWeight: 400 }}>(Excel / CSV)</span>
                </label>
                <div
                  onDragOver={e => { e.preventDefault(); setDragging("actuals"); }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={e => { e.preventDefault(); setDragging(null); handleFile("actuals", e.dataTransfer.files[0]); }}
                  onClick={() => actualsRef.current.click()}
                  style={{ border: `2px dashed ${dragging === "actuals" ? "#1260a8" : fileNames.actuals ? "#22c55e" : "#cbd5e1"}`, borderRadius: 10, padding: "20px 16px", textAlign: "center", cursor: "pointer", background: fileNames.actuals ? "#f0fdf4" : dragging === "actuals" ? "#eff6ff" : "#f8fafc", transition: "all 0.2s" }}>
                  <div style={{ fontSize: 28, marginBottom: 6 }}>{fileNames.actuals ? "✅" : "📂"}</div>
                  <div style={{ fontSize: 13, color: fileNames.actuals ? "#16a34a" : "#64748b", fontWeight: fileNames.actuals ? 600 : 400 }}>
                    {fileNames.actuals || "Drag & drop or click to upload"}
                  </div>
                  {!fileNames.actuals && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>Tally export, SAP, ERP, any Excel</div>}
                </div>
                <input ref={actualsRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => handleFile("actuals", e.target.files[0])} />
              </div>

              {/* Budget */}
              <div>
                <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>
                  Budget File <span style={{ color: "#94a3b8", fontWeight: 400 }}>(Optional)</span>
                </label>
                <div
                  onDragOver={e => { e.preventDefault(); setDragging("budget"); }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={e => { e.preventDefault(); setDragging(null); handleFile("budget", e.dataTransfer.files[0]); }}
                  onClick={() => budgetRef.current.click()}
                  style={{ border: `2px dashed ${dragging === "budget" ? "#1260a8" : fileNames.budget ? "#22c55e" : "#cbd5e1"}`, borderRadius: 10, padding: "20px 16px", textAlign: "center", cursor: "pointer", background: fileNames.budget ? "#f0fdf4" : dragging === "budget" ? "#eff6ff" : "#f8fafc", transition: "all 0.2s" }}>
                  <div style={{ fontSize: 28, marginBottom: 6 }}>{fileNames.budget ? "✅" : "📋"}</div>
                  <div style={{ fontSize: 13, color: fileNames.budget ? "#16a34a" : "#64748b", fontWeight: fileNames.budget ? 600 : 400 }}>
                    {fileNames.budget || "Drag & drop or click to upload"}
                  </div>
                  {!fileNames.budget && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>Skip if budget is in actuals file</div>}
                </div>
                <input ref={budgetRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => handleFile("budget", e.target.files[0])} />
              </div>
            </div>

            {/* File format tip */}
            <div style={{ background: "#fefce8", border: "1px solid #fde047", borderRadius: 8, padding: "12px 16px", marginBottom: 24, fontSize: 12, color: "#713f12" }}>
              💡 <strong>File format tip:</strong> Your Excel should have columns like: <em>Particulars / Account Head</em> (row names), <em>Actual</em>, <em>Budget</em>. The app auto-detects column names — no reformatting needed.
            </div>

            <div style={{ display: "flex", gap: 12 }}>
              <button onClick={() => setStage("landing")}
                style={{ padding: "12px 24px", background: "white", border: "1.5px solid #e2e8f0", borderRadius: 10, cursor: "pointer", fontSize: 14, color: "#475569" }}>
                ← Back
              </button>
              <button onClick={generate}
                style={{ flex: 1, background: "linear-gradient(135deg, #0a3d6b, #1260a8)", color: "white", border: "none", borderRadius: 10, padding: "13px 24px", fontSize: 15, fontWeight: 700, cursor: "pointer", boxShadow: "0 4px 14px rgba(10,61,107,0.25)" }}>
                🚀 Generate MIS Report
              </button>
            </div>
          </div>
        )}

        {/* ══ GENERATING ═══════════════════════════════════════════════════ */}
        {stage === "generating" && (
          <div style={{ background: "white", borderRadius: 16, padding: "52px 40px", textAlign: "center", boxShadow: "0 4px 24px rgba(10,61,107,0.10)" }}>
            <div style={{ fontSize: 52, marginBottom: 16 }}>⚙️</div>
            <h2 style={{ color: "#0a3d6b", fontSize: 20, marginBottom: 8 }}>Generating Your MIS Report</h2>
            <p style={{ color: "#64748b", fontSize: 14, marginBottom: 32 }}>{progressMsg}</p>
            <div style={{ background: "#f0f4f8", borderRadius: 20, height: 10, marginBottom: 14, overflow: "hidden", maxWidth: 400, margin: "0 auto 14px" }}>
              <div style={{ height: "100%", width: `${progress}%`, background: "linear-gradient(90deg, #0a3d6b, #1260a8)", borderRadius: 20, transition: "width 0.6s ease" }} />
            </div>
            <div style={{ color: "#0a3d6b", fontWeight: 600, fontSize: 13 }}>{progress}% complete</div>
          </div>
        )}

        {/* ══ REPORT ═══════════════════════════════════════════════════════ */}
        {stage === "report" && (
          <>
            {/* Report header */}
            <div style={{ background: "white", borderRadius: 12, padding: "20px 24px", boxShadow: "0 2px 10px rgba(0,0,0,0.06)", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, color: "#0a3d6b" }}>{companyName}</div>
                <div style={{ fontSize: 13, color: "#64748b" }}>MIS Report — {period} {industry ? `| ${industry}` : ""} | All figures in {currency}</div>
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={reset} style={{ padding: "8px 16px", background: "white", border: "1px solid #e2e8f0", borderRadius: 8, cursor: "pointer", fontSize: 13, color: "#475569" }}>
                  ← New Report
                </button>
                <button onClick={() => window.print()} style={{ padding: "8px 16px", background: "#0a3d6b", color: "white", border: "none", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
                  ⬇ Download PDF
                </button>
              </div>
            </div>

            {/* Summary KPI strip */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12, marginBottom: 16 }}>
              {[
                { label: "Total Actuals", value: fmt(totalActual), sub: `Budget: ${fmt(totalBudget)}`, icon: "💰", color: "#0a3d6b" },
                { label: "Above Budget Items", value: overBudgetCount, sub: "line items over budget", icon: "🔴", color: "#dc2626" },
                { label: "Within Budget Items", value: underBudgetCount, sub: "line items on track", icon: "🟢", color: "#16a34a" },
              ].map((k, i) => (
                <div key={i} style={{ background: "white", borderRadius: 12, padding: "16px 18px", boxShadow: "0 2px 10px rgba(0,0,0,0.06)", borderLeft: `4px solid ${k.color}` }}>
                  <div style={{ fontSize: 11, color: "#94a3b8", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>{k.label}</div>
                  <div style={{ fontSize: 22, fontWeight: 700, color: k.color }}>{k.icon} {k.value}</div>
                  <div style={{ fontSize: 12, color: "#64748b" }}>{k.sub}</div>
                </div>
              ))}
            </div>

            {/* P&L Table */}
            <div style={{ background: "white", borderRadius: 12, boxShadow: "0 2px 10px rgba(0,0,0,0.06)", marginBottom: 16, overflow: "hidden" }}>
              <div style={{ background: "#0a3d6b", color: "white", padding: "12px 20px", fontWeight: 700, fontSize: 14 }}>
                📋 Variance Analysis — {period}
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: "#f8fafc" }}>
                      {["Particulars", "Actual", "Budget", "Variance", "Var %", ...(rows[0]?.ly ? ["Last Year", "YoY %"] : [])].map((h) => (
                        <th key={h} style={{ padding: "10px 14px", textAlign: h === "Particulars" ? "left" : "right", color: "#475569", fontWeight: 600, borderBottom: "2px solid #e2e8f0", whiteSpace: "nowrap", fontSize: 12 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row, i) => {
                      const pct = row.budget ? calcPct(row.actual, row.budget) : null;
                      const yoy = row.ly ? calcPct(row.actual, row.ly) : null;
                      return (
                        <tr key={i} style={{ background: i % 2 === 0 ? "white" : "#fafafa", borderBottom: "1px solid #f1f5f9" }}>
                          <td style={{ padding: "9px 14px", color: "#1e293b" }}>{row.label}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontWeight: 600, color: "#1e293b" }}>{fmt(row.actual)}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", color: "#64748b" }}>{row.budget ? fmt(row.budget) : "—"}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", color: pct !== null ? (pct >= 0 ? "#16a34a" : "#dc2626") : "#94a3b8", fontWeight: 600 }}>
                            {pct !== null ? `${pct >= 0 ? "+" : ""}${fmt(row.actual - row.budget)}` : "—"}
                          </td>
                          <td style={{ padding: "9px 14px", textAlign: "right" }}>
                            {pct !== null ? (
                              <span style={{ background: pctBg(pct), color: pctColor(pct), borderRadius: 6, padding: "2px 8px", fontSize: 11, fontWeight: 700 }}>
                                {pct >= 0 ? "▲" : "▼"} {Math.abs(pct).toFixed(1)}%
                              </span>
                            ) : "—"}
                          </td>
                          {rows[0]?.ly && <>
                            <td style={{ padding: "9px 14px", textAlign: "right", color: "#64748b" }}>{row.ly ? fmt(row.ly) : "—"}</td>
                            <td style={{ padding: "9px 14px", textAlign: "right" }}>
                              {yoy !== null ? (
                                <span style={{ color: yoy >= 0 ? "#16a34a" : "#dc2626", fontSize: 12, fontWeight: 600 }}>
                                  {yoy >= 0 ? "▲" : "▼"} {Math.abs(yoy).toFixed(1)}%
                                </span>
                              ) : "—"}
                            </td>
                          </>}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* AI Commentary */}
            <div style={{ background: "white", borderRadius: 12, boxShadow: "0 2px 10px rgba(0,0,0,0.06)", overflow: "hidden", marginBottom: 16 }}>
              <div style={{ background: "linear-gradient(90deg, #0a3d6b, #1260a8)", color: "white", padding: "12px 20px", fontWeight: 700, fontSize: 14, display: "flex", alignItems: "center", gap: 8 }}>
                <span>🤖</span> AI Management Commentary
                <span style={{ marginLeft: "auto", background: "rgba(255,255,255,0.18)", borderRadius: 20, padding: "2px 12px", fontSize: 11 }}>Generated by Claude AI</span>
              </div>
              <div style={{ padding: "22px 24px", lineHeight: 1.7 }}>
                {renderCommentary(commentary)}
              </div>
            </div>

            {/* Footer */}
            <div style={{ textAlign: "center", color: "#94a3b8", fontSize: 12, paddingBottom: 16 }}>
              Generated by MIS Autopilot • {new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" })} • Powered by Claude AI
            </div>
          </>
        )}
      </div>
    </div>
  );
} 
