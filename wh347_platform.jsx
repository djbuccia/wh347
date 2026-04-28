import { useState, useRef, useCallback } from "react";

const ANTHROPIC_MODEL = "claude-sonnet-4-20250514";

// ── Embedded payroll report template (xlsx, base64) ──────────────────────────
const TEMPLATE_B64 = "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";

function downloadTemplate() {
  const byteChars = atob(TEMPLATE_B64);
  const byteNums = new Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) byteNums[i] = byteChars.charCodeAt(i);
  const blob = new Blob([new Uint8Array(byteNums)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "WH347_Payroll_Report_Template.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

// ── Parsing prompt ──────────────────────────────────────────────────────────
const PARSE_SYSTEM = `You are a certified payroll specialist. Extract payroll data from uploaded reports and return ONLY valid JSON (no markdown, no explanation).

Return this exact structure:
{
  "contractor_name": "string",
  "contractor_address": "string",
  "project_name": "string",
  "project_location": "string",
  "project_number": "string",
  "contract_number": "string",
  "payroll_number": "string",
  "week_ending": "MM/DD/YYYY",
  "period_start": "MM/DD/YYYY",
  "signatory_name": "string",
  "signatory_title": "string",
  "signature_date": "MM/DD/YYYY",
  "employees": [
    {
      "name": "Last, First",
      "ss_last4": "last 4 digits only (use **** if unavailable)",
      "classification": "trade classification e.g. Carpenter, Electrician Journeyman",
      "daily_hours": [Sun, Mon, Tue, Wed, Thu, Fri, Sat],
      "st_hours": number,
      "ot_hours": number,
      "rate_of_pay": "hourly rate as string e.g. 45.50",
      "ot_rate": "OT hourly rate as string",
      "gross_wages": "total gross as string e.g. 1820.00",
      "fica": "FICA deduction",
      "fed_tax": "federal income tax",
      "state_tax": "state income tax",
      "other_ded1": "other deduction if any",
      "other_ded2": "",
      "other_ded3": "",
      "total_deductions": "sum of all deductions",
      "net_wages": "gross minus total deductions"
    }
  ]
}

If a field is missing, use empty string "". For daily_hours, use 0 for days not worked.`;

// ── WH-347 PDF Generator (pure JS / pdf-lib) ─────────────────────────────────
// We embed a PDF generation function using jsPDF since we're browser-only
async function generateWH347PDF(data) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "portrait", unit: "pt", format: "letter" });

  const W = 612, H = 792;
  const margin = 30;
  const top = H - 30;

  const drawRect = (x, y, w, h) => {
    doc.rect(x, y, w, h);
  };

  const setFont = (size, bold = false) => {
    doc.setFont("helvetica", bold ? "bold" : "normal");
    doc.setFontSize(size);
  };

  const text = (str, x, y, opts = {}) => {
    if (!str) return;
    const s = String(str);
    if (opts.align === "center") {
      doc.text(s, x, y, { align: "center" });
    } else {
      doc.text(s, x, y);
    }
  };

  doc.setTextColor(0, 0, 0);

  // ── HEADER
  setFont(9, true);
  text("U.S. DEPARTMENT OF LABOR", W / 2, top, { align: "center" });
  setFont(7, true);
  text("WAGE AND HOUR DIVISION", W / 2, top - 10, { align: "center" });
  setFont(11, true);
  text("PAYROLL", W / 2, top - 22, { align: "center" });
  setFont(7);
  text("(For Contractor's Optional Use; See Instructions at www.dol.gov/whd)", W / 2, top - 32, { align: "center" });
  setFont(7);
  doc.text("Form WH-347 (Rev. 08/2008)", W - margin, top, { align: "right" });
  doc.text("OMB No. 1235-0008  Expires: 08/31/2026", W - margin, top - 9, { align: "right" });

  // ── ROW 1 BOXES
  const y1 = H - 105;
  const row1h = 45;

  drawRect(margin, y1, 240, row1h);
  setFont(5.5);
  doc.setTextColor(80, 80, 80);
  text("NAME AND INDIVIDUAL TITLE OF SIGNATORY PARTY", margin + 2, y1 + row1h - 7);
  setFont(7.5);
  doc.setTextColor(0, 0, 0);
  text((data.contractor_name || "").substring(0, 44), margin + 4, y1 + row1h - 18);
  text((data.contractor_address || "").substring(0, 44), margin + 4, y1 + row1h - 28);

  drawRect(margin + 240, y1, 200, row1h);
  setFont(5.5);
  doc.setTextColor(80, 80, 80);
  text("PROJECT AND LOCATION", margin + 242, y1 + row1h - 7);
  setFont(7.5);
  doc.setTextColor(0, 0, 0);
  text((data.project_name || "").substring(0, 36), margin + 244, y1 + row1h - 18);
  text((data.project_location || "").substring(0, 36), margin + 244, y1 + row1h - 28);

  drawRect(margin + 440, y1, 80, row1h / 2);
  setFont(5.5);
  doc.setTextColor(80, 80, 80);
  text("PAYROLL NO.", margin + 442, y1 + row1h / 2 - 5);
  setFont(8, true);
  doc.setTextColor(0, 0, 0);
  text(String(data.payroll_number || ""), margin + 444, y1 + row1h / 2 - 15);

  drawRect(margin + 440, y1 + row1h / 2, 80, row1h / 2);
  setFont(5.5);
  doc.setTextColor(80, 80, 80);
  text("WEEK ENDING", margin + 442, y1 + row1h - 5);
  setFont(8, true);
  doc.setTextColor(0, 0, 0);
  text(data.week_ending || "", margin + 444, y1 + row1h - 15);

  // ── ROW 2
  const y2 = y1 - 20;
  drawRect(margin, y2, 240, 20);
  setFont(5.5);
  doc.setTextColor(80, 80, 80);
  text("PROJECT OR CONTRACT NO.", margin + 2, y2 + 13);
  setFont(7.5);
  doc.setTextColor(0, 0, 0);
  text(data.contract_number || "", margin + 4, y2 + 5);

  drawRect(margin + 240, y2, 200, 20);
  setFont(5.5);
  doc.setTextColor(80, 80, 80);
  text("PROJECT OR CONTRACT NO.", margin + 242, y2 + 13);
  setFont(7.5);
  doc.setTextColor(0, 0, 0);
  text(data.project_number || "", margin + 244, y2 + 5);

  // ── COLUMN HEADERS
  const y3 = y2 - 32;
  const headerH = 32;
  const colWidths = [20, 108, 55, 22, 22, 22, 22, 22, 22, 22, 30, 52, 34, 34];
  const colHeaders = [
    "(1)\nNO.", "(2)\nNAME / SS LAST 4", "(3)\nCLASSIFICATION",
    "S", "M", "T", "W", "T", "F", "S", "TOT",
    "(5)\nTOTAL HRS", "(6)\nRATE", "(7)\nGROSS"
  ];

  let xp = margin;
  doc.setDrawColor(0, 0, 0);
  colWidths.forEach((cw, i) => {
    drawRect(xp, y3, cw, headerH);
    setFont(5);
    doc.setTextColor(0, 0, 0);
    const lines = colHeaders[i].split("\n");
    lines.forEach((ln, li) => {
      doc.text(ln, xp + cw / 2, y3 + headerH - 6 - li * 7, { align: "center" });
    });
    xp += cw;
  });

  // ── EMPLOYEE ROWS
  const employees = data.employees || [];
  const rowH = 30;
  let yEmp = y3 - rowH;

  employees.slice(0, 7).forEach((emp, idx) => {
    let xr = margin;
    const num = String(idx + 1);
    const hours = emp.daily_hours || [0, 0, 0, 0, 0, 0, 0];
    const totalH = hours.reduce((a, b) => a + (Number(b) || 0), 0);

    // (1) number
    drawRect(xr, yEmp, colWidths[0], rowH);
    setFont(7);
    doc.text(num, xr + colWidths[0] / 2, yEmp + rowH / 2 + 2, { align: "center" });
    xr += colWidths[0];

    // (2) name + SS
    drawRect(xr, yEmp, colWidths[1], rowH);
    setFont(7, true);
    doc.text((emp.name || "").substring(0, 20), xr + 2, yEmp + rowH - 8);
    setFont(6);
    doc.text(`***-**-${emp.ss_last4 || "****"}`, xr + 2, yEmp + rowH - 18);
    xr += colWidths[1];

    // (3) classification
    drawRect(xr, yEmp, colWidths[2], rowH);
    setFont(6);
    const classLines = doc.splitTextToSize(emp.classification || "", colWidths[2] - 4);
    classLines.slice(0, 3).forEach((ln, li) => {
      doc.text(ln, xr + 2, yEmp + rowH - 8 - li * 7);
    });
    xr += colWidths[2];

    // daily hours
    for (let d = 0; d < 7; d++) {
      drawRect(xr, yEmp, colWidths[3 + d], rowH);
      const hv = hours[d] || 0;
      if (hv) {
        setFont(7);
        doc.text(String(hv), xr + colWidths[3 + d] / 2, yEmp + rowH / 2 + 2, { align: "center" });
      }
      xr += colWidths[3 + d];
    }

    // total
    drawRect(xr, yEmp, colWidths[10], rowH);
    setFont(7, true);
    doc.text(String(totalH), xr + colWidths[10] / 2, yEmp + rowH / 2 + 2, { align: "center" });
    xr += colWidths[10];

    // (5) total hours detail
    drawRect(xr, yEmp, colWidths[11], rowH);
    setFont(6);
    doc.text(`ST: ${emp.st_hours || totalH}`, xr + 2, yEmp + rowH - 8);
    if (emp.ot_hours) doc.text(`OT: ${emp.ot_hours}`, xr + 2, yEmp + rowH - 16);
    xr += colWidths[11];

    // (6) rate
    drawRect(xr, yEmp, colWidths[12], rowH);
    setFont(6);
    doc.text(`$${emp.rate_of_pay || ""}`, xr + 2, yEmp + rowH - 8);
    if (emp.ot_rate) doc.text(`$${emp.ot_rate} OT`, xr + 2, yEmp + rowH - 16);
    xr += colWidths[12];

    // (7) gross
    drawRect(xr, yEmp, colWidths[13], rowH);
    setFont(7, true);
    doc.text(`$${emp.gross_wages || ""}`, xr + 2, yEmp + rowH / 2 + 2);

    yEmp -= rowH;
  });

  // ── DEDUCTIONS HEADER
  const dedCols = [20, 108, 55, 32, 32, 32, 32, 32, 32, 38, 38];
  const dedLabels = ["#", "NAME", "CLASS", "FICA", "FED TAX", "STATE TAX", "OTHER1", "OTHER2", "OTHER3", "TOTAL DED.", "NET WAGES"];
  const yDed = yEmp - 5;
  const dedHdrH = 18;

  let xd = margin;
  dedCols.forEach((cw, i) => {
    drawRect(xd, yDed - dedHdrH, cw, dedHdrH);
    setFont(5);
    doc.setTextColor(0, 0, 0);
    doc.text(dedLabels[i], xd + cw / 2, yDed - dedHdrH + 11, { align: "center" });
    xd += cw;
  });

  let yDed2 = yDed - dedHdrH;
  employees.slice(0, 7).forEach((emp, idx) => {
    const rowH2 = 14;
    const vals = [
      String(idx + 1), (emp.name || "").substring(0, 18), (emp.classification || "").substring(0, 10),
      emp.fica, emp.fed_tax, emp.state_tax,
      emp.other_ded1, emp.other_ded2, emp.other_ded3,
      emp.total_deductions, emp.net_wages
    ];
    let xdr = margin;
    dedCols.forEach((cw, i) => {
      drawRect(xdr, yDed2 - rowH2, cw, rowH2);
      if (vals[i]) {
        setFont(5.5);
        const pref = i >= 3 ? "$" : "";
        doc.text(`${pref}${vals[i]}`.substring(0, Math.floor(cw / 3.5)), xdr + cw / 2, yDed2 - rowH2 + 8, { align: "center" });
      }
      xdr += cw;
    });
    yDed2 -= rowH2;
  });

  // ── STATEMENT OF COMPLIANCE
  const yCert = yDed2 - 18;
  setFont(7, true);
  doc.text("STATEMENT OF COMPLIANCE", margin, yCert);
  setFont(6);
  const complianceText = `I, the undersigned, do hereby state: (1) That I pay or supervise the payment of the persons employed by ${data.contractor_name || "__________________"} on the ${data.project_name || "__________________"}; that during the payroll period commencing on ${data.period_start || "______"} and ending ${data.week_ending || "______"} all persons employed on said project have been paid the full weekly wages earned, that no rebates have been or will be made either directly or indirectly to or on behalf of said contractor from the full weekly wages earned by any person and that no deductions have been made either directly or indirectly from the full wages earned by any person, except such payroll deductions as are permitted by regulations issued by the Secretary of Labor under the Copeland Act (29 CFR Part 3).`;
  const cLines = doc.splitTextToSize(complianceText, W - 2 * margin);
  let yCl = yCert - 10;
  cLines.slice(0, 5).forEach(ln => {
    doc.text(ln, margin, yCl);
    yCl -= 8;
  });

  // Signature lines
  const ySig = yCl - 20;
  doc.line(margin, ySig, margin + 190, ySig);
  doc.line(margin + 230, ySig, margin + 430, ySig);
  setFont(6);
  doc.text("Signature", margin, ySig + 8);
  doc.text("Title", margin + 230, ySig + 8);

  doc.line(margin, ySig - 28, margin + 190, ySig - 28);
  doc.line(margin + 230, ySig - 28, margin + 340, ySig - 28);
  doc.text("Name (Please Print)", margin, ySig - 22);
  doc.text("Date", margin + 230, ySig - 22);

  setFont(8, true);
  doc.text(data.signatory_name || "", margin + 2, ySig - 4);
  doc.text(data.signatory_title || "", margin + 232, ySig - 4);
  doc.text(data.signature_date || "", margin + 232, ySig - 32);

  // Footer
  setFont(5);
  doc.setTextColor(120, 120, 120);
  doc.text("THE WILLFUL FALSIFICATION OF ANY OF THE ABOVE STATEMENTS MAY SUBJECT THE CONTRACTOR OR SUBCONTRACTOR TO CIVIL OR CRIMINAL PROSECUTION. SEE SECTION 1001 OF TITLE 18 AND SECTION 231 OF TITLE 31 OF THE UNITED STATES CODE.", W / 2, 18, { align: "center", maxWidth: W - 60 });

  return doc.output("blob");
}

// ── Main App Component ────────────────────────────────────────────────────────
export default function WH347Platform() {
  const [step, setStep] = useState("upload"); // upload | parsing | review | done
  const [parsedData, setParsedData] = useState(null);
  const [editedData, setEditedData] = useState(null);
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [fileName, setFileName] = useState("");
  const [pdfBlob, setPdfBlob] = useState(null);
  const fileRef = useRef();

  const parseFile = useCallback(async (file) => {
    setStep("parsing");
    setError("");
    setFileName(file.name);

    try {
      // Read file as text or base64
      const isCSV = file.type === "text/csv" || file.name.endsWith(".csv");
      const isText = isCSV || file.name.endsWith(".txt") || file.name.endsWith(".tsv");

      let content;
      if (isText) {
        content = await file.text();
      } else {
        // For xlsx, pdf, etc - read as base64
        content = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result.split(",")[1]);
          reader.onerror = reject;
          reader.readAsDataURL(file);
        });
      }

      const messages = isText
        ? [{ role: "user", content: `Parse this payroll report and extract WH-347 data:\n\n${content}` }]
        : [{
            role: "user",
            content: [
              {
                type: file.type === "application/pdf" ? "document" : "image",
                source: {
                  type: "base64",
                  media_type: file.type || "application/octet-stream",
                  data: content
                }
              },
              { type: "text", text: "Parse this payroll report and extract WH-347 data." }
            ]
          }];

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: ANTHROPIC_MODEL,
          max_tokens: 4000,
          system: PARSE_SYSTEM,
          messages
        })
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.error?.message || "API error");
      }

      const result = await response.json();
      const rawText = result.content.map(b => b.text || "").join("").trim();
      const clean = rawText.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);

      setParsedData(parsed);
      setEditedData(JSON.parse(JSON.stringify(parsed)));
      setStep("review");
    } catch (e) {
      setError(`Error: ${e.message}`);
      setStep("upload");
    }
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) parseFile(file);
  }, [parseFile]);

  const handleFileChange = useCallback((e) => {
    const file = e.target.files[0];
    if (file) parseFile(file);
  }, [parseFile]);

  const handleGenerate = async () => {
    try {
      setStep("parsing");
      // Load jsPDF if not loaded
      if (!window.jspdf) {
        await new Promise((resolve, reject) => {
          const s = document.createElement("script");
          s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
          s.onload = resolve;
          s.onerror = reject;
          document.head.appendChild(s);
        });
      }
      const blob = await generateWH347PDF(editedData);
      setPdfBlob(blob);
      setStep("done");
    } catch (e) {
      setError(`PDF generation error: ${e.message}`);
      setStep("review");
    }
  };

  const handleDownload = () => {
    if (!pdfBlob) return;
    const url = URL.createObjectURL(pdfBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `WH347_Payroll_${editedData?.week_ending?.replace(/\//g, "-") || "form"}.pdf`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const updateField = (path, value) => {
    setEditedData(prev => {
      const next = JSON.parse(JSON.stringify(prev));
      const parts = path.split(".");
      let obj = next;
      for (let i = 0; i < parts.length - 1; i++) {
        const p = parts[i];
        if (p.includes("[")) {
          const [key, idx] = p.replace("]", "").split("[");
          obj = obj[key][parseInt(idx)];
        } else {
          obj = obj[p];
        }
      }
      obj[parts[parts.length - 1]] = value;
      return next;
    });
  };

  const styles = {
    wrap: { fontFamily: "'Georgia', serif", maxWidth: 900, margin: "0 auto", padding: "24px 16px", color: "var(--color-text-primary)" },
    header: { textAlign: "center", marginBottom: 32 },
    logo: { fontSize: 11, letterSpacing: 3, color: "var(--color-text-secondary)", textTransform: "uppercase", marginBottom: 6 },
    title: { fontSize: 26, fontWeight: 600, margin: "0 0 6px", letterSpacing: -0.5 },
    subtitle: { fontSize: 14, color: "var(--color-text-secondary)", margin: 0 },
    card: { background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, padding: "28px 32px", marginBottom: 20 },
    dropZone: {
      border: `2px dashed var(--color-border-${dragOver ? "primary" : "secondary"})`,
      borderRadius: 10,
      padding: "48px 32px",
      textAlign: "center",
      cursor: "pointer",
      background: dragOver ? "var(--color-background-secondary)" : "transparent",
      transition: "all 0.2s"
    },
    dropIcon: { fontSize: 40, marginBottom: 12, display: "block" },
    dropTitle: { fontSize: 16, fontWeight: 500, marginBottom: 6, color: "var(--color-text-primary)" },
    dropSub: { fontSize: 13, color: "var(--color-text-secondary)", lineHeight: 1.6 },
    btn: { background: "var(--color-text-primary)", color: "var(--color-background-primary)", border: "none", borderRadius: 8, padding: "12px 28px", fontSize: 14, fontWeight: 500, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 8 },
    btnOutline: { background: "transparent", color: "var(--color-text-primary)", border: "0.5px solid var(--color-border-primary)", borderRadius: 8, padding: "10px 20px", fontSize: 13, cursor: "pointer" },
    sectionTitle: { fontSize: 12, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: "var(--color-text-secondary)", marginBottom: 14, borderBottom: "0.5px solid var(--color-border-tertiary)", paddingBottom: 8 },
    grid2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 },
    grid3: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 12 },
    fieldGroup: { display: "flex", flexDirection: "column", gap: 4 },
    label: { fontSize: 11, color: "var(--color-text-secondary)", fontWeight: 500 },
    input: { border: "0.5px solid var(--color-border-secondary)", borderRadius: 6, padding: "7px 10px", fontSize: 13, background: "var(--color-background-primary)", color: "var(--color-text-primary)", width: "100%", boxSizing: "border-box" },
    empCard: { border: "0.5px solid var(--color-border-tertiary)", borderRadius: 10, padding: "16px 20px", marginBottom: 14, background: "var(--color-background-secondary)" },
    empName: { fontSize: 15, fontWeight: 600, marginBottom: 10, color: "var(--color-text-primary)" },
    badge: { display: "inline-block", fontSize: 10, padding: "3px 10px", borderRadius: 20, background: "var(--color-background-info)", color: "var(--color-text-info)", marginLeft: 8 },
    spinner: { width: 36, height: 36, border: "3px solid var(--color-border-tertiary)", borderTopColor: "var(--color-text-primary)", borderRadius: "50%", animation: "spin 0.8s linear infinite", margin: "0 auto 16px" },
    error: { background: "var(--color-background-danger)", color: "var(--color-text-danger)", border: "0.5px solid var(--color-border-danger)", borderRadius: 8, padding: "12px 16px", fontSize: 13, marginBottom: 16 },
    successCard: { textAlign: "center", padding: 48 },
    checkCircle: { width: 64, height: 64, borderRadius: "50%", background: "var(--color-background-success)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 20px" },
    hoursGrid: { display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6, marginBottom: 10 },
    dayLabel: { fontSize: 10, color: "var(--color-text-secondary)", textAlign: "center", marginBottom: 2 },
    dayInput: { border: "0.5px solid var(--color-border-secondary)", borderRadius: 6, padding: "5px 4px", fontSize: 12, textAlign: "center", background: "var(--color-background-primary)", color: "var(--color-text-primary)", width: "100%" },
  };

  const Field = ({ label, path, value }) => (
    <div style={styles.fieldGroup}>
      <label style={styles.label}>{label}</label>
      <input style={styles.input} value={value || ""} onChange={e => updateField(path, e.target.value)} />
    </div>
  );

  return (
    <div style={styles.wrap}>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } } input:focus { outline: 2px solid var(--color-border-info); }`}</style>

      <div style={styles.header}>
        <div style={styles.logo}>DOL Compliance Suite</div>
        <h1 style={styles.title}>WH-347 Certified Payroll</h1>
        <p style={styles.subtitle}>Upload your payroll report — we'll fill the WH-347 automatically</p>
      </div>

      {/* Steps indicator */}
      <div style={{ display: "flex", justifyContent: "center", gap: 0, marginBottom: 32 }}>
        {["Upload", "Extract", "Review", "Download"].map((s, i) => {
          const stepIdx = ["upload", "parsing", "review", "done"].indexOf(step);
          const active = i === stepIdx;
          const done = i < stepIdx;
          return (
            <div key={s} style={{ display: "flex", alignItems: "center" }}>
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", minWidth: 70 }}>
                <div style={{ width: 28, height: 28, borderRadius: "50%", background: done || active ? "var(--color-text-primary)" : "var(--color-background-secondary)", color: done || active ? "var(--color-background-primary)" : "var(--color-text-secondary)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 600, border: "0.5px solid var(--color-border-secondary)", transition: "all 0.3s" }}>
                  {done ? "✓" : i + 1}
                </div>
                <div style={{ fontSize: 11, marginTop: 4, color: active ? "var(--color-text-primary)" : "var(--color-text-secondary)", fontWeight: active ? 600 : 400 }}>{s}</div>
              </div>
              {i < 3 && <div style={{ width: 40, height: 1, background: "var(--color-border-tertiary)", margin: "0 0 18px" }} />}
            </div>
          );
        })}
      </div>

      {error && <div style={styles.error}>{error}</div>}

      {/* UPLOAD STEP */}
      {step === "upload" && (
        <div style={styles.card}>
          <div
            style={styles.dropZone}
            onDragOver={e => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={handleDrop}
            onClick={() => fileRef.current?.click()}
          >
            <span style={styles.dropIcon}>📄</span>
            <div style={styles.dropTitle}>Drop your payroll report here</div>
            <div style={styles.dropSub}>
              Supports CSV, Excel (.xlsx), PDF payroll exports<br />
              from QuickBooks, ADP, Gusto, Sage, and most payroll systems
            </div>
          </div>
          <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,.pdf,.txt" style={{ display: "none" }} onChange={handleFileChange} />
          <div style={{ textAlign: "center", marginTop: 20 }}>
            <button style={styles.btn} onClick={() => fileRef.current?.click()}>
              Choose File
            </button>
          </div>
          {/* Template download banner */}
          <div style={{ marginTop: 20, padding: "16px 20px", background: "var(--color-background-secondary)", borderRadius: 10, border: "0.5px solid var(--color-border-secondary)", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
            <div style={{ flex: 1, minWidth: 200 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4, color: "var(--color-text-primary)", display: "flex", alignItems: "center", gap: 6 }}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                Need a starting point?
              </div>
              <div style={{ fontSize: 12, color: "var(--color-text-secondary)", lineHeight: 1.6 }}>
                Download our template showing every field the WH-347 requires — with sample data so you know exactly how to format your payroll export before uploading.
              </div>
            </div>
            <button onClick={downloadTemplate} style={{ background: "var(--color-text-primary)", color: "var(--color-background-primary)", border: "none", borderRadius: 8, padding: "10px 18px", fontSize: 12, fontWeight: 500, cursor: "pointer", display: "flex", alignItems: "center", gap: 7, whiteSpace: "nowrap", flexShrink: 0 }}>
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Download Template (.xlsx)
            </button>
          </div>

          <div style={{ marginTop: 12, padding: "12px 16px", background: "var(--color-background-secondary)", borderRadius: 8 }}>
            <div style={{ fontSize: 11, fontWeight: 600, marginBottom: 6, color: "var(--color-text-secondary)" }}>FIELDS WE EXTRACT FROM YOUR REPORT</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
              {["Contractor & project info", "Employee names & SS last 4", "Work classifications", "Daily/weekly hours", "Straight & overtime rates", "Gross wages & deductions"].map(item => (
                <div key={item} style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "flex", gap: 6 }}>
                  <span style={{ color: "var(--color-text-success)" }}>✓</span> {item}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* PARSING STEP */}
      {step === "parsing" && (
        <div style={{ ...styles.card, textAlign: "center", padding: 64 }}>
          <div style={styles.spinner} />
          <div style={{ fontSize: 16, fontWeight: 500, marginBottom: 8 }}>Analyzing payroll data…</div>
          <div style={{ fontSize: 13, color: "var(--color-text-secondary)" }}>{fileName && `Reading ${fileName}`}</div>
          <div style={{ fontSize: 12, color: "var(--color-text-secondary)", marginTop: 8 }}>Extracting fields for WH-347 form</div>
        </div>
      )}

      {/* REVIEW STEP */}
      {step === "review" && editedData && (
        <>
          <div style={styles.card}>
            <div style={styles.sectionTitle}>Project & Contractor Info</div>
            <div style={styles.grid2}>
              <Field label="Contractor Name" path="contractor_name" value={editedData.contractor_name} />
              <Field label="Contractor Address" path="contractor_address" value={editedData.contractor_address} />
              <Field label="Project Name" path="project_name" value={editedData.project_name} />
              <Field label="Project Location" path="project_location" value={editedData.project_location} />
              <Field label="Contract Number" path="contract_number" value={editedData.contract_number} />
              <Field label="Project Number" path="project_number" value={editedData.project_number} />
            </div>
            <div style={styles.grid3}>
              <Field label="Payroll Number" path="payroll_number" value={editedData.payroll_number} />
              <Field label="Week Ending (MM/DD/YYYY)" path="week_ending" value={editedData.week_ending} />
              <Field label="Period Start (MM/DD/YYYY)" path="period_start" value={editedData.period_start} />
            </div>
          </div>

          <div style={styles.card}>
            <div style={styles.sectionTitle}>Signatory</div>
            <div style={styles.grid3}>
              <Field label="Signatory Name" path="signatory_name" value={editedData.signatory_name} />
              <Field label="Title" path="signatory_title" value={editedData.signatory_title} />
              <Field label="Signature Date (MM/DD/YYYY)" path="signature_date" value={editedData.signature_date} />
            </div>
          </div>

          <div style={styles.card}>
            <div style={styles.sectionTitle}>Employees ({(editedData.employees || []).length} found)</div>
            {(editedData.employees || []).map((emp, idx) => (
              <div key={idx} style={styles.empCard}>
                <div style={styles.empName}>
                  {emp.name || `Employee ${idx + 1}`}
                  <span style={styles.badge}>{emp.classification}</span>
                </div>
                <div style={styles.grid3}>
                  <Field label="Name (Last, First)" path={`employees[${idx}].name`} value={emp.name} />
                  <Field label="Last 4 SS" path={`employees[${idx}].ss_last4`} value={emp.ss_last4} />
                  <Field label="Work Classification" path={`employees[${idx}].classification`} value={emp.classification} />
                </div>

                <div style={{ marginBottom: 10 }}>
                  <div style={styles.label}>Daily Hours (Sun – Sat)</div>
                  <div style={styles.hoursGrid}>
                    {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((d, di) => (
                      <div key={d}>
                        <div style={styles.dayLabel}>{d}</div>
                        <input
                          type="number"
                          min="0"
                          max="24"
                          style={styles.dayInput}
                          value={(emp.daily_hours || [])[di] ?? 0}
                          onChange={e => {
                            const hrs = [...(emp.daily_hours || [0,0,0,0,0,0,0])];
                            hrs[di] = Number(e.target.value);
                            updateField(`employees[${idx}].daily_hours`, hrs);
                          }}
                        />
                      </div>
                    ))}
                  </div>
                </div>

                <div style={styles.grid3}>
                  <Field label="ST Hours" path={`employees[${idx}].st_hours`} value={emp.st_hours} />
                  <Field label="OT Hours" path={`employees[${idx}].ot_hours`} value={emp.ot_hours} />
                  <Field label="Rate of Pay ($/hr)" path={`employees[${idx}].rate_of_pay`} value={emp.rate_of_pay} />
                  <Field label="OT Rate ($/hr)" path={`employees[${idx}].ot_rate`} value={emp.ot_rate} />
                  <Field label="Gross Wages" path={`employees[${idx}].gross_wages`} value={emp.gross_wages} />
                  <Field label="Net Wages" path={`employees[${idx}].net_wages`} value={emp.net_wages} />
                </div>

                <div style={{ ...styles.grid3, gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr 1fr" }}>
                  {["fica", "fed_tax", "state_tax", "other_ded1", "other_ded2", "total_deductions"].map(f => (
                    <Field key={f} label={f.replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase())} path={`employees[${idx}].${f}`} value={emp[f]} />
                  ))}
                </div>
              </div>
            ))}
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
            <button style={styles.btnOutline} onClick={() => { setStep("upload"); setParsedData(null); setEditedData(null); }}>
              ← Start Over
            </button>
            <button style={styles.btn} onClick={handleGenerate}>
              Generate WH-347 PDF →
            </button>
          </div>
        </>
      )}

      {/* DONE STEP */}
      {step === "done" && (
        <div style={styles.card}>
          <div style={styles.successCard}>
            <div style={styles.checkCircle}>
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="var(--color-text-success)" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="20 6 9 17 4 12"/>
              </svg>
            </div>
            <h2 style={{ fontSize: 22, fontWeight: 600, margin: "0 0 8px" }}>WH-347 Ready</h2>
            <p style={{ fontSize: 14, color: "var(--color-text-secondary)", margin: "0 0 28px" }}>
              Your certified payroll form has been generated for week ending {editedData?.week_ending}.<br/>
              Review, sign, and submit to the appropriate agency.
            </p>
            <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap" }}>
              <button style={styles.btn} onClick={handleDownload}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                Download PDF
              </button>
              <button style={styles.btnOutline} onClick={() => setStep("review")}>
                Edit Data
              </button>
              <button style={styles.btnOutline} onClick={() => { setStep("upload"); setParsedData(null); setEditedData(null); setPdfBlob(null); }}>
                New Payroll
              </button>
            </div>
            <div style={{ marginTop: 28, padding: "14px 20px", background: "var(--color-background-secondary)", borderRadius: 8, textAlign: "left" }}>
              <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6, color: "var(--color-text-secondary)" }}>NEXT STEPS</div>
              {["Print and review for accuracy", "Have authorized signatory sign the form", "Submit to contracting officer with weekly payroll", "Retain copy for 3 years per DOL requirements"].map((s, i) => (
                <div key={i} style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "flex", gap: 8, marginBottom: 4 }}>
                  <span style={{ fontWeight: 600, minWidth: 16 }}>{i + 1}.</span> {s}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      <div style={{ textAlign: "center", marginTop: 24, fontSize: 11, color: "var(--color-text-secondary)", lineHeight: 1.8 }}>
        Always verify generated forms for accuracy before submission.<br/>
        This tool is provided as a convenience — final compliance responsibility rests with the contractor.
      </div>
    </div>
  );
}
