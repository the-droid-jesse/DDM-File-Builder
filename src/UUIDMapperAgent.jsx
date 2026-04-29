import { useState, useRef } from "react";

const ANTHROPIC_MODEL = "claude-sonnet-4-20250514";

// PAR Retail brand tokens — light mode
const C = {
  bg:        "#f4f5f9",
  surface:   "#ffffff",
  surface2:  "#f0f1f7",
  border:    "#d9dae8",
  borderDim: "#e8e9f2",
  orange:    "#f97316",
  orangeDim: "#fff0e6",
  purple:    "#7c3aed",
  purpleDim: "#f0ebff",
  blue:      "#2563eb",
  green:     "#16a34a",
  amber:     "#d97706",
  text:      "#0f0e2a",
  muted:     "#4b5563",
  dim:       "#6b7280",
  faint:     "#9ca3af",
};

// ── Retailer MAN lookup (source: Confluence SUPW/pages/21484208129) ──────────
const RETAILER_LOOKUP = [
  { patterns: ["delek"],                                    name: "Delek",                  man: "070948" },
  { patterns: ["egretailamerica", "egamerica", "egretail"], name: "EG America",             man: "026142" },
  { patterns: ["extramileconvenience", "extramile"],        name: "EMCS",                   man: "068418" },
  { patterns: ["casey"],                                    name: "Caseys",                 man: "028030" },
  { patterns: ["dashin"],                                   name: "Dash In",                man: "041309" },
  { patterns: ["hsenergyproducts", "hsenergy"],             name: "HNS",                    man: "072797" },
  { patterns: ["hucksconvenience", "hucks"],                name: "Hucks",                  man: "036376" },
  { patterns: ["saneholtz", "mckarns", "saneholtzmckarns"], name: "Marathon (Sane Holtz)",  man: "053447" },
  { patterns: ["jacksons"],                                 name: "Jacksons",               man: "069824" },
  { patterns: ["parkers"],                                  name: "Parkers",                man: "049394" },
  { patterns: ["raceway", "racewayventure"],                name: "Raceway",                man: "079170" },
  { patterns: ["refuel"],                                   name: "Refuel",                 man: "076353" },
  { patterns: ["royalfarms", "cloverlandfarms"],            name: "Royal Farms",            man: "015512" },
  { patterns: ["stinker"],                                  name: "Stinker",                man: null     },
  { patterns: ["brookwood", "bwgas", "yesway"],             name: "Yesway",                 man: "072605" },
];

function detectRetailerFromFilename(filename) {
  // Strip extension, then strip trailing date block (8+ digits) and everything after
  const base = filename
    .replace(/\.[^.]+$/, "")
    .replace(/\d{8,}.*$/, "")
    .toLowerCase()
    .replace(/[^a-z]/g, "");
  for (const entry of RETAILER_LOOKUP) {
    if (entry.patterns.some(p => base.includes(p))) return entry;
  }
  return null;
}

// ── AGDC field definitions (output spec) ─────────────────────────────────────
const AGDC_OUTPUT_FIELDS = [
  { num: 1,  key: "Management Account Number",         label: "Management Account Number (MAN)",                   required: true  },
  { num: 2,  key: "Activity Date",                     label: "Activity Date",                                     required: true  },
  { num: 3,  key: "Activity Time",                     label: "Activity Time",                                     required: true  },
  { num: 4,  key: "Ad ID",                             label: "Ad ID",                                             required: true  },
  { num: 5,  key: "Loyalty ID/Rewards Number",         label: "Loyalty ID/Rewards Number",                        required: true, isLoyaltyId: true },
  { num: 6,  key: "Activity Type",                     label: "Activity Type",                                     required: true  },
  { num: 7,  key: "Channel Type",                      label: "Channel Type",                                      required: true  },
  { num: 8,  key: "Marketing Transaction ID",          label: "Marketing Transaction ID / Activity Log ID",       required: false },
  { num: 9,  key: "Activity State",                    label: "Activity State",                                    required: false },
  { num: 10, key: "OPTIONAL EAIV Indicator",           label: "OPTIONAL EAIV Indicator",                          required: false },
  { num: 11, key: "RESERVE FOR FUTURE USE 11",         label: "RESERVE FOR FUTURE USE",                           required: false },
  { num: 12, key: "RESERVE FOR FUTURE USE 12",         label: "RESERVE FOR FUTURE USE",                           required: false },
  { num: 13, key: "RESERVE FOR FUTURE USE 13",         label: "RESERVE FOR FUTURE USE",                           required: false },
  { num: 14, key: "RESERVE FOR FUTURE USE 14",         label: "RESERVE FOR FUTURE USE",                           required: false },
  { num: 15, key: "RESERVE FOR FUTURE USE 15",         label: "RESERVE FOR FUTURE USE",                           required: false },
];

const OUTPUT_HEADERS = AGDC_OUTPUT_FIELDS.map(f => f.label);

// ── CSV parser ────────────────────────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return { headers: [], rows: [] };
  const headers = lines[0].split(",").map(h => h.trim().replace(/^"|"$/g, ""));
  const rows = lines.slice(1).filter(l => l.trim()).map(line => {
    const values = [];
    let cur = "", inQ = false;
    for (const ch of line) {
      if (ch === '"') { inQ = !inQ; continue; }
      if (ch === "," && !inQ) { values.push(cur.trim()); cur = ""; }
      else cur += ch;
    }
    values.push(cur.trim());
    return Object.fromEntries(headers.map((h, i) => [h, values[i] ?? ""]));
  });
  return { headers, rows };
}

// ── XLSX parser (SheetJS from cdnjs) ───────────────────────────────────────────
async function parseXLSX(arrayBuffer) {
  if (!window._XLSX) {
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
    window._XLSX = window.XLSX;
  }
  const workbook = window._XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
  if (!workbook.SheetNames.length) return { headers: [], rows: [] };
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  if (!worksheet) return { headers: [], rows: [] };
  const rawData = window._XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  if (!rawData || rawData.length < 1) return { headers: [], rows: [] };
  let headerRowIdx = 0;
  while (headerRowIdx < rawData.length && rawData[headerRowIdx].every(c => String(c).trim() === "")) {
    headerRowIdx++;
  }
  if (headerRowIdx >= rawData.length) return { headers: [], rows: [] };
  let rawHeaders = rawData[headerRowIdx].map(h => String(h).trim());
  const headers = rawHeaders.map((h, i) => h || `Column_${i + 1}`);
  const rows = rawData.slice(headerRowIdx + 1)
    .filter(rowArr => rowArr.some(cell => String(cell).trim() !== ""))
    .map(rowArr => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = String(rowArr[i] ?? "").trim(); });
      return obj;
    });
  return { headers, rows };
}

// ── CSV output ────────────────────────────────────────────────────────────────
function toCSV(headers, rows) {
  const esc = v => { const s = String(v ?? ""); return s.includes(",") || s.includes('"') ? `"${s.replace(/"/g, '""')}"` : s; };
  return [headers.join(","), ...rows.map(r => headers.map(h => esc(r[h])).join(","))].join("\n");
}

// ── normalization ─────────────────────────────────────────────────────────────
function normalizePhone(p) {
  const d = String(p ?? "").replace(/\D/g, "");
  return d.length === 11 && d.startsWith("1") ? d.slice(1) : d;
}
function normalizeEmail(e) { return String(e ?? "").trim().toLowerCase(); }
function normalizeKey(v, t) { return t === "email" ? normalizeEmail(v) : normalizePhone(v); }

// ── fuzzy column matcher ──────────────────────────────────────────────────────
function fuzzyMatch(agdcKey, ddmHeaders) {
  const needle = agdcKey.toLowerCase().replace(/[^a-z0-9]/g, "");
  let match = ddmHeaders.find(h => h.toLowerCase().replace(/[^a-z0-9]/g, "") === needle);
  if (match) return match;
  match = ddmHeaders.find(h => {
    const hay = h.toLowerCase().replace(/[^a-z0-9]/g, "");
    return hay.includes(needle) || needle.includes(hay);
  });
  return match ?? null;
}

// ── component ─────────────────────────────────────────────────────────────────
export default function UUIDMapperAgent() {
  const [uuidFile,      setUuidFile]      = useState(null);
  const [ddmFile,       setDdmFile]       = useState(null);
  const [matchType,     setMatchType]     = useState("phone");
  const [refIdCol,      setRefIdCol]      = useState("");
  const [refUuidCol,    setRefUuidCol]    = useState("");
  const [ddmIdCol,      setDdmIdCol]      = useState("");
  const [manAcctNum,    setManAcctNum]    = useState("");
  const [retailerName,  setRetailerName]  = useState("");
  const [selectedMonth, setSelectedMonth] = useState(String(new Date().getMonth() + 1).padStart(2, "0"));
  const [selectedYear,  setSelectedYear]  = useState(String(new Date().getFullYear()));
  const [colMap,        setColMap]        = useState({});
  const [messages,      setMessages]      = useState([]);
  const [input,         setInput]         = useState("");
  const [loading,       setLoading]       = useState(false);
  const [result,        setResult]        = useState(null);
  const [showCopy,      setShowCopy]      = useState(false);
  const [showColMap,    setShowColMap]    = useState(false);
  const [fallbackState, setFallbackState] = useState("");
  const [autoFilled,    setAutoFilled]    = useState(false);
  const chatRef = useRef(null);
  const copyRef = useRef(null);

  const addMsg = (role, text) =>
    setMessages(prev => [...prev, { role, text, ts: Date.now() }]);

  // ── file reading ─────────────────────────────────────────────────────────────
  const readFile = (file, setter, label) => {
    const isXLSX = /\.xlsx?$/i.test(file.name);
    const finish = ({ headers, rows }) => {
      setter({ name: file.name, headers, rows });
      if (label === "DDM") {
        const auto = {};
        AGDC_OUTPUT_FIELDS.forEach(f => {
          if (f.num === 1 || f.num >= 11) return;
          const found = fuzzyMatch(f.key, headers);
          if (found) auto[f.key] = found;
        });
        setColMap(auto);
      }

      // Auto-fill retailer name + MAN from filename if not already set manually
      const match = detectRetailerFromFilename(file.name);
      if (match) {
        let filled = false;
        if (!retailerName.trim()) { setRetailerName(match.name); filled = true; }
        if (!manAcctNum.trim() && match.man) { setManAcctNum(match.man); filled = true; }
        if (filled) {
          setAutoFilled(true);
          setTimeout(() => setAutoFilled(false), 4000);
        }
        addMsg("agent",
          `✅ Loaded ${label}: ${file.name} — ${rows.length} rows\nColumns: ${headers.join(", ")}\n\n` +
          `🏪 Detected retailer: **${match.name}**${match.man ? ` · MAN: ${match.man}` : " · MAN not on file"}`
        );
      } else {
        addMsg("agent", `✅ Loaded ${label}: ${file.name} — ${rows.length} rows\nColumns: ${headers.join(", ")}`);
      }
    };
    if (isXLSX) {
      const reader = new FileReader();
      reader.onload = async e => {
        try { finish(await parseXLSX(e.target.result)); }
        catch (err) { addMsg("agent", `⚠️ Could not parse ${file.name}: ${err.message}`); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = e => finish(parseCSV(e.target.result));
      reader.readAsText(file);
    }
  };

  // ── AI chat ───────────────────────────────────────────────────────────────────
  const buildSystem = () =>
    `You are a data mapping copilot for PAR Retail working with AGDC Digital Trade Program Activity Log files.
Output must have exactly 15 fields in this order: ${AGDC_OUTPUT_FIELDS.map(f => `${f.num}. ${f.label}`).join(", ")}.
Field 5 (Loyalty ID/Rewards Number) gets the UUID looked up by matching ${matchType} across both files.
Field 1 (Management Account Number) is entered manually by the user.
Fields 11–15 are always blank (RESERVE FOR FUTURE USE).
Reference file columns: ${uuidFile?.headers?.join(", ") ?? "not loaded"}
DDM file columns: ${ddmFile?.headers?.join(", ") ?? "not loaded"}
Be concise. User is a data-savvy PM at PAR Retail.`;

  const sendMessage = async text => {
    if (!text.trim()) return;
    addMsg("user", text);
    setInput("");
    setLoading(true);
    const history = [...messages, { role: "user", text }];
    try {
      const res = await fetch("/anthropic-api/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: ANTHROPIC_MODEL, max_tokens: 1000,
          system: buildSystem(),
          messages: history.map(m => ({ role: m.role === "agent" ? "assistant" : "user", content: m.text })),
        }),
      });
      const data = await res.json();
      addMsg("agent", data.content?.find(b => b.type === "text")?.text ?? "No response.");
    } catch { addMsg("agent", "⚠️ Network error."); }
    setLoading(false);
    setTimeout(() => chatRef.current?.scrollTo({ top: 99999, behavior: "smooth" }), 100);
  };

  // ── auto-detect ───────────────────────────────────────────────────────────────
  const autoDetect = async () => {
    if (!uuidFile || !ddmFile) return;
    setLoading(true);
    addMsg("agent", "🔍 Auto-detecting columns...");
    try {
      const res = await fetch("/anthropic-api/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: ANTHROPIC_MODEL, max_tokens: 400,
          messages: [{ role: "user", content:
            `Match type: ${matchType}
Ref file columns: ${uuidFile.headers.join(", ")}
Ref sample row: ${JSON.stringify(uuidFile.rows[0])}
DDM file columns: ${ddmFile.headers.join(", ")}
DDM sample row: ${JSON.stringify(ddmFile.rows[0])}

Return ONLY valid JSON, no markdown:
{
  "refIdCol": "column in ref file containing ${matchType}",
  "refUuidCol": "column in ref file containing UUID",
  "ddmIdCol": "column in DDM file containing ${matchType} to be replaced by UUID",
  "colMap": {
    "Activity Date": "matching DDM column or null",
    "Activity Time": "matching DDM column or null",
    "Ad ID": "matching DDM column or null",
    "Activity Type": "matching DDM column or null",
    "Channel Type": "matching DDM column or null",
    "Marketing Transaction ID": "matching DDM column or null",
    "Loyalty ID/Rewards Number": "matching DDM column or null",
    "Activity State": "matching DDM column or null",
    "OPTIONAL EAIV Indicator": "matching DDM column or null"
  }
}`
          }],
        }),
      });
      const data = await res.json();
      const raw = data.content?.find(b => b.type === "text")?.text ?? "";
      const d = JSON.parse(raw.replace(/```json|```/g, "").trim());
      if (d.refIdCol)   setRefIdCol(d.refIdCol);
      if (d.refUuidCol) setRefUuidCol(d.refUuidCol);
      if (d.ddmIdCol)   setDdmIdCol(d.ddmIdCol);
      if (d.colMap) {
        const cleaned = {};
        Object.entries(d.colMap).forEach(([k, v]) => { if (v && v !== "null") cleaned[k] = v; });
        setColMap(cleaned);
      }
      addMsg("agent",
        `🤖 Auto-detected:\n` +
        `- Ref ${matchType} col: ${d.refIdCol}\n` +
        `- Ref UUID col: ${d.refUuidCol}\n` +
        `- DDM ${matchType} col: ${d.ddmIdCol}\n` +
        `- Field mappings: ${Object.entries(d.colMap || {}).filter(([,v]) => v && v !== "null").map(([k,v]) => `${k} → ${v}`).join(", ") || "none detected"}\n\n` +
        `Review and click ▶ Run Mapping.`
      );
    } catch (e) { addMsg("agent", `⚠️ Auto-detect failed: ${e.message}`); }
    setLoading(false);
  };

  // ── mapping ───────────────────────────────────────────────────────────────────
  const runMapping = () => {
    if (!ddmFile) { addMsg("agent", "⚠️ Load the DDM file first."); return; }
    const lookupEnabled = uuidFile && refIdCol && refUuidCol && ddmIdCol;
    const lookup = new Map();
    if (lookupEnabled) {
      for (const row of uuidFile.rows) {
        const key = normalizeKey(row[refIdCol], matchType);
        const uuid = (row[refUuidCol] ?? "").trim();
        if (key && uuid) lookup.set(key, uuid);
      }
    }
    const sampleRef = lookupEnabled ? [...lookup.keys()].slice(0, 3) : [];
    const sampleDDM = lookupEnabled ? ddmFile.rows.slice(0, 3).map(r => normalizeKey(r[ddmIdCol], matchType)) : [];
    let matched = 0, unmatched = 0;
    const newRows = ddmFile.rows.map(row => {
      let uuid = undefined;
      if (lookupEnabled) {
        uuid = lookup.get(normalizeKey(row[ddmIdCol], matchType));
        uuid ? matched++ : unmatched++;
      } else { unmatched++; }
      const out = {};
      AGDC_OUTPUT_FIELDS.forEach(f => {
        if (f.num === 1) {
          out[f.label] = manAcctNum.trim();
        } else if (f.isLoyaltyId) {
          const ddmCol = colMap[f.key];
          const original = ddmCol ? (row[ddmCol] ?? "") : "";
          out[f.label] = uuid ? uuid : original;
        } else if (f.num >= 11) {
          out[f.label] = "";
        } else {
          const ddmCol = colMap[f.key];
          let val = ddmCol ? (row[ddmCol] ?? "") : "";
          if (f.num === 9 && !val && fallbackState) val = fallbackState;
          out[f.label] = val;
        }
      });
      return out;
    });
    setResult({ headers: OUTPUT_HEADERS, rows: newRows, stats: { matched, unmatched, total: ddmFile.rows.length } });
    setShowCopy(false);
    let msg = `✅ Done! Output has all 15 AGDC fields in order.\n`;
    if (lookupEnabled) {
      const fb = colMap["Loyalty ID/Rewards Number"] ? "Fallback (DDM value)" : "Blank";
      msg += `- ${matched} rows matched → Field 5 = UUID\n` +
             `- ${unmatched} rows unmatched → Field 5 = ${fb}\n` +
             `- Total: ${ddmFile.rows.length} rows\n\n` +
             `🔍 Ref sample ${matchType}s: ${sampleRef.join(", ") || "none"}\n` +
             `🔍 DDM sample ${matchType}s: ${sampleDDM.join(", ") || "none"}` +
             (matched === 0 ? `\n\n⚠️ 0 matches — paste a sample value from each file so I can diagnose the format difference.` : "");
    } else {
      msg += `⚠️ UUID lookup was skipped (identifier columns not fully mapped).\n- Total: ${ddmFile.rows.length} rows processed.`;
    }
    addMsg("agent", msg);
  };

  // ── download / copy ───────────────────────────────────────────────────────────
  const csvContent = result ? toCSV(result.headers, result.rows) : "";
  const handleDownload = () => {
    const retailer = retailerName.trim().replace(/\s+/g, "") || "Retailer";
    const acct = manAcctNum.trim() || "00000";
    const filename = `${retailer}_${acct}_CampaignActivityLog_${selectedMonth}${selectedYear}.csv`;
    try {
      const a = document.createElement("a");
      a.href = "data:text/csv;charset=utf-8," + encodeURIComponent(csvContent);
      a.download = filename;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
    } catch {
      setShowCopy(true);
      addMsg("agent", "Download unavailable — use Copy CSV Text instead.");
    }
  };
  const handleCopy = () => {
    try {
      navigator.clipboard.writeText(csvContent).then(() => addMsg("agent", "✅ CSV copied to clipboard!"));
    } catch {
      if (copyRef.current) { copyRef.current.select(); document.execCommand("copy"); addMsg("agent", "✅ Copied!"); }
    }
  };

  // ── shared styles ─────────────────────────────────────────────────────────────
  const inputStyle = {
    width: "100%", background: C.surface, border: `1px solid ${C.border}`,
    borderRadius: 6, color: C.text, padding: "6px 9px", fontSize: 12,
    fontFamily: "inherit", outline: "none", boxSizing: "border-box",
  };
  const labelStyle = {
    display: "block", fontSize: 10, fontWeight: 700, letterSpacing: 1,
    color: C.orange, marginBottom: 3, textTransform: "uppercase",
  };
  const sectionStyle = {
    background: C.surface, border: `1px solid ${C.border}`,
    borderRadius: 10, padding: 12,
  };
  const sectionTitle = {
    fontSize: 10, fontWeight: 700, color: C.dim, letterSpacing: 1,
    marginBottom: 8, textTransform: "uppercase",
  };

  // ── sub-components ────────────────────────────────────────────────────────────
  const Sel = ({ label, value, onChange, options, dim }) => (
    <div style={{ marginBottom: 8 }}>
      <label style={{ ...labelStyle, color: dim ? C.faint : C.orange }}>{label}</label>
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ ...inputStyle, color: value ? C.text : C.dim, border: `1px solid ${dim ? C.borderDim : C.border}` }}>
        <option value="">— none —</option>
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );

  const DropZone = ({ label, file, onFile }) => {
    const ref = useRef();
    return (
      <div onClick={() => ref.current.click()}
        onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}
        onDragOver={e => e.preventDefault()}
        style={{
          border: `2px dashed ${file ? C.orange : C.border}`, borderRadius: 9,
          padding: "13px 10px", textAlign: "center", cursor: "pointer",
          background: file ? `${C.orange}11` : C.surface2, marginBottom: 8,
          transition: "all .15s",
        }}>
        <input ref={ref} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
          onChange={e => { const f = e.target.files[0]; if (f) onFile(f); }} />
        <div style={{ fontSize: 18, marginBottom: 2 }}>{file ? "📄" : "📂"}</div>
        <div style={{ fontSize: 11, fontWeight: 700, color: file ? C.orange : C.dim }}>{label}</div>
        {file
          ? <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>{file.name} · {file.rows.length} rows</div>
          : <div style={{ fontSize: 10, color: C.faint, marginTop: 2 }}>Click or drop · CSV / XLSX</div>}
      </div>
    );
  };

  const statsBar = result?.stats && (
    <div style={{ display: "flex", gap: 5, margin: "7px 0" }}>
      {[["Matched", result.stats.matched, C.green], ["Unmatched", result.stats.unmatched, C.amber], ["Total", result.stats.total, C.orange]].map(([l, v, col]) => (
        <div key={l} style={{ flex: 1, background: C.surface2, border: `1px solid ${col}44`, borderRadius: 6, padding: "6px 0", textAlign: "center" }}>
          <div style={{ fontSize: 16, fontWeight: 800, color: col }}>{v}</div>
          <div style={{ fontSize: 9, color: C.dim, letterSpacing: 1, textTransform: "uppercase" }}>{l}</div>
        </div>
      ))}
    </div>
  );

  // ── render ────────────────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Inter',system-ui,sans-serif", color: C.text, display: "flex", flexDirection: "column", alignItems: "center", padding: "20px 10px" }}>

      {/* ── top bar ── */}
      <div style={{ width: "100%", maxWidth: 860, marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {/* PAR Retail wordmark */}
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{
                background: `linear-gradient(135deg, ${C.orange}, ${C.purple})`,
                borderRadius: 8, width: 36, height: 36,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 16, fontWeight: 900, color: "#fff", letterSpacing: -1,
              }}>P</div>
              <div>
                <div style={{ fontSize: 11, fontWeight: 800, color: C.orange, letterSpacing: 2, textTransform: "uppercase", lineHeight: 1 }}>PAR Retail</div>
                <div style={{ fontSize: 9, color: C.dim, letterSpacing: 1, textTransform: "uppercase" }}>DDM File Builder</div>
              </div>
            </div>
          </div>
          <div style={{ fontSize: 10, color: C.faint, letterSpacing: 1, textTransform: "uppercase", border: `1px solid ${C.border}`, borderRadius: 20, padding: "3px 10px" }}>
            AGDC · 15-Field Output
          </div>
        </div>
        {/* orange accent bar */}
        <div style={{ height: 2, background: `linear-gradient(90deg, ${C.orange}, ${C.purple}, transparent)`, borderRadius: 2, marginTop: 12 }} />
      </div>

      <div style={{ width: "100%", maxWidth: 860, display: "flex", gap: 10 }}>

        {/* ── left panel ── */}
        <div style={{ width: 240, flexShrink: 0, display: "flex", flexDirection: "column", gap: 8 }}>

          {/* 1 · Load Files */}
          <div style={sectionStyle}>
            <div style={sectionTitle}>1 · Load Files</div>
            <DropZone label="UUID Reference File" file={uuidFile} onFile={f => readFile(f, setUuidFile, "UUID Reference")} />
            <DropZone label="DDM File" file={ddmFile} onFile={f => readFile(f, setDdmFile, "DDM")} />
          </div>

          {/* 2 · MAN + Retailer */}
          <div style={sectionStyle}>
            <div style={{ ...sectionTitle, display: "flex", alignItems: "center", gap: 6 }}>
              <span>2 · MAN + Retailer Name</span>
              {autoFilled && (
                <span style={{ fontSize: 9, fontWeight: 700, background: `${C.green}22`, color: C.green, border: `1px solid ${C.green}44`, borderRadius: 10, padding: "1px 7px", letterSpacing: .5, textTransform: "uppercase" }}>
                  Auto-filled
                </span>
              )}
            </div>
            <div style={{ marginBottom: 6 }}>
              <label style={labelStyle}>Retailer Name</label>
              <input value={retailerName} onChange={e => setRetailerName(e.target.value)}
                placeholder="e.g. QuickGas" style={{ ...inputStyle, marginBottom: 8 }} />
              <label style={labelStyle}>MAN (Managed Account Number)</label>
              <input value={manAcctNum} onChange={e => setManAcctNum(e.target.value)}
                placeholder="e.g. 012345" style={inputStyle} />
              <div style={{ fontSize: 9, color: C.dim, marginTop: 5, lineHeight: 1.6 }}>
                Filename: <span style={{ color: C.muted }}>
                  {(retailerName.trim().replace(/\s+/g, "") || "Retailer")}_{manAcctNum.trim() || "00000"}_CampaignActivityLog_{selectedMonth}{selectedYear}.csv
                </span>
              </div>
              <div style={{ display: "flex", gap: 5, marginTop: 8 }}>
                <div style={{ flex: 1 }}>
                  <label style={labelStyle}>Month</label>
                  <select value={selectedMonth} onChange={e => setSelectedMonth(e.target.value)}
                    style={{ ...inputStyle }}>
                    {["01","02","03","04","05","06","07","08","09","10","11","12"].map((m, i) => (
                      <option key={m} value={m}>{["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][i]} ({m})</option>
                    ))}
                  </select>
                </div>
                <div style={{ flex: 1 }}>
                  <label style={labelStyle}>Year</label>
                  <select value={selectedYear} onChange={e => setSelectedYear(e.target.value)}
                    style={{ ...inputStyle }}>
                    {[2023,2024,2025,2026,2027].map(y => (
                      <option key={y} value={String(y)}>{y}</option>
                    ))}
                  </select>
                </div>
              </div>
            </div>
          </div>

          {/* 3 · Map Columns */}
          {uuidFile && ddmFile && (
            <div style={sectionStyle}>
              <div style={sectionTitle}>3 · Map Identifier Column</div>
              <div style={{ marginBottom: 8 }}>
                <label style={labelStyle}>Match Key</label>
                <div style={{ display: "flex", gap: 4 }}>
                  {["phone", "email"].map(t => (
                    <button key={t} onClick={() => { setMatchType(t); setRefIdCol(""); setDdmIdCol(""); }}
                      style={{
                        flex: 1, padding: "5px 0", borderRadius: 6, fontSize: 10, fontWeight: 700,
                        fontFamily: "inherit", cursor: "pointer", textTransform: "uppercase",
                        background: matchType === t ? `linear-gradient(135deg,${C.orange},${C.purple})` : C.surface2,
                        color: matchType === t ? "#fff" : C.dim,
                        border: matchType === t ? "none" : `1px solid ${C.border}`,
                      }}>
                      {t === "phone" ? "📱 Phone" : "✉️ Email"}
                    </button>
                  ))}
                </div>
              </div>
              <Sel label={`Ref → ${matchType}`}  value={refIdCol}   onChange={setRefIdCol}   options={uuidFile.headers} />
              <Sel label="Ref → UUID col"         value={refUuidCol} onChange={setRefUuidCol} options={uuidFile.headers} />
              <Sel label={`DDM → ${matchType}`}   value={ddmIdCol}   onChange={setDdmIdCol}   options={ddmFile.headers} />

              <button onClick={() => setShowColMap(v => !v)}
                style={{ width: "100%", background: C.surface2, color: C.dim, border: `1px solid ${C.borderDim}`, borderRadius: 5, padding: "5px 0", fontSize: 10, cursor: "pointer", marginBottom: 7, fontFamily: "inherit" }}>
                {showColMap ? "▲ Hide" : "▼ Show"} field column mapping (fields 2–10)
              </button>

              {showColMap && (
                <div style={{ background: C.bg, borderRadius: 6, padding: "8px 8px 2px", marginBottom: 7, border: `1px solid ${C.borderDim}` }}>
                  <div style={{ fontSize: 9, color: C.faint, marginBottom: 6, letterSpacing: .5 }}>MAP DDM COLUMNS → AGDC FIELDS</div>
                  {AGDC_OUTPUT_FIELDS.filter(f => f.num >= 2 && f.num <= 10).map(f => (
                    <div key={f.key} style={{ display: "flex", gap: 6, alignItems: "flex-end" }}>
                      <div style={{ flex: 1 }}>
                        <Sel dim label={`F${f.num}: ${f.key}`}
                          value={colMap[f.key] ?? ""}
                          onChange={v => setColMap(prev => ({ ...prev, [f.key]: v }))}
                          options={ddmFile.headers} />
                      </div>
                      {f.num === 9 && (
                        <button onClick={() => {
                          const s = window.prompt("Enter fallback state (e.g. CA) to auto-fill missing rows for Field 9:");
                          if (s !== null) setFallbackState(s.trim());
                        }} style={{ marginBottom: 8, height: 26, background: fallbackState ? C.green : C.border, color: "#fff", border: "none", borderRadius: 5, padding: "0 10px", fontSize: 10, cursor: "pointer", fontFamily: "inherit" }}>
                          {fallbackState ? `F9: ${fallbackState}` : "➕ F9 State"}
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              )}

              <button onClick={autoDetect} disabled={loading}
                style={{ width: "100%", background: C.surface2, color: C.orange, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 0", fontSize: 11, cursor: "pointer", marginBottom: 6, fontFamily: "inherit", fontWeight: 600 }}>
                🤖 Auto-detect all columns
              </button>
              <button onClick={runMapping} disabled={!ddmFile}
                style={{ width: "100%", background: ddmFile ? `linear-gradient(135deg,${C.orange},${C.purple})` : C.surface2, color: "#fff", border: "none", borderRadius: 6, padding: "9px 0", fontSize: 12, fontWeight: 700, cursor: ddmFile ? "pointer" : "not-allowed", fontFamily: "inherit" }}>
                ▶ Run Mapping
              </button>
            </div>
          )}

          {/* 4 · Export */}
          {result && (
            <div style={{ ...sectionStyle, border: `1px solid ${C.green}55` }}>
              <div style={sectionTitle}>4 · Export</div>
              {statsBar}
              <button onClick={handleDownload}
                style={{ width: "100%", background: `linear-gradient(135deg,${C.orange},${C.purple})`, color: "#fff", border: "none", borderRadius: 6, padding: "8px 0", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", marginBottom: 5 }}>
                ⬇ Download CSV
              </button>
              <button onClick={() => setShowCopy(v => !v)}
                style={{ width: "100%", background: C.surface2, color: C.orange, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 0", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>
                📋 {showCopy ? "Hide" : "Copy"} CSV Text
              </button>
            </div>
          )}
        </div>

        {/* ── chat panel ── */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, overflow: "hidden", minHeight: 520 }}>
          <div style={{ padding: "10px 14px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 7, height: 7, borderRadius: "50%", background: C.green, boxShadow: `0 0 5px ${C.green}` }} />
            <span style={{ fontSize: 10, color: C.dim, letterSpacing: 1, fontWeight: 700, textTransform: "uppercase" }}>Mapping Copilot</span>
            <div style={{ marginLeft: "auto", fontSize: 9, color: C.faint, letterSpacing: .5 }}>Powered by Claude</div>
          </div>

          <div ref={chatRef} style={{ flex: 1, overflowY: "auto", padding: 12, display: "flex", flexDirection: "column", gap: 8 }}>
            {messages.length === 0 && (
              <div style={{ color: C.faint, fontSize: 12, textAlign: "center", marginTop: 50 }}>
                <div style={{ fontSize: 28, marginBottom: 10 }}>🤖</div>
                <div style={{ color: C.dim, marginBottom: 4 }}>Load your files to get started.</div>
                <span style={{ fontSize: 10, color: C.faint }}>CSV & XLSX · Phone or email · 15-field AGDC output</span>
              </div>
            )}
            {messages.map(m => (
              <div key={m.ts} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start" }}>
                <div style={{
                  maxWidth: "88%", borderRadius: 8, padding: "8px 11px", fontSize: 12, lineHeight: 1.6, whiteSpace: "pre-wrap",
                  background: m.role === "user" ? `linear-gradient(135deg,${C.orange},${C.purple})` : C.surface2,
                  border: m.role === "agent" ? `1px solid ${C.border}` : "none",
                  color: C.text,
                }}>
                  {m.text}
                </div>
              </div>
            ))}
            {loading && (
              <div style={{ display: "flex" }}>
                <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 11px", fontSize: 11, color: C.dim }}>⏳ Thinking...</div>
              </div>
            )}
          </div>

          <div style={{ padding: "8px 10px", borderTop: `1px solid ${C.border}`, display: "flex", gap: 6 }}>
            <input value={input} onChange={e => setInput(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendMessage(input); } }}
              placeholder="Ask about columns, formats, or the mapping..."
              style={{ flex: 1, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, color: C.text, padding: "7px 10px", fontSize: 11, fontFamily: "inherit", outline: "none" }} />
            <button onClick={() => sendMessage(input)} disabled={loading || !input.trim()}
              style={{ background: `linear-gradient(135deg,${C.orange},${C.purple})`, border: "none", borderRadius: 6, color: "#fff", padding: "0 14px", fontSize: 14, cursor: loading ? "not-allowed" : "pointer", opacity: (!input.trim() || loading) ? 0.5 : 1 }}>↑</button>
          </div>
        </div>
      </div>

      {/* ── copy textarea ── */}
      {showCopy && result && (
        <div style={{ width: "100%", maxWidth: 860, marginTop: 10, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, overflow: "hidden" }}>
          <div style={{ padding: "7px 12px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 10, color: C.orange, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase" }}>CSV Output</span>
            <button onClick={handleCopy} style={{ background: C.orange, color: "#fff", border: "none", borderRadius: 4, padding: "3px 10px", fontSize: 10, cursor: "pointer", fontFamily: "inherit", fontWeight: 700 }}>Copy All</button>
          </div>
          <textarea ref={copyRef} readOnly value={csvContent}
            style={{ width: "100%", height: 150, background: C.surface2, color: C.muted, border: "none", padding: 10, fontSize: 11, fontFamily: "inherit", resize: "vertical", outline: "none", boxSizing: "border-box" }} />
        </div>
      )}

      {/* ── output preview ── */}
      {result && (
        <div style={{ width: "100%", maxWidth: 860, marginTop: 10, background: C.surface, border: `1px solid ${C.green}44`, borderRadius: 10, overflow: "hidden" }}>
          <div style={{ padding: "8px 12px", borderBottom: `1px solid ${C.green}44`, fontSize: 10, color: C.green, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase" }}>
            Output Preview — first 5 rows · all 15 AGDC fields
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
              <thead>
                <tr>{result.headers.map((h, i) => (
                  <th key={h+i} style={{
                    padding: "6px 9px", textAlign: "left", whiteSpace: "nowrap",
                    borderBottom: `1px solid ${C.borderDim}`,
                    color: h === "Loyalty ID/Rewards Number" ? C.green : h === "Management Account Number (MAN)" ? C.orange : C.blue,
                    background: h === "Loyalty ID/Rewards Number" ? `${C.green}11` : h === "Management Account Number (MAN)" ? `${C.orange}11` : "transparent",
                  }}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {result.rows.slice(0, 5).map((row, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.borderDim}` }}>
                    {result.headers.map((h, j) => (
                      <td key={h+j} style={{
                        padding: "5px 9px", whiteSpace: "nowrap", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis",
                        color: h === "Loyalty ID/Rewards Number" ? C.green : h === "Management Account Number (MAN)" ? C.orange : C.muted,
                        background: h === "Loyalty ID/Rewards Number" ? `${C.green}0a` : h === "Management Account Number (MAN)" ? `${C.orange}0a` : "transparent",
                        fontWeight: (h === "Loyalty ID/Rewards Number" || h === "Management Account Number (MAN)") ? 700 : 400,
                      }}>
                        {row[h] || <span style={{ color: C.faint }}>—</span>}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
