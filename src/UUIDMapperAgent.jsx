import { useState, useRef, useEffect, useMemo } from "react";

const ANTHROPIC_MODEL = "claude-sonnet-4-20250514";

// ── AGDC field definitions (output spec) ─────────────────────────────────────
const AGDC_OUTPUT_FIELDS = [
    { num: 1, key: "Management Account Number", label: "Management Account Number (MAN)", required: true },
    { num: 2, key: "Activity Date", label: "Activity Date", required: true },
    { num: 3, key: "Activity Time", label: "Activity Time", required: true },
    { num: 4, key: "Ad ID", label: "Ad ID", required: true },
    { num: 5, key: "Loyalty ID/Rewards Number", label: "Loyalty ID/Rewards Number", required: true, isLoyaltyId: true },
    { num: 6, key: "Activity Type", label: "Activity Type", required: true },
    { num: 7, key: "Channel Type", label: "Channel Type", required: true },
    { num: 8, key: "Marketing Transaction ID", label: "Marketing Transaction ID / Activity Log ID", required: false },
    { num: 9, key: "Activity State", label: "Activity State", required: false },
    { num: 10, key: "OPTIONAL EAIV Indicator", label: "OPTIONAL EAIV Indicator", required: false },
    { num: 11, key: "RESERVE FOR FUTURE USE 11", label: "RESERVE FOR FUTURE USE", required: false },
    { num: 12, key: "RESERVE FOR FUTURE USE 12", label: "RESERVE FOR FUTURE USE", required: false },
    { num: 13, key: "RESERVE FOR FUTURE USE 13", label: "RESERVE FOR FUTURE USE", required: false },
    { num: 14, key: "RESERVE FOR FUTURE USE 14", label: "RESERVE FOR FUTURE USE", required: false },
    { num: 15, key: "RESERVE FOR FUTURE USE 15", label: "RESERVE FOR FUTURE USE", required: false },
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

// ── XLSX parser (JSZip from cdnjs) ───────────────────────────────────────────
async function parseXLSX(arrayBuffer) {
    if (!window._JSZip) {
        await new Promise((res, rej) => {
            const s = document.createElement("script");
            s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
            s.onload = res; s.onerror = rej;
            document.head.appendChild(s);
        });
        window._JSZip = window.JSZip;
    }
    const zip = await window._JSZip.loadAsync(arrayBuffer);
    let sharedStrings = [];
    if (zip.files["xl/sharedStrings.xml"]) {
        const xml = await zip.files["xl/sharedStrings.xml"].async("text");
        sharedStrings = [...xml.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(m => m[1]);
    }
    const sheetEntry = zip.files["xl/worksheets/sheet1.xml"];
    if (!sheetEntry) throw new Error("No sheet1.xml found");
    const sheetXml = await sheetEntry.async("text");
    const rowMatches = [...sheetXml.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)];
    const colIdx = col => col.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    const grid = rowMatches.map(rm =>
        [...rm[1].matchAll(/<c r="([A-Z]+)\d+"[^>]*(?:t="([^"]*)")?[^>]*>[\s\S]*?<v>([^<]*)<\/v>/g)].map(c => ({
            col: c[1], value: c[2] === "s" ? (sharedStrings[parseInt(c[3])] ?? "") : c[3]
        }))
    );
    if (!grid.length) return { headers: [], rows: [] };
    const maxCol = Math.max(...grid.flat().map(c => colIdx(c.col)));
    const toArr = cells => {
        const arr = Array(maxCol + 1).fill("");
        cells.forEach(c => { arr[colIdx(c.col)] = c.value; });
        return arr;
    };
    const headers = toArr(grid[0]).map(h => String(h).trim());
    const rows = grid.slice(1).map(cells =>
        Object.fromEntries(headers.map((h, i) => [h, String(toArr(cells)[i] ?? "").trim()]))
    );
    return { headers, rows };
}

// ── CSV output ────────────────────────────────────────────────────────────────
function toCSV(headers, rows) {
    const esc = v => {
        const s = String(v ?? "");
        return s.includes(",") || s.includes('"') || s.includes("\n") || s.includes("\r")
            ? `"${s.replace(/"/g, '""')}"`
            : s;
    };
    return headers.join(",") + "\r\n" + rows.map(r => headers.map(h => esc(r[h])).join(",")).join("\r\n");
}

// ── normalization ─────────────────────────────────────────────────────────────
function normalizePhone(p) {
    const d = String(p ?? "").replace(/\D/g, "");
    return d.length === 11 && d.startsWith("1") ? d.slice(1) : d;
}
function normalizeEmail(e) { return String(e ?? "").trim().toLowerCase(); }
function normalizeKey(v, t) { return t === "email" ? normalizeEmail(v) : normalizePhone(v); }

// ── fuzzy column matcher — finds best DDM col for an AGDC field key ───────────
// Keywords let e.g. 'date' match 'SendDate', 'time' match 'SendTime', etc.
const FIELD_KEYWORDS = {
    "Activity Date": ["date", "activitydate", "senddate", "send_date", "actdate"],
    "Activity Time": ["time", "activitytime", "sendtime", "send_time", "acttime"],
    "Ad ID": ["adid", "ad_id", "campaignid", "campaign_id", "adcode"],
    "Activity Type": ["activitytype", "type", "eventtype", "event_type"],
    "Channel Type": ["channeltype", "channel", "medium", "messagetype"],
    "Marketing Transaction ID": ["transactionid", "transaction_id", "marketingid", "activitylogid"],
    "Activity State": ["status", "state", "activitystate", "deliverystatus"],
    "OPTIONAL EAIV Indicator": ["eaiv", "indicator", "eaivindicator"],
};

function fuzzyMatch(agdcKey, ddmHeaders) {
    const normalize = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
    const needle = normalize(agdcKey);
    const keywords = FIELD_KEYWORDS[agdcKey] ?? [];

    // 1. Exact normalized match
    let match = ddmHeaders.find(h => normalize(h) === needle);
    if (match) return match;

    // 2. Keyword prefix/exact match (highest specificity first)
    for (const kw of keywords) {
        match = ddmHeaders.find(h => normalize(h) === kw);
        if (match) return match;
    }

    // 3. Keyword contained in header (e.g. 'SendDate' contains 'date')
    for (const kw of keywords) {
        match = ddmHeaders.find(h => normalize(h).includes(kw) || kw.includes(normalize(h)));
        if (match) return match;
    }

    // 4. Partial match on the full AGDC key
    match = ddmHeaders.find(h => {
        const hay = normalize(h);
        return hay.includes(needle) || needle.includes(hay);
    });
    return match ?? null;
}

// ── component ─────────────────────────────────────────────────────────────────
export default function UUIDMapperAgent() {
    const [uuidFile, setUuidFile] = useState(null);
    const [ddmFile, setDdmFile] = useState(null);
    const [matchType, setMatchType] = useState("phone");
    const [refIdCol, setRefIdCol] = useState("");
    const [refUuidCol, setRefUuidCol] = useState("");
    const [ddmIdCol, setDdmIdCol] = useState("");
    const [manAcctNum, setManAcctNum] = useState("");  // Field 1 override
    const [retailerName, setRetailerName] = useState(""); // for filename
    const [selectedMonth, setSelectedMonth] = useState(String(new Date().getMonth() + 1).padStart(2, "0"));
    const [selectedYear, setSelectedYear] = useState(String(new Date().getFullYear()));
    const [colMap, setColMap] = useState({});  // agdcKey → ddmCol for fields 2–10
    const [messages, setMessages] = useState([]);
    const [input, setInput] = useState("");
    const [loading, setLoading] = useState(false);
    const [result, setResult] = useState(null);
    const [showCopy, setShowCopy] = useState(false);
    const [showColMap, setShowColMap] = useState(false);
    const chatRef = useRef(null);
    const copyRef = useRef(null);

    const addMsg = (role, text) =>
        setMessages(prev => [...prev, { role, text, ts: Date.now() }]);

    // ── file reading ─────────────────────────────────────────────────────────────
    const readFile = (file, setter, label) => {
        const isXLSX = /\.xlsx?$/i.test(file.name);
        const finish = ({ headers, rows }) => {
            setter({ name: file.name, headers, rows });
            // auto-build colMap when DDM loads
            if (label === "DDM") {
                const auto = {};
                AGDC_OUTPUT_FIELDS.forEach(f => {
                    if (f.num === 1 || f.isLoyaltyId || f.num >= 11) return;
                    const found = fuzzyMatch(f.key, headers);
                    if (found) auto[f.key] = found;
                });
                setColMap(auto);
            }
            addMsg("agent", `✅ Loaded ${label}: ${file.name} — ${rows.length} rows\nColumns: ${headers.join(", ")}`);
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
            // PROXY PATH: /anthropic-api is mapped to https://api.anthropic.com in vite.config.js
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
            if (data.error) {
                addMsg("agent", `⚠️ API Error: ${data.error.message}`);
            } else {
                addMsg("agent", data.content?.find(b => b.type === "text")?.text ?? "No response.");
            }
        } catch (err) { addMsg("agent", `⚠️ Error: ${err.message}`); }
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
                    messages: [{
                        role: "user", content:
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
    "Activity State": "matching DDM column or null",
    "OPTIONAL EAIV Indicator": "matching DDM column or null"
  }
}`
                    }],
                }),
            });
            const data = await res.json();
            if (data.error) throw new Error(data.error.message);

            const raw = data.content?.find(b => b.type === "text")?.text ?? "";
            const d = JSON.parse(raw.replace(/```json|```/g, "").trim());
            if (d.refIdCol) setRefIdCol(d.refIdCol);
            if (d.refUuidCol) setRefUuidCol(d.refUuidCol);
            if (d.ddmIdCol) setDdmIdCol(d.ddmIdCol);
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
                `- Field mappings: ${Object.entries(d.colMap || {}).filter(([, v]) => v && v !== "null").map(([k, v]) => `${k} → ${v}`).join(", ") || "none detected"}\n\n` +
                `Review and click ▶ Run Mapping.`
            );
        } catch (e) { addMsg("agent", `⚠️ Auto-detect failed: ${e.message}`); }
        setLoading(false);
    };

    // ── mapping ───────────────────────────────────────────────────────────────────
    const runMapping = () => {
        if (!uuidFile || !ddmFile || !refIdCol || !refUuidCol || !ddmIdCol) {
            addMsg("agent", "⚠️ Load both files and select all identifier columns first.");
            return;
        }

        // build UUID lookup
        const lookup = new Map();
        for (const row of uuidFile.rows) {
            const key = normalizeKey(row[refIdCol], matchType);
            const uuid = (row[refUuidCol] ?? "").trim();
            if (key && uuid) lookup.set(key, uuid);
        }

        const sampleRef = [...lookup.keys()].slice(0, 3);
        const sampleDDM = ddmFile.rows.slice(0, 3).map(r => normalizeKey(r[ddmIdCol], matchType));
        let matched = 0, unmatched = 0;

        const newRows = ddmFile.rows.map(row => {
            const uuid = lookup.get(normalizeKey(row[ddmIdCol], matchType));
            uuid ? matched++ : unmatched++;

            // Build output row with all 15 AGDC fields in order
            const out = {};
            AGDC_OUTPUT_FIELDS.forEach(f => {
                if (f.num === 1) {
                    // Field 1: user-supplied Management Account Number
                    out[f.label] = manAcctNum.trim();
                } else if (f.isLoyaltyId) {
                    // Field 5: UUID
                    out[f.label] = uuid ?? "";
                } else if (f.num >= 11) {
                    // Fields 11–15: always blank
                    out[f.label] = "";
                } else {
                    // Fields 2–4, 6–10: map from DDM via colMap; with field-specific defaults
                    const ddmCol = colMap[f.key];
                    const raw = ddmCol ? (row[ddmCol] ?? "") : "";
                    // Activity Time defaults to 0:00:00 if empty or unmapped
                    out[f.label] = (f.key === "Activity Time" && !raw.trim()) ? "0:00:00" : raw;
                }
            });
            return out;
        });

        setResult({ headers: OUTPUT_HEADERS, rows: newRows, stats: { matched, unmatched, total: ddmFile.rows.length } });
        setShowCopy(false);
        addMsg("agent",
            `✅ Done! Output has all 15 AGDC fields in order.\n` +
            `- ${matched} rows matched → Field 5 = UUID\n` +
            `- ${unmatched} rows unmatched → Field 5 blank\n` +
            `- Total: ${ddmFile.rows.length} rows\n\n` +
            `🔍 Ref sample ${matchType}s: ${sampleRef.join(", ") || "none"}\n` +
            `🔍 DDM sample ${matchType}s: ${sampleDDM.join(", ") || "none"}` +
            (matched === 0 ? `\n\n⚠️ 0 matches — paste a sample value from each file so I can diagnose the format difference.` : "")
        );
    };

    // ── download / copy ───────────────────────────────────────────────────────────
    const csvContent = useMemo(
        () => (result ? toCSV(result.headers, result.rows) : ""),
        [result]
    );

    const sanitize = s => String(s ?? "").trim().replace(/[^a-z0-9_-]/gi, "_");
    const downloadFilename = `${sanitize(retailerName) || "Retailer"}_${sanitize(manAcctNum) || "00000"}_CampaignActivityLog_${selectedMonth}${selectedYear}.csv`;

    const handleDownload = async () => {
        if (!result || !csvContent) return;

        // Primary: File System Access API — shows native Save As dialog w/ correct filename
        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: downloadFilename,
                    types: [{ description: "CSV File", accept: { "text/csv": [".csv"] } }],
                });
                const writable = await handle.createWritable();
                await writable.write(csvContent);
                await writable.close();
                addMsg("agent", `✅ Saved: ${downloadFilename}\n(${result.rows.length.toLocaleString()} rows)`);
                return;
            } catch (e) {
                if (e.name === "AbortError") return; // user cancelled the picker
                // fall through to blob fallback
            }
        }

        // Fallback: blob URL with data-uri trick for reliable naming
        try {
            const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8" });
            const url = URL.createObjectURL(blob);
            const a = Object.assign(document.createElement("a"), {
                href: url,
                download: downloadFilename,
            });
            document.body.appendChild(a);
            a.click();
            setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 60000);
            addMsg("agent", `🚀 Download initiated: ${downloadFilename}`);
        } catch (err) {
            addMsg("agent", `⚠️ Download failed: ${err.message}`);
            setShowCopy(true);
        }
    };

    const handleCopy = () => {
        try {
            navigator.clipboard.writeText(csvContent).then(() => addMsg("agent", "✅ CSV copied to clipboard!"));
        } catch {
            if (copyRef.current) { copyRef.current.select(); document.execCommand("copy"); addMsg("agent", "✅ Copied!"); }
        }
    };

    // ── ui helpers ────────────────────────────────────────────────────────────────
    const Sel = ({ label, value, onChange, options, dim }) => (
        <div style={{ marginBottom: 8 }}>
            <label style={{ display: "block", fontSize: 10, fontWeight: 700, letterSpacing: 1, color: dim ? "#334155" : "#7dd3fc", marginBottom: 3, textTransform: "uppercase" }}>{label}</label>
            <select value={value} onChange={e => onChange(e.target.value)}
                style={{ width: "100%", background: "#0f172a", border: `1px solid ${dim ? "#1a2a3a" : "#1e3a5f"}`, color: value ? "#e2e8f0" : "#475569", borderRadius: 5, padding: "5px 8px", fontSize: 11 }}>
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
                style={{ border: `2px dashed ${file ? "#38bdf8" : "#1e3a5f"}`, borderRadius: 9, padding: "13px 10px", textAlign: "center", cursor: "pointer", background: file ? "#0c1f36" : "#060f1c", marginBottom: 8 }}>
                <input ref={ref} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (f) onFile(f); }} />
                <div style={{ fontSize: 18, marginBottom: 2 }}>{file ? "📄" : "📂"}</div>
                <div style={{ fontSize: 11, fontWeight: 700, color: file ? "#38bdf8" : "#475569" }}>{label}</div>
                {file
                    ? <div style={{ fontSize: 10, color: "#94a3b8", marginTop: 2 }}>{file.name} · {file.rows.length} rows</div>
                    : <div style={{ fontSize: 10, color: "#334155", marginTop: 2 }}>Click or drop · CSV / XLSX</div>}
            </div>
        );
    };

    const statsBar = result?.stats && (
        <div style={{ display: "flex", gap: 5, margin: "7px 0" }}>
            {[["Matched", result.stats.matched, "#22c55e"], ["Unmatched", result.stats.unmatched, "#f59e0b"], ["Total", result.stats.total, "#38bdf8"]].map(([l, v, c]) => (
                <div key={l} style={{ flex: 1, background: "#0f172a", border: `1px solid ${c}33`, borderRadius: 6, padding: "6px 0", textAlign: "center" }}>
                    <div style={{ fontSize: 16, fontWeight: 800, color: c }}>{v}</div>
                    <div style={{ fontSize: 9, color: "#64748b", letterSpacing: 1, textTransform: "uppercase" }}>{l}</div>
                </div>
            ))}
        </div>
    );

    return (
        <div style={{ minHeight: "100vh", background: "#020b18", fontFamily: "'JetBrains Mono','Fira Code',monospace", color: "#e2e8f0", display: "flex", flexDirection: "column", alignItems: "center", padding: "20px 10px" }}>

            {/* header */}
            <div style={{ width: "100%", maxWidth: 820, marginBottom: 14 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <div style={{ width: 34, height: 34, borderRadius: 8, background: "linear-gradient(135deg,#0ea5e9,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>🔀</div>
                    <div>
                        <div style={{ fontSize: 15, fontWeight: 800, color: "#f1f5f9" }}>UUID Mapper Agent</div>
                        <div style={{ fontSize: 10, color: "#475569", letterSpacing: .5 }}>PAR RETAIL · AGDC DDM · 15-FIELD OUTPUT</div>
                    </div>
                </div>
            </div>

            <div style={{ width: "100%", maxWidth: 820, display: "flex", gap: 10 }}>

                {/* ── left panel ── */}
                <div style={{ width: 232, flexShrink: 0, display: "flex", flexDirection: "column", gap: 8 }}>

                    {/* 1 load */}
                    <div style={{ background: "#070f1d", border: "1px solid #0f2744", borderRadius: 10, padding: 11 }}>
                        <div style={{ fontSize: 10, fontWeight: 700, color: "#475569", letterSpacing: 1, marginBottom: 7, textTransform: "uppercase" }}>1 · Load Files</div>
                        <DropZone label="UUID Reference File" file={uuidFile} onFile={f => readFile(f, setUuidFile, "UUID Reference")} />
                        <DropZone label="DDM File" file={ddmFile} onFile={f => readFile(f, setDdmFile, "DDM")} />
                    </div>

                    {/* 2 field 1 + match key */}
                    <div style={{ background: "#070f1d", border: "1px solid #0f2744", borderRadius: 10, padding: 11 }}>
                        <div style={{ fontSize: 10, fontWeight: 700, color: "#475569", letterSpacing: 1, marginBottom: 7, textTransform: "uppercase" }}>2 · MAN + Retailer Name</div>
                        <div style={{ marginBottom: 6 }}>
                            <label style={{ display: "block", fontSize: 10, fontWeight: 700, letterSpacing: 1, color: "#7dd3fc", marginBottom: 3, textTransform: "uppercase" }}>Retailer Name</label>
                            <input value={retailerName} onChange={e => setRetailerName(e.target.value)}
                                placeholder="e.g. QuickGas"
                                style={{ width: "100%", background: "#0f172a", border: "1px solid #1e3a5f", borderRadius: 5, color: "#e2e8f0", padding: "6px 8px", fontSize: 12, fontFamily: "inherit", outline: "none", boxSizing: "border-box", marginBottom: 6 }} />
                            <label style={{ display: "block", fontSize: 10, fontWeight: 700, letterSpacing: 1, color: "#7dd3fc", marginBottom: 3, textTransform: "uppercase" }}>MAN (Managed Account Number)</label>
                            <input value={manAcctNum} onChange={e => setManAcctNum(e.target.value)}
                                placeholder="e.g. 012345"
                                style={{ width: "100%", background: "#0f172a", border: "1px solid #1e3a5f", borderRadius: 5, color: "#e2e8f0", padding: "6px 8px", fontSize: 12, fontFamily: "inherit", outline: "none", boxSizing: "border-box" }} />
                            <div style={{ fontSize: 9, color: "#334155", marginTop: 4, lineHeight: 1.5 }}>
                                Filename: <span style={{ color: "#475569" }}>{(retailerName.trim().replace(/\s+/g, "") || "Retailer")}_{manAcctNum.trim() || "00000"}_CampaignActivityLog_{selectedMonth}{selectedYear}.csv</span>
                            </div>
                            <div style={{ display: "flex", gap: 5, marginTop: 8 }}>
                                <div style={{ flex: 1 }}>
                                    <label style={{ display: "block", fontSize: 10, fontWeight: 700, letterSpacing: 1, color: "#7dd3fc", marginBottom: 3, textTransform: "uppercase" }}>Month</label>
                                    <select value={selectedMonth} onChange={e => setSelectedMonth(e.target.value)}
                                        style={{ width: "100%", background: "#0f172a", border: "1px solid #1e3a5f", color: "#e2e8f0", borderRadius: 5, padding: "5px 6px", fontSize: 11, fontFamily: "inherit" }}>
                                        {["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"].map((m, i) => (
                                            <option key={m} value={m}>{["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][i]} ({m})</option>
                                        ))}
                                    </select>
                                </div>
                                <div style={{ flex: 1 }}>
                                    <label style={{ display: "block", fontSize: 10, fontWeight: 700, letterSpacing: 1, color: "#7dd3fc", marginBottom: 3, textTransform: "uppercase" }}>Year</label>
                                    <select value={selectedYear} onChange={e => setSelectedYear(e.target.value)}
                                        style={{ width: "100%", background: "#0f172a", border: "1px solid #1e3a5f", color: "#e2e8f0", borderRadius: 5, padding: "5px 6px", fontSize: 11, fontFamily: "inherit" }}>
                                        {[2023, 2024, 2025, 2026, 2027].map(y => (
                                            <option key={y} value={String(y)}>{y}</option>
                                        ))}
                                    </select>
                                </div>
                            </div>
                        </div>
                    </div>

                    {/* 3 map columns */}
                    {uuidFile && ddmFile && (
                        <div style={{ background: "#070f1d", border: "1px solid #0f2744", borderRadius: 10, padding: 11 }}>
                            <div style={{ fontSize: 10, fontWeight: 700, color: "#475569", letterSpacing: 1, marginBottom: 7, textTransform: "uppercase" }}>3 · Map Identifier Column</div>
                            <div style={{ marginBottom: 8 }}>
                                <label style={{ display: "block", fontSize: 10, fontWeight: 700, letterSpacing: 1, color: "#7dd3fc", marginBottom: 4, textTransform: "uppercase" }}>Match Key</label>
                                <div style={{ display: "flex", gap: 4 }}>
                                    {["phone", "email"].map(t => (
                                        <button key={t} onClick={() => { setMatchType(t); setRefIdCol(""); setDdmIdCol(""); }}
                                            style={{
                                                flex: 1, padding: "5px 0", borderRadius: 5, fontSize: 10, fontWeight: 700, fontFamily: "inherit", cursor: "pointer", textTransform: "uppercase",
                                                background: matchType === t ? "linear-gradient(135deg,#0ea5e9,#6366f1)" : "#0f172a",
                                                color: matchType === t ? "#fff" : "#475569",
                                                border: matchType === t ? "none" : "1px solid #1e3a5f"
                                            }}>
                                            {t === "phone" ? "📱 Phone" : "✉️ Email"}
                                        </button>
                                    ))}
                                </div>
                            </div>
                            <Sel label={`Ref → ${matchType}`} value={refIdCol} onChange={setRefIdCol} options={uuidFile.headers} />
                            <Sel label="Ref → UUID col" value={refUuidCol} onChange={setRefUuidCol} options={uuidFile.headers} />
                            <Sel label={`DDM → ${matchType}`} value={ddmIdCol} onChange={setDdmIdCol} options={ddmFile.headers} />

                            {/* optional: field mapping overrides */}
                            <button onClick={() => setShowColMap(v => !v)}
                                style={{ width: "100%", background: "#0a1628", color: "#475569", border: "1px solid #0f2744", borderRadius: 5, padding: "5px 0", fontSize: 10, cursor: "pointer", marginBottom: 7, fontFamily: "inherit" }}>
                                {showColMap ? "▲ Hide" : "▼ Show"} field column mapping (fields 2–10)
                            </button>
                            {showColMap && (
                                <div style={{ background: "#040b14", borderRadius: 6, padding: "8px 8px 2px", marginBottom: 7, border: "1px solid #0f2744" }}>
                                    <div style={{ fontSize: 9, color: "#334155", marginBottom: 6, letterSpacing: .5 }}>MAP DDM COLUMNS → AGDC FIELDS</div>
                                    {AGDC_OUTPUT_FIELDS.filter(f => f.num >= 2 && f.num <= 10 && !f.isLoyaltyId).map(f => (
                                        <Sel key={f.key} dim label={`F${f.num}: ${f.key}`}
                                            value={colMap[f.key] ?? ""}
                                            onChange={v => setColMap(prev => ({ ...prev, [f.key]: v }))}
                                            options={ddmFile.headers} />
                                    ))}
                                </div>
                            )}

                            <button onClick={autoDetect} disabled={loading}
                                style={{ width: "100%", background: "#0f2744", color: "#7dd3fc", border: "1px solid #1e3a5f", borderRadius: 5, padding: "6px 0", fontSize: 11, cursor: "pointer", marginBottom: 5, fontFamily: "inherit" }}>
                                🤖 Auto-detect all columns
                            </button>
                            <button onClick={runMapping} disabled={!refIdCol || !refUuidCol || !ddmIdCol}
                                style={{ width: "100%", background: refIdCol && refUuidCol && ddmIdCol ? "linear-gradient(135deg,#0ea5e9,#6366f1)" : "#1e293b", color: "#fff", border: "none", borderRadius: 5, padding: "8px 0", fontSize: 12, fontWeight: 700, cursor: refIdCol && refUuidCol && ddmIdCol ? "pointer" : "not-allowed", fontFamily: "inherit" }}>
                                ▶ Run Mapping
                            </button>
                        </div>
                    )}

                    {/* 4 export */}
                    {result && (
                        <div style={{ background: "#070f1d", border: "1px solid #166534", borderRadius: 10, padding: 11 }}>
                            <div style={{ fontSize: 10, fontWeight: 700, color: "#475569", letterSpacing: 1, marginBottom: 5, textTransform: "uppercase" }}>4 · Export</div>
                            {statsBar}
                            <button onClick={handleDownload}
                                style={{ width: "100%", background: "linear-gradient(135deg,#16a34a,#15803d)", color: "#fff", border: "none", borderRadius: 5, padding: "7px 0", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", marginBottom: 5 }}>
                                ⬇ Download CSV
                            </button>
                            <button onClick={() => setShowCopy(v => !v)}
                                style={{ width: "100%", background: "#0f2744", color: "#7dd3fc", border: "1px solid #1e3a5f", borderRadius: 5, padding: "5px 0", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>
                                📋 {showCopy ? "Hide" : "Copy"} CSV Text
                            </button>
                        </div>
                    )}
                </div>

                {/* ── chat panel ── */}
                <div style={{ flex: 1, display: "flex", flexDirection: "column", background: "#070f1d", border: "1px solid #0f2744", borderRadius: 10, overflow: "hidden", minHeight: 500 }}>
                    <div style={{ padding: "9px 13px", borderBottom: "1px solid #0f2744", display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#22c55e", boxShadow: "0 0 4px #22c55e" }} />
                        <span style={{ fontSize: 10, color: "#64748b", letterSpacing: .5, fontWeight: 700 }}>MAPPING COPILOT</span>
                    </div>
                    <div ref={chatRef} style={{ flex: 1, overflowY: "auto", padding: 11, display: "flex", flexDirection: "column", gap: 7 }}>
                        {messages.length === 0 && (
                            <div style={{ color: "#334155", fontSize: 12, textAlign: "center", marginTop: 40 }}>
                                <div style={{ fontSize: 26, marginBottom: 8 }}>🤖</div>
                                Load your files to get started.<br />
                                <span style={{ fontSize: 10, color: "#1e3a5f" }}>CSV & XLSX · Phone or email · 15-field AGDC output</span>
                            </div>
                        )}
                        {messages.map(m => (
                            <div key={m.ts} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start" }}>
                                <div style={{
                                    maxWidth: "88%", borderRadius: 8, padding: "7px 10px", fontSize: 12, lineHeight: 1.6, whiteSpace: "pre-wrap",
                                    background: m.role === "user" ? "linear-gradient(135deg,#1e40af,#4f46e5)" : "#0f1f36",
                                    border: m.role === "agent" ? "1px solid #1e3a5f" : "none", color: "#e2e8f0"
                                }}>
                                    {m.text}
                                </div>
                            </div>
                        ))}
                        {loading && (
                            <div style={{ display: "flex" }}>
                                <div style={{ background: "#0f1f36", border: "1px solid #1e3a5f", borderRadius: 8, padding: "7px 10px", fontSize: 11, color: "#64748b" }}>⏳ Thinking...</div>
                            </div>
                        )}
                    </div>
                    <div style={{ padding: "7px 9px", borderTop: "1px solid #0f2744", display: "flex", gap: 5 }}>
                        <input value={input} onChange={e => setInput(e.target.value)}
                            onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendMessage(input); } }}
                            placeholder="Ask about columns, formats, or the mapping..."
                            style={{ flex: 1, background: "#0a1628", border: "1px solid #1e3a5f", borderRadius: 5, color: "#e2e8f0", padding: "6px 9px", fontSize: 11, fontFamily: "inherit", outline: "none" }} />
                        <button onClick={() => sendMessage(input)} disabled={loading || !input.trim()}
                            style={{ background: "linear-gradient(135deg,#0ea5e9,#6366f1)", border: "none", borderRadius: 5, color: "#fff", padding: "0 12px", fontSize: 13, cursor: loading ? "not-allowed" : "pointer" }}>↑</button>
                    </div>
                </div>
            </div>

            {/* copy textarea */}
            {showCopy && result && (
                <div style={{ width: "100%", maxWidth: 820, marginTop: 10, background: "#070f1d", border: "1px solid #1e3a5f", borderRadius: 10, overflow: "hidden" }}>
                    <div style={{ padding: "7px 12px", borderBottom: "1px solid #1e3a5f", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <span style={{ fontSize: 10, color: "#7dd3fc", fontWeight: 700, letterSpacing: 1 }}>CSV OUTPUT</span>
                        <button onClick={handleCopy} style={{ background: "#0ea5e9", color: "#fff", border: "none", borderRadius: 4, padding: "3px 10px", fontSize: 10, cursor: "pointer", fontFamily: "inherit", fontWeight: 700 }}>Copy All</button>
                    </div>
                    <textarea ref={copyRef} readOnly value={csvContent}
                        style={{ width: "100%", height: 150, background: "#020b18", color: "#94a3b8", border: "none", padding: 10, fontSize: 11, fontFamily: "inherit", resize: "vertical", outline: "none", boxSizing: "border-box" }} />
                </div>
            )}

            {/* preview */}
            {result && (
                <div style={{ width: "100%", maxWidth: 820, marginTop: 10, background: "#070f1d", border: "1px solid #166534", borderRadius: 10, overflow: "hidden" }}>
                    <div style={{ padding: "7px 12px", borderBottom: "1px solid #166534", fontSize: 10, color: "#22c55e", fontWeight: 700, letterSpacing: 1 }}>
                        OUTPUT PREVIEW — first 5 rows · all 15 AGDC fields
                    </div>
                    <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
                            <thead>
                                <tr>{result.headers.map((h, i) => (
                                    <th key={h + i} style={{
                                        padding: "6px 9px", textAlign: "left", whiteSpace: "nowrap", borderBottom: "1px solid #0f2744",
                                        color: h === "Loyalty ID/Rewards Number" ? "#22c55e" : h === "Management Account Number (MAN)" ? "#f59e0b" : "#38bdf8",
                                        background: h === "Loyalty ID/Rewards Number" ? "#0c2a1a" : h === "Management Account Number (MAN)" ? "#1a1200" : "transparent"
                                    }}>{h}</th>
                                ))}</tr>
                            </thead>
                            <tbody>
                                {result.rows.slice(0, 5).map((row, i) => (
                                    <tr key={i} style={{ borderBottom: "1px solid #0a1628" }}>
                                        {result.headers.map((h, j) => (
                                            <td key={h + j} style={{
                                                padding: "5px 9px", whiteSpace: "nowrap", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis",
                                                color: h === "Loyalty ID/Rewards Number" ? "#22c55e" : h === "Management Account Number (MAN)" ? "#f59e0b" : "#94a3b8",
                                                background: h === "Loyalty ID/Rewards Number" ? "#0a1e0f" : h === "Management Account Number (MAN)" ? "#120e00" : "transparent",
                                                fontWeight: h === "Loyalty ID/Rewards Number" || h === "Management Account Number (MAN)" ? 700 : 400
                                            }}>
                                                {row[h] || <span style={{ color: "#374151" }}>—</span>}
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
