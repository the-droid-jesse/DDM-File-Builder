# DDM File Builder

> **PAR Retail · AGDC Digital Trade Program Activity Log Builder**

A local React tool that maps loyalty UUIDs into DDM files and exports a properly formatted 15-field AGDC Campaign Activity Log CSV — with an AI-powered column detection copilot.

---

## Features

- **UUID Mapping** — Match customers by **phone** or **email** across a UUID reference file and a DDM file, populating the AGDC Loyalty ID/Rewards Number field (Field 5)
- **Optional UUID Mapping** — UUID lookup can be skipped or dynamically overridden with fallback DDM columns
- **Fallback Field Handling** — Tools to inject default values for missing data (e.g. Field 9 Activity State)
- **15-field AGDC output** — Output strictly conforms to the AGDC Digital Trade Program specification, fields in order
- **CSV & XLSX support** — Drag-and-drop or click to upload both file types (powered by SheetJS for robust XLSX parsing)
- **Smart column detection** — Fuzzy keyword matcher auto-maps DDM columns to AGDC fields on upload, with AI-assisted override
- **AI Mapping Copilot** — Powered by Claude (Anthropic) via a local Vite proxy; helps you diagnose match failures and column mappings
- **Named file download** — Uses the native File System Access API (`showSaveFilePicker`) to save files with the correct name: `RetailerName_MAN_CampaignActivityLog_MMYYYY.csv`
- **Handles large files** — Tested with 390,000+ row DDM files (~3 MB output)

---

## AGDC Output Fields

| # | Field | Source |
|---|-------|--------|
| 1 | Management Account Number (MAN) | Entered manually by user |
| 2 | Activity Date | Mapped from DDM file |
| 3 | Activity Time | Mapped from DDM file (defaults to `0:00:00`) |
| 4 | Ad ID | Mapped from DDM file |
| 5 | Loyalty ID/Rewards Number | **UUID looked up from reference file (or mapped natively)** |
| 6 | Activity Type | Mapped from DDM file |
| 7 | Channel Type | Mapped from DDM file |
| 8 | Marketing Transaction ID | Mapped from DDM file |
| 9 | Activity State | Mapped from DDM file |
| 10 | OPTIONAL EAIV Indicator | Mapped from DDM file |
| 11–15 | RESERVE FOR FUTURE USE | Always blank |

---

## Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) v18+
- An [Anthropic API key](https://console.anthropic.com/) *(required for AI auto-detect only)*

### Installation

```bash
git clone git@github.com:the-droid-jesse/DDM-File-Builder.git
cd DDM-File-Builder
npm install
```

### Configuration

Copy `.env.example` to `.env` and add your Anthropic API key:

```bash
cp .env.example .env
```

Edit `.env`:
```
VITE_ANTHROPIC_API_KEY=sk-ant-api03-your-key-here
```

> ⚠️ **Never commit `.env`** — it is in `.gitignore` by default.

### Running Locally

```bash
npm run dev
```

Open [http://localhost:5173](http://localhost:5173) in Chrome.

---

## Usage

1. **Load Files** — Drop your UUID reference file and DDM file (CSV or XLSX)
2. **Set MAN + Retailer** — Enter the Management Account Number and retailer name (used for the output filename)
3. **Map Columns** — Select the month/year, then map the identifier columns (phone or email). Use **🤖 Auto-detect** for AI-assisted mapping
4. **Run Mapping** — Click **▶ Run Mapping** to process all rows. Missing UUIDs will natively fallback to mapped DDM columns if configured.
5. **Export** — Click **⬇ Download CSV** to save the named output file via the OS Save dialog

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| UI | React 19 + Vite |
| Styling | Inline styles (no deps) |
| XLSX parsing | SheetJS / xlsx (loaded from CDN on demand) |
| AI | Anthropic Claude (via Vite dev proxy) |
| File download | File System Access API (`showSaveFilePicker`) |

---

## Project Structure

```
uuid-mapper-app/
├── src/
│   ├── UUIDMapperAgent.jsx   # Main component (all logic + UI)
│   ├── App.jsx               # Root render
│   └── index.css             # Global resets
├── .env.example              # Required env vars template
├── vite.config.js            # Dev server + Anthropic proxy
└── package.json
```

---

## Notes

- The Anthropic API proxy runs **server-side** in Vite's dev server — your API key is never exposed to the browser
- For production deployment, you would need a backend proxy (e.g. Vercel Edge Function, Express) to replace the Vite dev proxy
- Files up to ~50k rows have been tested successfully

---

## License

Internal tool — PAR Retail. Not for public distribution.
