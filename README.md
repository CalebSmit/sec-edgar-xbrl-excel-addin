# SEC EDGAR XBRL Financial Statements — Excel Add-in

**Version:** 1.0.0 | **Platform:** Excel for Windows | **Cost:** Free

Pull live financial statement data from the SEC EDGAR database directly into Excel.
No sign-up, no subscriptions, no backend — data comes straight from the SEC.

---

## Download & Install

### Step 1 — Download the add-in

**[View the latest release →](https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin/releases/tag/v1.0.0)**

On the release page, look for `SEC_XBRL_Addin.xlam` in the **Assets** section and click it to download.

Alternatively, you can download the file directly from the repository:
1. Go to the [dist folder](https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin/tree/master/dist)
2. Click `SEC_XBRL_Addin.xlam`
3. Click the **Download raw file** button (the download icon in the top-right of the file view)

> Save `SEC_XBRL_Addin.xlam` to a folder on your computer (e.g. `Documents\ExcelAddins\`).  
> Do **not** open it directly from your browser — save it first, then install from Excel.

---

### Step 2 — Install in Excel

1. Open Excel
2. Click **File → Options**
3. Click **Add-ins** in the left sidebar
4. At the bottom, make sure **"Excel Add-ins"** is selected → click **Go...**
5. Click **Browse...**
6. Navigate to where you saved `SEC_XBRL_Addin.xlam` → select it → click **OK**
7. Make sure **SEC_XBRL_Addin** is checked → click **OK**

---

### Step 3 — Enable macros (if prompted)

If Excel shows a yellow bar saying "Macros have been disabled" → click **Enable Content**.

If Excel blocks macros entirely:
1. **File → Options → Trust Center → Trust Center Settings**
2. Click **Macro Settings**
3. Select **"Disable all macros with notification"** → OK

---

### Step 4 — Use it

| Method | How |
|--------|-----|
| **Keyboard** | Press `Ctrl + Shift + S` → type a ticker → Enter |
| **Macro** | `Alt + F8` → select `PullSECFinancials` → Run → type a ticker |
| **Ribbon button** | Click **Pull SEC Financials** in the **SEC EDGAR** tab *(requires extra setup — see below)* |

Enter any US stock ticker: `AAPL`, `MSFT`, `TSLA`, `JPM`, `GOOGL`, `BRK-B`, etc.

The add-in fetches data from SEC EDGAR and fills three worksheets:
- **Income Statement** — Revenue, gross profit, net income, EPS, R&D, SG&A
- **Balance Sheet** — Assets, liabilities, equity, cash, goodwill, debt
- **Cash Flow** — Operating/investing/financing flows, capex, D&A

Both **annual (10-K)** and **quarterly (10-Q)** data appear side by side.

---

## Requirements

- Windows 10 or 11
- Microsoft Excel 2016 or newer (Microsoft 365 works too)
- Internet connection (to fetch SEC data)

> **Mac not supported.** The add-in uses Windows-only COM objects.

---

## What the data looks like

Each cell traces 1:1 to the `val` field in the raw SEC EDGAR JSON — no estimates, no rounding, no LLM inference. Column headers are ISO end-dates (e.g. `2024-09-28`). Row labels are the XBRL concept names as filed.

Large companies like Apple (~15 MB) or JPMorgan (~25 MB) may take 15–30 seconds to load. This is normal — the SEC serves the full filing history in one file.

---

## Error Messages

| Message | What it means |
|---------|---------------|
| `Ticker 'XYZ' not found in SEC database.` | Check spelling. Use the SEC ticker (e.g. `BRK-B` not `BRKB`). |
| `SEC rate-limited. Please wait 30 seconds and try again.` | Too many requests. Wait and retry. |
| `Cannot connect to SEC servers. Check your internet connection.` | No internet, or SEC is temporarily down. |
| `No US-GAAP XBRL data found for this company.` | Company doesn't file US-GAAP XBRL (e.g. foreign filers). |
| `Failed to parse SEC response. The data format may have changed.` | SEC changed their API format. Check GitHub for an update. |

---

## Known Limitations

| Limitation | Detail |
|------------|--------|
| **Windows only** | Mac Excel not supported in v1. |
| **Q4 not in quarterly** | Q4 is not filed in 10-Q. It appears only in the Annual section. |
| **10-K/A excluded** | Amended filings are excluded; original 10-K values are shown. |
| **XBRL tag names** | Row labels are raw XBRL names, not human-friendly labels. |
| **One unit per concept** | Prefers USD > USD/shares > shares > pure. |
| **No IFRS** | US-GAAP only. Foreign IFRS filers will show E4. |

---

## Uninstall

**File → Options → Add-ins → Go...** → uncheck **SEC_XBRL_Addin** → OK.

---
---

## For Developers — Build from Source

> Normal users do not need this section. Install the `.xlam` from the `dist/` folder above.

### Repository Structure

```
sec-edgar-xbrl-excel-addin/
├── dist/
│   ├── SEC_XBRL_Addin.xlam   ← Prebuilt add-in (download this)
│   └── INSTALL.txt            ← Beginner install guide
├── modules/                   ← VBA source (10 .bas / .cls files)
├── dependencies/
│   ├── JsonConverter.bas      ← VBA-JSON v2.3.1 (MIT license, bundled)
│   └── LICENSE-VBA-JSON.txt
├── customUI/
│   └── customUI14.xml         ← Ribbon XML (Office 2010+)
└── README.md
```

### Building the .xlam from Source

If you want to build the add-in yourself from VBA source:

1. Open Excel and press `Alt + F11` to open the VBA editor
2. **File → Import File** → import `dependencies/JsonConverter.bas`
3. **File → Import File** → import all 9 files from `modules/`:
   - `modConfig.bas`, `modHTTP.bas`, `modTickerLookup.bas`
   - `modJSONParser.bas`, `modClassifier.bas`, `modExcelWriter.bas`
   - `modProgress.bas`, `modRibbon.bas`, `modMain.bas`
4. Import `modules/ThisWorkbook.cls` into the **existing ThisWorkbook** module (do not create a new one)
5. **Tools → References** → check:
   - ✅ Microsoft Scripting Runtime
   - ✅ Microsoft XML, v6.0
6. *(Optional)* Use [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) to inject `customUI/customUI14.xml` for the ribbon button
7. **File → Save As** → type: **Excel Add-in (.xlam)** → save as `SEC_XBRL_Addin.xlam`

### Architecture

```
SEC_XBRL_Addin.xlam
├── modConfig        — Constants, API endpoints, keyword lists, error codes
├── modHTTP          — WinHttp GET, User-Agent header, rate-limit, retries
├── modTickerLookup  — Ticker → 10-digit CIK via SEC company_tickers.json
├── modJSONParser    — JSON navigation, annual/quarterly filters, dedup
├── modClassifier    — Keyword-based XBRL concept → IS/BS/CFS bucketing
├── modExcelWriter   — Worksheet creation, sparse-matrix layout, CDbl() writes
├── modProgress      — Progress bar, error dialogs, PROG_* / ERR_* constants
├── modRibbon        — RibbonCallback_PullSECFinancials, Ctrl+Shift+S shortcut
├── modMain          — PullSECFinancials pipeline + TestPhase1–6 macros
└── ThisWorkbook.cls — Workbook_Open / BeforeClose event handlers
```

### SEC API Endpoints

| Endpoint | URL |
|----------|-----|
| Ticker → CIK | `https://www.sec.gov/files/company_tickers.json` |
| Company Facts | `https://data.sec.gov/api/xbrl/companyfacts/CIK{10-digit}.json` |

All requests use `User-Agent: SECExcelAddin contact@example.com`. Rate cap: 5 req/sec.

### Verification Macros

Run any via `Alt + F8`:

| Macro | What it tests |
|-------|---------------|
| `TestPhase1` | Ticker → CIK resolution |
| `TestPhase2` | JSON fetch + us-gaap navigation |
| `TestPhase3` | Concept classification + annual/quarterly filters |
| `TestPhase4` | Full pipeline + 5 cell spot-checks |
| `TestPhase5` | All 5 error messages match PRD §4.5 exactly |
| `TestPhase6` | 100-cell audit table vs live SEC data (10 tickers × 10 checks) |

### QA

100/100 spot-checks verified against live SEC EDGAR JSON across 10 tickers
(AAPL, MSFT, GOOGL, AMZN, NEM, JPM, XOM, PFE, TSLA, BRK-B) on 2026-04-27.

---

## License

MIT. The bundled `dependencies/JsonConverter.bas` is [VBA-JSON v2.3.1](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, also MIT licensed. See `dependencies/LICENSE-VBA-JSON.txt`.
