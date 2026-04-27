# SEC EDGAR XBRL Financial Statements — Excel Add-in

**Version:** 1.0.0 (Phase 6 — Release)  
**Platform:** Excel for Windows (VBA .xlam)  
**Cost:** $0 — no subscriptions, no hosting, no backend  
**Data Source:** SEC EDGAR public API (no auth required)  
**QA:** 100/100 spot-checks verified against live SEC JSON (2026-04-27)

---

## What It Does

Enter a ticker symbol (e.g. `AAPL`) → the add-in fetches XBRL financial statement data directly from the SEC EDGAR API and populates three worksheets:

| Sheet | Contents |
|-------|----------|
| Income Statement | Revenue, gross profit, net income, EPS, R&D, SG&A, etc. |
| Balance Sheet | Assets, liabilities, equity, cash, goodwill, debt, etc. |
| Cash Flow | Operating/investing/financing cash flows, capex, D&A, etc. |

Each sheet shows both **annual (10-K)** and **quarterly (10-Q)** data, exactly as reported in SEC filings — no renaming, no estimation, no LLM inference. Each cell value traces 1:1 to the `val` field in the raw SEC JSON.

---

## Installation

### Step 1 — Download VBA-JSON

The add-in uses [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) (MIT license) for JSON parsing.

1. Go to https://github.com/VBA-tools/VBA-JSON/releases
2. Download `JsonConverter.bas`
3. Save it to the same folder as `SEC_XBRL_Addin.xlam`

### Step 2 — Open the Add-in in Excel

1. Open Excel
2. Press `Alt + F11` to open the VBA editor
3. In the VBA editor: **File → Import File** → import `JsonConverter.bas`
4. **File → Import File** → import each `.bas` / `.cls` file from the `modules/` folder:
   - `modConfig.bas`
   - `modHTTP.bas`
   - `modTickerLookup.bas`
   - `modJSONParser.bas`
   - `modClassifier.bas`
   - `modExcelWriter.bas`
   - `modProgress.bas`
   - `modRibbon.bas`
   - `modMain.bas`
   - `ThisWorkbook.cls` ← import into the **ThisWorkbook** module (do not create a new module)

### Step 3 — Install the Custom Ribbon

1. Close the VBA editor
2. Copy `customUI/customUI14.xml` into the `.xlam` package:
   - Save the file as `.xlam` first (Step 4)
   - Use a tool like [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) to inject `customUI14.xml`
   - The ribbon adds a **"Pull SEC Financials"** button to the **Add-ins** tab

### Step 4 — Required VBA References

In the VBA editor, go to **Tools → References** and ensure these are checked:
- ✅ Microsoft Scripting Runtime (`scrrun.dll`) — required for `Scripting.Dictionary`
- ✅ Microsoft XML, v6.0 — required for HTTP (usually already checked)

### Step 5 — Save as .xlam

1. In Excel: **File → Save As**
2. Choose file type: **Excel Add-in (.xlam)**
3. Save as `SEC_XBRL_Addin.xlam`

### Step 6 — Install Add-in

1. **File → Options → Add-ins → Manage: Excel Add-ins → Go**
2. Click **Browse** → select `SEC_XBRL_Addin.xlam`
3. Check the box next to it → OK

---

## Usage

### Option A — Ribbon Button
Click **"Pull SEC Financials"** in the **Add-ins** tab.

### Option B — Keyboard Shortcut
Press `Ctrl + Shift + S`

### Option C — Macro
1. Press `Alt + F8`
2. Select `PullSECFinancials` → Run
3. Enter a ticker (e.g. `AAPL`, `MSFT`, `GOOGL`)

### Workflow
1. Enter a ticker → add-in resolves it to a CIK via `company_tickers.json`
2. Fetches the full `companyfacts` JSON from SEC EDGAR
3. Classifies all US-GAAP XBRL concepts into IS / BS / CFS buckets
4. Writes annual and quarterly data to three worksheets in a sparse-matrix layout

---

## Architecture

```
SEC_XBRL_Addin.xlam
├── modConfig        — Constants, API endpoints, keyword lists, error codes, sheet names
├── modHTTP          — WinHttp GET with User-Agent header, rate-limit delay, gzip, retries
├── modTickerLookup  — Ticker → 10-digit CIK resolution via SEC company_tickers.json
├── modJSONParser    — JSON navigation, IsAnnualFact, IsQuarterlyFact,
│                      DeduplicateFacts, FilterAndDedup, GetSortedEndDates
├── modClassifier    — Keyword-based concept → IS/BS/CFS bucketing,
│                      unit preference (USD > USD/shares > shares > pure)
├── modExcelWriter   — Worksheet creation, sparse-matrix layout, CDbl() cell writes,
│                      GetCellAuditInfo for TestPhase6
├── modProgress      — ShowProgress, ClearProgress, ShowError, GetErrorMessage;
│                      all PROG_* and ERR_* constants (PRD §4.5)
├── modRibbon        — RibbonCallback_PullSECFinancials, Ctrl+Shift+S shortcut
├── modMain          — PullSECFinancials (full pipeline),
│                      TestPhase1..TestPhase6 verification macros
└── ThisWorkbook.cls — Workbook_Open (register shortcut), Workbook_BeforeClose (cleanup)

customUI/
└── customUI14.xml   — Office 2010+ Ribbon XML (Add-ins tab button)
```

---

## Data Integrity

- **Zero LLM calls** — entirely deterministic VBA + SEC API
- **Every cell** traces 1:1 to a `val` field in the raw SEC JSON — no estimation
- **No interpolation** — missing data = empty cell (PRD FR-14)
- **Annual filter:** `form=10-K`, `fp=FY`, duration ≥ 300 days (or instant/BS fact)
- **Quarterly filter:** `form=10-Q`, `fp ∈ {Q1,Q2,Q3}`, 60–110 day duration (or instant/BS fact)
- **Dedup:** Latest `filed` date wins per `end_date` — handles restatements
- **Numeric type:** `CDbl()` for all cell writes — handles values up to 4 trillion (JPM assets)
- **Audit:** 100/100 spot-checks passed across 10 tickers × 10 checks (2026-04-27)

---

## SEC API Endpoints Used

| Endpoint | URL |
|----------|-----|
| Ticker → CIK | `https://www.sec.gov/files/company_tickers.json` |
| Company Facts | `https://data.sec.gov/api/xbrl/companyfacts/CIK{10-digit}.json` |

All requests include `User-Agent: SECExcelAddin contact@example.com` per SEC requirements.  
Rate limit: 5 requests/second (SEC allows 10; we cap at 5).

---

## Keyword Classification

Concepts are classified by XBRL tag name matching (case-insensitive substring):

**Income Statement (14 keywords):**  
`Revenue|Sales|CostOfGoods|CostOfRevenue|GrossProfit|OperatingExpenses|OperatingIncome|InterestExpense|IncomeTax|NetIncome|EarningsPerShare|WeightedAverage|ResearchAndDevelopment|SellingGeneralAndAdministrative`

**Cash Flow Statement (13 keywords):**  
`CashFlow|NetCashProvided|NetCashUsed|PaymentsFor|PaymentsTo|ProceedsFrom|DepreciationDepletionAndAmortization|DepreciationAndAmortization|AllocatedShareBasedCompensation|StockBasedCompensation|CapitalExpenditures|DividendsPaid|RepurchaseOfCommonStock`

**Balance Sheet (15 keywords):**  
`Assets|Liabilities|StockholdersEquity|CashAndCashEquivalents|Inventory|Receivable|Payable|Debt|Goodwill|IntangibleAssets|PropertyPlantAndEquipment|RetainedEarnings|CommonStock|TreasuryStock|AccumulatedOtherComprehensive`

Concepts matching no keyword are silently skipped (written to no sheet).

---

## Error Codes

| Code | Condition | Message |
|------|-----------|---------| 
| E1 | Invalid/not found ticker | `"Ticker '{ticker}' not found in SEC database."` |
| E2 | HTTP 403/429 rate limit | `"SEC rate-limited. Please wait 30 seconds and try again."` |
| E3 | No network | `"Cannot connect to SEC servers. Check your internet connection."` |
| E4 | No us-gaap facts | `"No US-GAAP XBRL data found for this company."` |
| E5 | JSON parse failure | `"Failed to parse SEC response. The data format may have changed."` |

---

## Verification Macros

Run any of these via `Alt + F8`:

| Macro | What it tests |
|-------|--------------|
| `TestPhase1` | Ticker → CIK resolution (AAPL, MSFT, FAKEXYZ) |
| `TestPhase2` | JSON fetch, parse, us-gaap navigation |
| `TestPhase3` | Concept classification (30-concept table), annual/quarterly filter |
| `TestPhase4` | Full pipeline + 5 cell spot-checks vs SEC JSON |
| `TestPhase5` | All 5 error messages match PRD §4.5 exactly |
| `TestPhase6` | 100-cell audit table (10 tickers × 10 checks) vs pre-verified expected values |

---

## License

MIT. VBA-JSON dependency is also MIT licensed (https://github.com/VBA-tools/VBA-JSON).
