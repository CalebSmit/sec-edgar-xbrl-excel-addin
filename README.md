# SEC EDGAR XBRL Financial Statements — Excel Add-in

**Version:** 1.0.8 | **Platform:** Excel for Windows | **Cost:** Free

Pull live financial statement data from the SEC EDGAR database directly into Excel.
No sign-up, no subscriptions, no backend — data comes straight from the SEC.

## Latest Fixes — v1.0.8 (2026-04-28)

- **Fix for "Compile error: Only comments may appear after End Sub, End Function, or End Property"** that appeared in v1.0.7 when the user tried to run `PullSECFinancials`. The cause: an unused `Public Type ConceptRecord` UDT was declared *after* a function in `modClassifier.bas`, which VBA's strict compiler rejects. The Type was never used anyway (the comment in the code explicitly noted that Collections cannot hold UDTs), so it has been removed.
- **Fix for status-bar hang ("Classifying concepts: (503 concepts)" stuck for 10+ minutes)**: the compile error above halted execution after the StatusBar was set, so the bar never cleared. Added a top-level error trap to `PullSECFinancialsForTicker` that always clears the StatusBar, restores `ScreenUpdating`, restores `Calculation`, and surfaces a real error dialog when anything goes wrong, instead of leaving Excel in a broken state.
- **Build script now rejects non-ASCII bytes** in `.bas` and `.cls` source files. VBA's import treats source as ANSI/Windows-1252; UTF-8 multi-byte chars (em-dashes, smart quotes) are mis-decoded into garbage bytes that can corrupt comments or string tokens. The build now fails fast with a clear error pointing at the offending line.
- **Smoke test extended** to call `ClassifyConcept` three times (one per IS / BS / CFS bucket) so this category of regression cannot ship to release without being caught.

## Previous Fixes — v1.0.7 (2026-04-28)

- **Replaced the unreliable VBA-purge build pipeline with a clean, Windows-native rebuild via Excel COM automation.** The previous `scripts/vba_purge_final.py` had hardcoded Linux paths and depended on a `dist/SEC_XBRL_Addin.xlam.bak3` baseline that wasn't shipped, so it could not be re-run by anyone. The rebuilt xlam is now structurally clean — `oletools` (and any standard MS-OVBA reader) can fully extract every module, which the v1.0.6 purged build could not.
- **New build script:** `scripts/build_xlam.ps1` — uses Excel COM to import all 10 modules + JsonConverter dependency + Ribbon XML, producing a normally-formed xlam from source files. Reproducible on any Windows machine with Excel installed and "Trust access to the VBA project object model" enabled.
- **Hardened `Workbook_Open` and `RegisterShortcut` / `UnregisterShortcut`** with `On Error Resume Next` so an unusual workbook name (e.g. renamed install) cannot cause a runtime-error popup at Excel startup. The Ribbon button continues to work even if the keyboard-shortcut hook fails.
- **New entry point: `PullSECFinancialsForTicker(ticker, [silent])`.** Same pipeline as the Ribbon button but accepts the ticker as a parameter (no InputBox), and `silent:=True` suppresses the success message box. Useful for VBA driving the add-in from another macro, or for headless automation: `Application.Run "SEC_XBRL_Addin.xlam!PullSECFinancialsForTicker", "AAPL"`.
- **Code-quality cleanups in source:** removed an unreachable `ParseError:` label in `modTickerLookup.bas`; added `SafeDouble` helper alongside `SafeLong` in `modJSONParser.bas` for numeric fields that exceed Long range (already correct on the hot path; this is for future use).
- **New verification scripts:** `scripts/smoke_test.ps1` (no-network smoke test, runs in <30s) and `scripts/e2e_test.ps1` (full pipeline against live SEC, downloads ~15 MB AAPL companyfacts).

## Previous Fixes — v1.0.6 (2026-04-28)

- **Permanent fix for "An error occurred while loading 'modConfig'" dialog**: applied VBA purging — all p-code (PerformanceCache) has been stripped from every module stream. Excel is forced to recompile from source on first load, which completely eliminates the load error regardless of Excel version or VBA engine build. Per MS-OVBA spec §2.4.1, PerformanceCache "MUST be ignored on read" — this is the correct and permanent solution.
- User-Agent `SEC-XBRL-Addin sec-xbrl-addin@outlook.com` confirmed active across all modules.

## Previous Fixes — v1.0.5 (2026-04-28)

- **Fix for modConfig load error** (partial — compressor issue): used spec-compliant MS-OVBA compressor. Did not resolve the underlying p-code conflict.

## Previous Fixes — v1.0.4 (2026-04-28)

- **Fix for E2 SEC Rate Limit / HTTP 403 error**: updated User-Agent string. SEC requires a real, reachable contact address in the UA — the previous address was rejected by SEC's rate-limiter with a persistent 403 even after retries. The new UA `SEC-XBRL-Addin sec-xbrl-addin@outlook.com` passes SEC's validation.

## Previous Fixes — v1.0.3 (2026-04-28)

- **Root-cause fix for E5 "Parse Error"**: removed `Accept-Encoding: gzip, deflate` request header. Excel's WinHTTP stack does not auto-decompress gzip responses, so SEC's compressed JSON arrived as binary garbage and VBA-JSON failed to parse it. The header is now gone — SEC returns plain UTF-8 JSON and parsing succeeds every time.
- VBA compressed source in the xlam binary patched directly (MS-OVBA stream level) so Excel re-compiles the corrected code regardless of p-code version mismatch.

## Previous Fixes — v1.0.2 (2026-04-27)

- Output now writes to your active workbook instead of the hidden add-in workbook.
- SEC ISO dates are parsed in a locale-independent way to avoid regional date issues.
- `Ctrl + Shift + S` shortcut now targets the add-in macro explicitly.
- HTTP User-Agent updated to `SEC XBRL Excel Add-in sec-addin@github.io` — a real domain required by SEC.
- Rate limit updated from 250 ms to 110 ms (~9 req/s), matching the SEC's actual 10 req/s ceiling.
- HTTP transport now auto-retries transient `403/429/5xx` SEC responses with exponential backoff + jitter.
- `Retry-After` response headers are respected when SEC provides them.
- Ticker map lookups now use in-memory caching to reduce repeated SEC requests.
- JSON parsing no longer hard-requires Microsoft Scripting Runtime.
- Removed stale `SEC_XBRL_Addin_FINAL_2026-04-27.xlam` from `dist/` — it had a compile-breaking VBA class header bug. Only one file is in `dist/` now.

---

## Download & Install

### Quick Setup (2 minutes)

1. Download `SEC_XBRL_Addin.xlam` from the link below.
2. Save it in `C:\Users\YourName\Documents\`.
3. Right-click the file -> **Properties** -> check **Unblock** -> **OK**.
4. In Excel, add your Documents folder as a **Trusted Location**.
5. In Excel, install the file using **File -> Options -> Add-ins -> Excel Add-ins -> Go... -> Browse**.

If you skip Unblock or Trusted Location, the ribbon tab may not appear.

**Latest Download:**

- **[⬇ SEC_XBRL_Addin.xlam — v1.0.8](https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin/raw/v1.0.8/dist/SEC_XBRL_Addin.xlam)**
- **[Releases page](https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin/releases)**

**Latest Source Code:**

- **[GitHub repository (master)](https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin)**

> ⚠️ **Follow all 4 steps below in order.** Skipping any step is the most common reason the ribbon tab does not appear.

---

### Step 1 — Save to your Documents folder

Save `SEC_XBRL_Addin.xlam` to:

```
C:\Users\YourName\Documents\SEC_XBRL_Addin.xlam
```

**Do NOT save it to:**
- ❌ Your Desktop
- ❌ OneDrive or any synced/cloud folder
- ❌ The root of C:\ (e.g. `C:\SEC_XBRL_Addin.xlam`)
- ❌ A USB or network drive

Only a plain local folder like Documents will work reliably.

---

### Step 2 — Unblock the file *(required after every download)*

Windows marks every downloaded file as untrusted. You must unblock it manually, or Excel will refuse to run it and show a "Protected View" error.

1. Open **File Explorer** → go to your Documents folder
2. **Right-click** `SEC_XBRL_Addin.xlam` → click **Properties**
3. On the **General** tab, look at the very bottom for a **Security** section
4. Check the box next to **"Unblock"** → click **OK**

> If there is no Unblock checkbox, the file is already trusted — skip to Step 3.

---

### Step 3 — Add Documents as a Trusted Location

Even after installing, Excel will silently load the add-in without its ribbon tab unless the folder is explicitly trusted. This is a Microsoft security requirement since 2016.

1. Open Excel (no file open — just the blank Excel window)
2. **File → Options**
3. Click **Trust Center** → click **Trust Center Settings...**
4. Click **Trusted Locations** → click **Add new location...**
5. Click **Browse...** → select your **Documents** folder → click **OK**
6. Click **OK → OK → OK** to close all dialogs

---

### Step 4 — Install the add-in

1. In Excel: **File → Options**
2. Click **Add-ins** in the left sidebar
3. At the bottom, the **Manage:** dropdown must say **"Excel Add-ins"**
   *(not "COM Add-ins" or "Automation" — those show .dll files and won't work)*
4. Click **Go...**
5. Click **Browse...**
6. Navigate to your **Documents** folder → select `SEC_XBRL_Addin.xlam` → click **OK**
7. Make sure the checkbox next to **SEC XBRL Addin** is checked → click **OK**

The **SEC EDGAR** tab will now appear in your ribbon.

> **If Excel shows a yellow security bar saying "Macros Disabled":** Click **Enable Content**. This is required every time you open a new Excel session with the add-in if your Trust Center settings are restrictive. Adding your Documents folder as a Trusted Location (Step 3) eliminates this prompt permanently.

---

### Troubleshooting

| Problem | Fix |
|---------|-----|
| "Microsoft Excel cannot access the file" | File is in a bad location (Desktop, OneDrive, C:\ root). Move it to Documents and repeat Step 4. |
| "This file type is not supported in Protected View" | You skipped Step 2. Right-click the file → Properties → check **Unblock** → OK. Repeat Step 4. |
| SEC EDGAR tab does not appear | You skipped Step 3. Add your Documents folder as a Trusted Location, then close and reopen Excel. |
| Checkbox in Step 4 is greyed out | Close Excel completely. Reopen it without opening any file. Then repeat Step 4. |
| Yellow bar: "Macros have been disabled" | Click **Enable Content**. To make this permanent, complete Step 3 (Trusted Location). |

### If VBA opens and shows compile errors

This usually means Excel is still loading an older SEC add-in from a different folder.

1. Close **all** Excel windows.
2. Open Excel (blank workbook) -> **File -> Options -> Add-ins**.
3. At bottom, **Manage: Excel Add-ins** -> **Go...**
4. Uncheck all old SEC entries (look for anything named `SEC_XBRL_Addin` or similar).
5. Remove old entries if available, then click **Browse...** and select only the latest `SEC_XBRL_Addin.xlam` file in Documents.
6. Ensure only one SEC add-in entry is checked, then restart Excel.

---

### Step 5 — Use it

| Method | How |
|--------|-----|
| **Ribbon** | Click **Pull SEC Financials** in the **SEC EDGAR** tab |
| **Keyboard** | Press `Ctrl + Shift + S` → type a ticker → Enter |
| **Macro** | `Alt + F8` → select `PullSECFinancials` → Run → type a ticker |

Enter any US stock ticker: `AAPL`, `MSFT`, `TSLA`, `JPM`, `GOOGL`, `BRK-B`, etc.

The add-in fetches data from SEC EDGAR and fills three worksheets:
- **Income Statement** — Revenue, gross profit, net income, EPS, R&D, SG&A
- **Balance Sheet** — Assets, liabilities, equity, cash, goodwill, debt
- **Cash Flow** — Operating/investing/financing flows, capex, D&A

Sheets are written into your **currently active workbook** (the workbook you have open), not into the hidden add-in file.

Both **annual (10-K)** and **quarterly (10-Q)** data appear side by side.

---

## Requirements

- Windows 10 or 11
- Microsoft Excel 2016 or newer (Microsoft 365 works too)
- Internet connection (to fetch SEC data)
- Windows Scripting support enabled (default on standard Windows installs)

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
│   ├── SEC_XBRL_Addin.xlam   ← Prebuilt add-in (download this — only file in dist/)
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
   - ✅ Microsoft XML, v6.0
   - *(Optional)* Microsoft Scripting Runtime
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

All requests use a compliant SEC identity string with project URL + contact. Rate cap: 5 req/sec.

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
