# Claude Code — SEC EDGAR XBRL Excel Add-in

## Role
Developer working on a VBA Excel add-in (.xlam) that pulls live financial statement data directly from the SEC EDGAR XBRL API. 100% VBA — no external dependencies, no runtime API calls to anything except SEC.

## Repository
`https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin`

```bash
git clone https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin.git
cd sec-edgar-xbrl-excel-addin
```

## GitHub Sync Protocol — MANDATORY for every code-change request
Always treat GitHub `master` as the source of truth. Local files and the working
directory may be stale; do **NOT** assume they match the remote.

**Before making any code change:**
1. `cd` into the local clone of the repo (under `repo/` if not already cloned).
2. `git fetch origin master` and `git pull --ff-only origin master` (or fresh
   `git clone` if no local clone exists). Resolve any divergence before editing.
3. Confirm the working tree is clean (`git status`) — if not, surface that to
   the user before overwriting their in-progress work.
4. Read the latest source from disk after the pull. Never edit based on
   in-context memory of what files used to look like.

**After making code changes:**
1. Run the build (`scripts/build_xlam.ps1` on Windows) and any verification
   tests that apply (`scripts/smoke_test.ps1` at minimum).
2. Stage only intended files (`git add <specific paths>` — never `git add -A`).
3. Commit with a clear conventional-commits-style message describing the why.
4. `git push origin master`.
5. If the change is user-facing, bump the version, update the README's
   "Latest Fixes" section + download link, and create a `gh release create vX.Y.Z`
   tag (see "Release Workflow" below).
6. Verify the new download URL resolves before reporting "done" to the user.

This applies to **every** code-change request, even small ones. Do not skip
the pull; do not skip the push. If you forget, the next session will fight
stale state and the user has to re-explain.

## Project Structure
```
sec-edgar-xbrl-excel-addin/
├── dist/
│   └── SEC_XBRL_Addin.xlam          ← Distributable (built artifact)
├── modules/                          ← VBA source (edit here)
│   ├── modConfig.bas                 ← URLs, User-Agent, delays, classifier keywords
│   ├── modMain.bas                   ← Entry point, orchestration
│   ├── modHTTP.bas                   ← WinHTTP wrapper, retry, rate limit
│   ├── modJSONParser.bas             ← Pure VBA JSON parser
│   ├── modExcelWriter.bas            ← Writes to ActiveWorkbook
│   ├── modTickerLookup.bas           ← Ticker → CIK with in-memory cache
│   ├── modClassifier.bas             ← XBRL concepts → IS / BS / CFS
│   ├── modProgress.bas               ← Progress bar / status
│   ├── modRibbon.bas                 ← Ribbon XML callbacks
│   └── ThisWorkbook.cls
├── dependencies/
│   └── JsonConverter.bas             ← VBA-JSON (vendored, MIT)
├── customUI/
│   └── customUI14.xml                ← Ribbon definition
├── scripts/
│   ├── build_xlam.ps1                ← PRIMARY build script (Windows + Excel COM)
│   ├── smoke_test.ps1                ← No-network verification (~30s)
│   ├── e2e_test.ps1                  ← Full live-SEC pipeline test for AAPL
│   └── vba_purge_final.py            ← LEGACY (Linux-only, broken — do not use)
└── README.md
```

## Build Workflow — CRITICAL
You **cannot** edit VBA inside the .xlam directly. Workflow is always:

1. Edit `.bas` / `.cls` files in `modules/` (and `dependencies/JsonConverter.bas`
   if updating the JSON parser).
2. Run the build script (Windows + Excel installed):
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/build_xlam.ps1
   ```
3. Verify output shows all 12 modules with reasonable line counts and
   `=== Build OK ===` at the end.
4. Run the smoke test to confirm macros are reachable from a freshly installed
   add-in:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/smoke_test.ps1
   ```
   Expect `Smoke OK: 4/4 PASS`.

The build script:
- Launches Excel via COM, creates a fresh workbook, imports `ThisWorkbook.cls`
  + 9 standard `.bas` modules + `JsonConverter.bas`, sets `IsAddin=True`,
  saves as `xlam`.
- Repacks the resulting ZIP container to inject `customUI/customUI14.xml`
  and a top-level `_rels/.rels` with the Custom UI relationship (Excel does
  not write that during `SaveAs`).
- Re-opens the built xlam and lists modules as final verification.

Requires `AccessVBOM=1` ("Trust access to the VBA project object model")
in Excel — the script auto-detects and prints a clear error if missing.

> ⚠ The legacy `scripts/vba_purge_final.py` is **broken** (hardcoded Linux
> paths, missing `bak3` baseline) and produces a structurally non-standard
> xlam that some MS-OVBA readers cannot parse. Do not use it. It is kept in
> the tree only as historical reference.

## Key Technical Details

### SEC API
- Ticker list: `https://www.sec.gov/files/company_tickers.json`
- Company facts: `https://data.sec.gov/api/xbrl/companyfacts/CIK{10-digit-zero-padded}.json`
- **User-Agent (NEVER change):** `SEC-XBRL-Addin sec-xbrl-addin@outlook.com`
- Rate limit: ~9 req/s → 110 ms delay between requests
- **NEVER add `Accept-Encoding: gzip`** — Excel's WinHTTP does not auto-decompress; SEC will send gzip and break JSON parsing

### VBA Constraints
- No external libraries (no Scripting Runtime, no JSON COM objects)
- Use `WinHttp.WinHttpRequest.5.1` only — never `MSXML2.ServerXMLHTTP`
- Output always to `ActiveWorkbook`, never the add-in workbook
- HTTP 403/429/5xx → retry with exponential backoff + jitter

### Configuration
All URLs, User-Agent, rate-limit delays, and classifier keyword lists live in **`modConfig.bas`**. Edit only that file for config changes.

### Classifier (`modClassifier.bas` + `modConfig.bas`)
Maps SEC XBRL concept names to IS/BS/CFS. Keyword lists were empirically verified against AAPL and MSFT — adding overly broad terms causes misclassification.

## Release Workflow

After testing changes (and after pulling latest master per the GitHub Sync
Protocol at the top of this file):

1. **Update source modules** in `modules/`
2. **Rebuild xlam:**
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/build_xlam.ps1
   ```
3. **Verify:**
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/smoke_test.ps1
   ```
   Expect `Smoke OK: 4/4 PASS`.
4. **Update README:**
   - Bump `**Version:** X.X.X` at top
   - Add new `## Latest Fixes — vX.X.X (YYYY-MM-DD)` section at top of changelog
   - Demote previous "Latest Fixes" → `## Previous Fixes — vX.X.X`
   - Update download link to new tag:
     `https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin/raw/vX.X.X/dist/SEC_XBRL_Addin.xlam`
5. **Commit & push:**
   ```bash
   git add dist/SEC_XBRL_Addin.xlam modules/ README.md scripts/
   git commit -m "vX.X.X: <concise description>"
   git push origin master
   ```
6. **Create GitHub release with tag:**
   ```bash
   gh release create vX.X.X \
     --title "vX.X.X — <short description>" \
     --notes "## What Changed
   <bullets>

   ## Download
   **[SEC_XBRL_Addin.xlam](https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin/raw/vX.X.X/dist/SEC_XBRL_Addin.xlam)**"
   ```
7. **Verify the download URL resolves** before reporting completion:
   ```powershell
   Invoke-WebRequest -Uri 'https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin/raw/vX.X.X/dist/SEC_XBRL_Addin.xlam' -OutFile $env:TEMP/verify.xlam
   ```
   Note: `gh release upload` for binary assets may not work in all environments. The raw tag URL above is the canonical download link.

## README Update Rules (always)
Every commit that changes functionality must update the README:
- Version badge at top
- New `## Latest Fixes — vX.X.X` section; demote old one to `## Previous Fixes`
- Download link → new tag URL
- Installation steps only if install process changed
- **Keep all existing troubleshooting sections intact**

## Version History
- **v1.0.8** — fixed compile error from misplaced UDT in modClassifier; added top-level error trap to clear StatusBar on any failure; build script now rejects non-ASCII source bytes
- **v1.0.7** — replaced VBA-purge build with Excel-COM rebuild (`scripts/build_xlam.ps1`); hardened Workbook_Open; added `PullSECFinancialsForTicker` entry point
- **v1.0.6** — VBA purging applied; modConfig load error permanently fixed (since superseded by v1.0.7's clean rebuild)
- **v1.0.5** — compressor fix attempt (partial)
- **v1.0.4** — User-Agent updated to fix SEC HTTP 403
- **v1.0.3** — removed `Accept-Encoding: gzip` header to fix E5 parse error
- **v1.0.2** — output to active workbook, locale-safe dates, retry logic, caching
- **v1.0.1** — initial release

## Common Mistakes — Do Not Do These

| Mistake | Consequence |
|---|---|
| Editing the xlam binary directly | Corrupts OLE streams, breaks all VBA |
| Skipping `scripts/build_xlam.ps1` after editing modules | Source changes never make it into the xlam |
| Running the legacy `scripts/vba_purge_final.py` | Hardcoded Linux paths fail; produces a non-standard xlam |
| Adding `Accept-Encoding: gzip` | JSON arrives as binary garbage, parse fails |
| Using `ActiveSheet` instead of `ActiveWorkbook` | Output goes to wrong/hidden sheet |
| Changing the User-Agent string | SEC returns HTTP 403 for all API calls |
| Skipping `git pull` before editing | Edits land on stale source; merge conflicts on push |
| Forgetting to `git push` after a build | Local xlam diverges from GitHub release; users download stale version |
| Bumping version without updating README | Users download old version thinking it's new |
