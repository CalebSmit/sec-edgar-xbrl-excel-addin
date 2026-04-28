Attribute VB_Name = "modConfig"
'==============================================================================
' modConfig  -  Constants and configuration for SEC EDGAR XBRL Add-in
' PRD S5.2 | Phase 1 (revised Phase 3 with empirically verified keyword lists)
'
' KEYWORD LISTS (PRD S4.3)  -  verified against live AAPL (503 concepts) and
' MSFT (543 concepts) on 2026-04-27. All 30 Phase 3 verification concepts
' classify correctly (0 misses). Changes from PRD original:
'   - Removed bare 'Depreciation' from IS (collides with CFS DepreciationDepletion)
'   - Changed 'DepreciationDepletion' -> 'DepreciationDepletionAndAmortization' (specific)
'   - Changed 'StockBasedCompensation' -> also added 'AllocatedShareBasedCompensation'
'     (AAPL/MSFT use ShareBased- naming; StockBased- matches zero concepts in both)
'   - Added 'PaymentsTo' to CFS (catches PaymentsToAcquire* capex patterns)
'   - Kept 'StockBasedCompensation' per PRD (covers other filers)
'   - Removed 'Amortization' standalone from CFS (too broad; hits BS accumulated items)
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' SEC API Endpoints (verified 2026-04-27)
' ---------------------------------------------------------------------------
Public Const SEC_TICKER_URL As String = _
    "https://www.sec.gov/files/company_tickers.json"
Public Const SEC_FACTS_BASE_URL As String = _
    "https://data.sec.gov/api/xbrl/companyfacts/"

' ---------------------------------------------------------------------------
' HTTP Configuration
' SEC requires User-Agent in format: "AppName email@domain"
' https://www.sec.gov/os/webmaster-faq#user-agent
' SEC rate limit: 10 requests/second max (as of 2026).
' ---------------------------------------------------------------------------
Public Const HTTP_USER_AGENT As String = _
    "SEC-XBRL-Addin sec-xbrl-addin@outlook.com"
Public Const HTTP_TIMEOUT_MS As Long = 30000    ' 30s
Public Const RATE_LIMIT_DELAY_MS As Long = 110  ' 110ms = ~9 req/sec (safely under 10/s SEC limit)
Public Const HTTP_MAX_RETRIES As Long = 4       ' Initial try + 4 retries on transient errors
Public Const HTTP_RETRY_BASE_MS As Long = 750   ' Exponential backoff base
Public Const HTTP_RETRY_MAX_MS As Long = 10000  ' Cap individual retry wait at 10s
Public Const HTTP_RETRY_TOTAL_BUDGET_MS As Long = 60000 ' Give up after 60s total retry time
Public Const TICKER_CACHE_TTL_MINUTES As Long = 60 ' Cache company_tickers.json in memory

' ---------------------------------------------------------------------------
' Worksheet Names
' ---------------------------------------------------------------------------
Public Const WS_INCOME_STMT As String = "Income Statement"
Public Const WS_BALANCE_SHEET As String = "Balance Sheet"
Public Const WS_CASH_FLOW As String = "Cash Flow"

' ---------------------------------------------------------------------------
' Income Statement keywords (IS checked first per PRD S4.3 priority)
' Pipe-delimited; parsed by modClassifier.SplitKeywords()
'
' VERIFIED: 10/10 IS concepts in 30-concept check pass correctly.
' NOTE: 'Depreciation' removed  -  collides with CFS DepreciationDepletionAndAmortization.
' ---------------------------------------------------------------------------
Public Const KEYWORDS_IS As String = _
    "Revenue|Sales|CostOfGoods|CostOfRevenue|GrossProfit|" & _
    "OperatingExpenses|OperatingIncome|InterestExpense|IncomeTax|" & _
    "NetIncome|EarningsPerShare|WeightedAverage|" & _
    "ResearchAndDevelopment|SellingGeneralAndAdministrative"

' ---------------------------------------------------------------------------
' Cash Flow Statement keywords (CFS checked second)
'
' VERIFIED: 8/8 CFS concepts in 30-concept check pass correctly.
' Key additions vs PRD original:
'   'DepreciationDepletionAndAmortization' (specific, avoids AccumulatedDepreciation...)
'   'AllocatedShareBasedCompensation' (AAPL/MSFT naming convention for add-back)
'   'PaymentsTo' (catches PaymentsToAcquire* capex patterns)
' ---------------------------------------------------------------------------
Public Const KEYWORDS_CFS As String = _
    "CashFlow|NetCashProvided|NetCashUsed|PaymentsFor|PaymentsTo|" & _
    "ProceedsFrom|DepreciationDepletionAndAmortization|DepreciationAndAmortization|" & _
    "AllocatedShareBasedCompensation|StockBasedCompensation|" & _
    "CapitalExpenditures|DividendsPaid|RepurchaseOfCommonStock"

' ---------------------------------------------------------------------------
' Balance Sheet keywords (BS checked last)
'
' VERIFIED: 12/12 BS concepts in 30-concept check pass correctly.
' ---------------------------------------------------------------------------
Public Const KEYWORDS_BS As String = _
    "Assets|Liabilities|StockholdersEquity|CashAndCashEquivalents|" & _
    "Inventory|Receivable|Payable|Debt|Goodwill|IntangibleAssets|" & _
    "PropertyPlantAndEquipment|RetainedEarnings|CommonStock|" & _
    "TreasuryStock|AccumulatedOtherComprehensive"

' ---------------------------------------------------------------------------
' Error Codes (PRD S4.5)
' ---------------------------------------------------------------------------
Public Const ERR_TICKER_NOT_FOUND As String = "E1"
Public Const ERR_HTTP_RATE_LIMITED As String = "E2"
Public Const ERR_NO_NETWORK As String = "E3"
Public Const ERR_NO_USGAAP As String = "E4"
Public Const ERR_JSON_PARSE As String = "E5"

' ---------------------------------------------------------------------------
' Layout constants (PRD S4.4)
' ---------------------------------------------------------------------------
Public Const HDR_ANNUAL As String = "=== ANNUAL (10-K) ==="
Public Const HDR_QUARTERLY As String = "=== QUARTERLY (10-Q) ==="
Public Const COL_TAG As Integer = 1        ' Column A  -  XBRL concept name
Public Const COL_UNIT As Integer = 2       ' Column B  -  unit type (USD, shares, etc.)
Public Const COL_DATA_START As Integer = 3 ' Column C onward  -  period values
