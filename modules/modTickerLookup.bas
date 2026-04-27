Attribute VB_Name = "modTickerLookup"
'==============================================================================
' modTickerLookup  -  Resolve a ticker symbol to a 10-digit zero-padded CIK
' PRD S4.2, S4.5 | Phase 1
'
' SEC company_tickers.json structure (verified 2026-04-27):
'   {
'     "0": { "cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc." },
'     "1": { "cik_str": 789019, "ticker": "MSFT", "title": "MICROSOFT CORP" },
'     ...
'   }
'
' Key fields used:
'   cik_str   -  integer CIK (NOT zero-padded in the JSON)
'   ticker    -  uppercase ticker string
'   title     -  company name (used for display only)
'
' VERIFIED: cik_str is an integer in JSON, ticker is uppercase.
' Zero-padding to 10 digits is applied in FormatCIK() below.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' ResolveTicker
' Looks up a ticker in SEC company_tickers.json and returns the 10-digit
' zero-padded CIK string (e.g., "0000320193").
'
' Returns "" and sets errCode/errMsg on any failure.
'
' Parameters:
'   ticker       -  any case (e.g. "aapl", "AAPL", "Aapl")
'   errCode      -  output: "" on success, ERR_* constant on failure
'   errMsg       -  output: user-facing error message
'   companyName  -  output: entity name from SEC (for display)
'------------------------------------------------------------------------------
Public Function ResolveTicker(ByVal ticker As String, _
                              ByRef errCode As String, _
                              ByRef errMsg As String, _
                              Optional ByRef companyName As String = "") As String
    ResolveTicker = ""
    errCode = ""
    errMsg = ""
    companyName = ""
    
    ' Normalize ticker to uppercase
    Dim tickerUC As String
    tickerUC = UCase(Trim(ticker))
    
    If Len(tickerUC) = 0 Then
        errCode = ERR_TICKER_NOT_FOUND
        errMsg = "Ticker '" & ticker & "' not found in SEC database."
        Exit Function
    End If
    
    ' Update status bar (PRD FR-3)
    Application.StatusBar = "Resolving ticker: " & tickerUC & "..."
    
    ' --- Fetch the ticker->CIK mapping file ---------------------------------
    Dim httpErr As String, httpMsg As String
    Dim jsonText As String
    jsonText = RateLimitedGet(SEC_TICKER_URL, httpErr, httpMsg)
    
    If httpErr <> "" Then
        errCode = httpErr
        errMsg = httpMsg
        Application.StatusBar = False
        Exit Function
    End If
    
    If Len(jsonText) = 0 Then
        errCode = ERR_JSON_PARSE
        errMsg = "Failed to parse SEC response. The data format may have changed."
        Application.StatusBar = False
        Exit Function
    End If
    
    ' --- Parse JSON --------------------------------------------------------
    ' VBA-JSON (JsonConverter) returns a Scripting.Dictionary for JSON objects
    Dim parsed As Object
    On Error GoTo ParseError
    Set parsed = JsonConverter.ParseJson(jsonText)
    On Error GoTo 0
    
    If parsed Is Nothing Then
        errCode = ERR_JSON_PARSE
        errMsg = "Failed to parse SEC response. The data format may have changed."
        Application.StatusBar = False
        Exit Function
    End If
    
    ' --- Iterate over entries to find matching ticker -----------------------
    ' The outer object is a dictionary keyed "0", "1", "2", etc.
    ' Each value is itself a dictionary with keys: cik_str, ticker, title
    Dim key As Variant
    Dim entry As Object
    
    For Each key In parsed.Keys
        Set entry = parsed(key)
        
        ' entry("ticker") is uppercase in the SEC file (verified)
        If entry.Exists("ticker") Then
            If UCase(entry("ticker")) = tickerUC Then
                ' Found  -  extract CIK integer and format to 10-digit string
                Dim cikInt As Long
                cikInt = CLng(entry("cik_str"))
                ResolveTicker = FormatCIK(cikInt)
                
                ' Capture company name if available
                If entry.Exists("title") Then
                    companyName = entry("title")
                End If
                
                Application.StatusBar = False
                Exit Function
            End If
        End If
    Next key
    
    ' --- Not found ----------------------------------------------------------
    errCode = ERR_TICKER_NOT_FOUND
    errMsg = "Ticker '" & tickerUC & "' not found in SEC database."
    Application.StatusBar = False
    Exit Function

ParseError:
    errCode = ERR_JSON_PARSE
    errMsg = "Failed to parse SEC response. The data format may have changed."
    Application.StatusBar = False
    Exit Function
End Function

'------------------------------------------------------------------------------
' FormatCIK
' Converts an integer CIK to a 10-digit zero-padded string.
' E.g., 320193 -> "0000320193"
' VERIFIED: SEC companyfacts endpoint requires exactly 10 digits.
'------------------------------------------------------------------------------
Public Function FormatCIK(ByVal cikInt As Long) As String
    FormatCIK = Format(cikInt, "0000000000")
End Function

'------------------------------------------------------------------------------
' BuildFactsURL
' Constructs the companyfacts URL for a given 10-digit CIK.
' VERIFIED endpoint: https://data.sec.gov/api/xbrl/companyfacts/CIK{10digit}.json
'------------------------------------------------------------------------------
Public Function BuildFactsURL(ByVal cik10 As String) As String
    BuildFactsURL = SEC_FACTS_BASE_URL & "CIK" & cik10 & ".json"
End Function
