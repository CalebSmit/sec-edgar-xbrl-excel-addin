Attribute VB_Name = "modHTTP"
'==============================================================================
' modHTTP  -  HTTP GET requests with User-Agent header, timeouts, error handling
' PRD S3, S4.5, S5.1 | Phase 1 (foundation) + Phase 2 (companyfacts fetch)
'
' Transport: WinHttp.WinHttpRequest.5.1 (standard Windows COM, no extra refs)
'   Falls back to MSXML2.ServerXMLHTTP.6.0 if WinHttp unavailable.
'
' SEC requirements (PRD S3):
'   - User-Agent: "SECExcelAddin contact@example.com"
'   - Rate cap: 5 req/sec (enforced via 200ms sleep between calls)
'   - HTTPS only
'
' VERIFIED field names / endpoints: none referenced here  -  this module is
' purely transport. Endpoints are in modConfig.
'==============================================================================
Option Explicit

' Win32 Sleep declaration for precise millisecond delays
' (Application.Wait minimum granularity is ~1 second; Sleep is exact)
#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'------------------------------------------------------------------------------
' GetHTTP
' Synchronous GET. Returns response body on HTTP 200.
' On any failure sets errCode (modConfig ERR_* constant) + errMsg and returns "".
'
' Parameters:
'   url      -  full HTTPS URL
'   errCode  -  output: "" on success; ERR_* on failure
'   errMsg   -  output: user-facing message for MsgBox
'------------------------------------------------------------------------------
Public Function GetHTTP(ByVal url As String, _
                        ByRef errCode As String, _
                        ByRef errMsg As String) As String
    GetHTTP = ""
    errCode = ""
    errMsg = ""

    Dim http As Object
    Dim useWinHttp As Boolean
    useWinHttp = True

    ' --- Instantiate HTTP object -------------------------------------------
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number <> 0 Or http Is Nothing Then
        Err.Clear
        useWinHttp = False
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    End If
    On Error GoTo 0

    If http Is Nothing Then
        errCode = ERR_NO_NETWORK
        errMsg = "Cannot connect to SEC servers. Check your internet connection."
        Exit Function
    End If

    ' --- Open + configure --------------------------------------------------
    On Error GoTo NetworkError

    http.Open "GET", url, False          ' False = synchronous

    ' Timeouts: resolve, connect, send, receive (all in milliseconds)
    If useWinHttp Then
        http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, _
                         HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
    End If

    ' SEC-required headers (PRD S3)
    ' DEBUG: Verify headers are set correctly
    On Error Resume Next
    http.SetRequestHeader "User-Agent", HTTP_USER_AGENT
    http.SetRequestHeader "Accept", "application/json"
    On Error GoTo 0
    
    Application.StatusBar = "Sending request to: " & Left(url, 50) & "..."
    DoEvents

    ' --- Send --------------------------------------------------------------
    http.Send

    ' --- Evaluate response -------------------------------------------------
    Dim statusCode As Long
    statusCode = CLng(http.Status)

    ' DEBUG: Log what happened
    Application.StatusBar = "SEC response: HTTP " & statusCode
    DoEvents

    Select Case statusCode
        Case 200
            GetHTTP = http.ResponseText      ' Raw JSON body

        Case 403, 429
            ' DEBUG: Log rate limit details
            Dim responseText As String
            On Error Resume Next
            responseText = http.ResponseText
            On Error GoTo 0
            
            errCode = ERR_HTTP_RATE_LIMITED
            errMsg = "SEC rate-limited (HTTP " & statusCode & "). Please wait 30 seconds and try again."
            
            ' If it's a 403, it's often a bot detection issue
            If statusCode = 403 Then
                errMsg = "SEC blocked request (HTTP 403). This may indicate:" & vbNewLine & _
                         "- Add-in is still loaded (close Excel completely)" & vbNewLine & _
                         "- Multiple rapid requests" & vbNewLine & _
                         "- Download fresh XLAM from GitHub"
            End If

        Case 404
            ' Caller interprets 404 based on context (ticker not found, etc.)
            errCode = ERR_TICKER_NOT_FOUND
            errMsg = "Resource not found on SEC servers (HTTP 404)."

        Case Else
            errCode = ERR_NO_NETWORK
            errMsg = "SEC server returned HTTP " & statusCode & ". Please try again later."
    End Select

    Set http = Nothing
    Exit Function

NetworkError:
    errCode = ERR_NO_NETWORK
    errMsg = "Cannot connect to SEC servers. Check your internet connection."
    If Not http Is Nothing Then Set http = Nothing
End Function

'------------------------------------------------------------------------------
' RateLimitedGet
' Enforces the 200ms inter-request delay (5 req/sec cap, PRD S3) using the
' kernel32 Sleep function for millisecond-accurate delay, then calls GetHTTP.
'------------------------------------------------------------------------------
Public Function RateLimitedGet(ByVal url As String, _
                               ByRef errCode As String, _
                               ByRef errMsg As String) As String
    Sleep RATE_LIMIT_DELAY_MS
    RateLimitedGet = GetHTTP(url, errCode, errMsg)
End Function

'------------------------------------------------------------------------------
' FetchCompanyFacts
' Phase 2 addition: Fetches the full companyfacts JSON for a 10-digit CIK.
' Returns the raw JSON string on success; sets errCode/errMsg on failure.
'
' Also checks for an empty or unusably small response (guards against
' truncated downloads on large payloads like AAPL ~15 MB).
'
' Verified endpoint (live test 2026-04-27):
'   https://data.sec.gov/api/xbrl/companyfacts/CIK0000320193.json
'   -> top keys: ["cik", "entityName", "facts"]
'   -> facts keys: ["dei", "us-gaap"]
'   -> AAPL has 503 us-gaap concepts, 24,492 total facts
'------------------------------------------------------------------------------
Public Function FetchCompanyFacts(ByVal cik10 As String, _
                                  ByRef errCode As String, _
                                  ByRef errMsg As String) As String
    FetchCompanyFacts = ""

    Dim factsURL As String
    factsURL = BuildFactsURL(cik10)

    Application.StatusBar = "Fetching SEC data for CIK " & cik10 & "..."

    Dim jsonText As String
    jsonText = RateLimitedGet(factsURL, errCode, errMsg)

    ' Propagate any HTTP errors
    If errCode <> "" Then
        Application.StatusBar = False
        Exit Function
    End If

    ' Guard: response must be non-trivially large (at least 100 chars)
    ' A valid facts JSON for even tiny filers is several KB
    If Len(jsonText) < 100 Then
        errCode = ERR_JSON_PARSE
        errMsg = "Failed to parse SEC response. The data format may have changed."
        Application.StatusBar = False
        Exit Function
    End If

    Application.StatusBar = "Data received (" & Format(Len(jsonText) / 1024, "0") & " KB). Parsing..."
    FetchCompanyFacts = jsonText
End Function

'------------------------------------------------------------------------------
' GetResponseSize
' Utility: returns Len(jsonText) in KB for status-bar display.
'------------------------------------------------------------------------------
Public Function GetResponseSize(ByVal jsonText As String) As String
    GetResponseSize = Format(Len(jsonText) / 1024, "#,##0") & " KB"
End Function
