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

Private mRequestInProgress As Boolean
Private mRandomSeeded As Boolean

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
Private Function GetHTTP(ByVal url As String, _
                         ByRef errCode As String, _
                         ByRef errMsg As String, _
                         Optional ByRef outStatusCode As Long = 0, _
                         Optional ByRef outRetryAfterSec As Long = 0) As String
    GetHTTP = ""
    errCode = ""
    errMsg = ""
    outStatusCode = 0
    outRetryAfterSec = 0

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

    ' SEC-required headers (PRD S3) + anti-bot headers
    ' The SEC bot detection checks for minimal HTTP compliance.
    ' Adding realistic headers reduces false-positive bot detection (HTTP 403).
    ' Per SEC developer FAQ: User-Agent format "AppName email" is required.
    Dim headerErr As Long

    On Error Resume Next
    http.SetRequestHeader "User-Agent", HTTP_USER_AGENT
    headerErr = Err.Number
    Err.Clear
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "Accept-Language", "en-US,en;q=0.9"
    http.SetRequestHeader "Accept-Encoding", "gzip, deflate"
    On Error GoTo 0

    If headerErr <> 0 Then
        errCode = ERR_NO_NETWORK
        errMsg = "Failed to configure SEC request headers. Close Excel and try again."
        Set http = Nothing
        Exit Function
    End If
    
    Application.StatusBar = "Sending request to: " & Left(url, 50) & "..."
    DoEvents

    ' --- Send --------------------------------------------------------------
    http.Send

    ' --- Evaluate response -------------------------------------------------
    Dim statusCode As Long
    statusCode = CLng(http.Status)
    outStatusCode = statusCode

    ' DEBUG: Log what happened
    Application.StatusBar = "SEC response: HTTP " & statusCode
    DoEvents

    Select Case statusCode
        Case 200
            GetHTTP = http.ResponseText      ' Raw JSON body

        Case 403, 429
            outRetryAfterSec = TryGetRetryAfterSeconds(http)

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
    Application.StatusBar = False
    errCode = ERR_NO_NETWORK
    errMsg = "Cannot connect to SEC servers. Check your internet connection."
    If Not http Is Nothing Then Set http = Nothing
End Function

'------------------------------------------------------------------------------
' RateLimitedGet
' Smart rate limiting: only sleep if the last request was recent.
' This avoids unnecessary delays on the first request while still enforcing
' the 200ms inter-request spacing for subsequent calls (5 req/sec cap).
' Uses millisecond-accurate kernel32 Sleep() for precision.
'
' IMPROVED: Eliminates first-request delay that could trigger SEC bot detection.
'------------------------------------------------------------------------------
Public Function RateLimitedGet(ByVal url As String, _
                               ByRef errCode As String, _
                               ByRef errMsg As String) As String
    Static lastRequestTimeMs As Double
    Dim startTimeMs As Double
    Dim attempt As Long
    Dim statusCode As Long
    Dim retryAfterSec As Long

    If mRequestInProgress Then
        errCode = ERR_HTTP_RATE_LIMITED
        errMsg = "A SEC request is already in progress. Please wait for it to finish."
        Exit Function
    End If

    mRequestInProgress = True
    On Error GoTo CleanupWithError

    startTimeMs = CDbl(Timer * 1000)

    For attempt = 1 To HTTP_MAX_RETRIES + 1
        If GetElapsedMs(startTimeMs, CDbl(Timer * 1000)) > HTTP_RETRY_TOTAL_BUDGET_MS Then
            errCode = ERR_NO_NETWORK
            errMsg = "SEC request timed out after multiple retries. Please try again."
            Exit For
        End If

        EnforceInterRequestDelay lastRequestTimeMs
        lastRequestTimeMs = CDbl(Timer * 1000)

        RateLimitedGet = GetHTTP(url, errCode, errMsg, statusCode, retryAfterSec)
        If errCode = "" Then Exit For

        If Not IsRetriableHTTPError(errCode, statusCode) Then
            If attempt > 1 Then
                errMsg = "Request ended with HTTP " & statusCode & " after " & attempt & " attempts. " & errMsg
            End If
            Exit For
        End If

        If attempt > HTTP_MAX_RETRIES Then
            errMsg = "Request failed after " & attempt & " attempts. " & errMsg
            Exit For
        End If

        Dim backoffMs As Long
        backoffMs = ComputeRetryDelayMs(attempt, retryAfterSec)

        Application.StatusBar = "SEC EDGAR: transient HTTP " & statusCode & _
                                ", retrying in " & Format(backoffMs / 1000, "0.0") & "s" & _
                                " (" & attempt & "/" & HTTP_MAX_RETRIES & ")"
        Sleep backoffMs
    Next attempt

Cleanup:
    mRequestInProgress = False
    Exit Function

CleanupWithError:
    mRequestInProgress = False
    errCode = ERR_NO_NETWORK
    errMsg = "Unexpected request error. Please retry."
End Function

Private Sub EnforceInterRequestDelay(ByVal lastRequestTimeMs As Double)
    Dim nowMs As Double
    nowMs = CDbl(Timer * 1000)

    If lastRequestTimeMs > 0 Then
        Dim elapsedMs As Double
        elapsedMs = GetElapsedMs(lastRequestTimeMs, nowMs)

        If elapsedMs < RATE_LIMIT_DELAY_MS Then
            Sleep CLng(RATE_LIMIT_DELAY_MS - elapsedMs)
        End If
    End If
End Sub

Private Function GetElapsedMs(ByVal fromMs As Double, ByVal toMs As Double) As Double
    Dim elapsedMs As Double
    elapsedMs = toMs - fromMs

    If elapsedMs < -43200000# Then
        elapsedMs = elapsedMs + 86400000#
    ElseIf elapsedMs < 0 Then
        elapsedMs = 0
    End If

    GetElapsedMs = elapsedMs
End Function

Private Function IsRetriableHTTPError(ByVal errCode As String, ByVal statusCode As Long) As Boolean
    Select Case statusCode
        Case 403, 429, 500, 502, 503, 504
            IsRetriableHTTPError = True
            Exit Function
    End Select

    If errCode = ERR_NO_NETWORK Then
        IsRetriableHTTPError = True
        Exit Function
    End If

    IsRetriableHTTPError = False
End Function

Private Function ComputeRetryDelayMs(ByVal attempt As Long, ByVal retryAfterSec As Long) As Long
    Dim expDelay As Double
    expDelay = CDbl(HTTP_RETRY_BASE_MS) * (2 ^ (attempt - 1))

    If expDelay > HTTP_RETRY_MAX_MS Then expDelay = HTTP_RETRY_MAX_MS

    ' Small jitter avoids repeated synchronized retries.
    Dim jitterMs As Long
    If Not mRandomSeeded Then
        Randomize
        mRandomSeeded = True
    End If
    jitterMs = CLng(Rnd() * 300)

    Dim computedMs As Long
    computedMs = CLng(expDelay) + jitterMs

    If retryAfterSec > 0 Then
        Dim retryAfterMs As Long
        retryAfterMs = CLng(retryAfterSec) * 1000
        If retryAfterMs > computedMs Then computedMs = retryAfterMs
    End If

    ComputeRetryDelayMs = computedMs
End Function

Private Function TryGetRetryAfterSeconds(ByVal http As Object) As Long
    On Error Resume Next
    Dim retryAfterRaw As String
    retryAfterRaw = Trim$(CStr(http.GetResponseHeader("Retry-After")))
    On Error GoTo 0

    If Len(retryAfterRaw) = 0 Then
        TryGetRetryAfterSeconds = 0
        Exit Function
    End If

    If IsNumeric(retryAfterRaw) Then
        TryGetRetryAfterSeconds = CLng(retryAfterRaw)
    Else
        ' HTTP-date Retry-After parsing omitted; use exponential backoff fallback.
        TryGetRetryAfterSeconds = 0
    End If
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
