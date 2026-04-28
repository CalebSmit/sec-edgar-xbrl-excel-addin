Attribute VB_Name = "modHTTP"
'==============================================================================
' modHTTP  -  HTTP GET requests with User-Agent header, timeouts, error handling
' PRD S3, S4.5, S5.1 | Phase 1 (foundation) + Phase 2 (companyfacts fetch)
'
' Transport: WinHttp.WinHttpRequest.5.1 (standard Windows COM, no extra refs)
'   Falls back to MSXML2.ServerXMLHTTP.6.0 if WinHttp unavailable.
'
' SEC requirements (PRD S3):
'   - User-Agent: "AppName email@domain" (real domain required per SEC FAQ)
'   - Rate cap: 10 req/sec max per SEC docs; 110ms delay = ~9 req/sec safe
'   - HTTPS only
'   - IMPORTANT: Do NOT send Accept-Encoding: gzip header.
'     WinHttp and MSXML2.ServerXMLHTTP do NOT auto-decompress gzip responses.
'     Without this header, SEC returns plain UTF-8 JSON that VBA can parse.
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

    Application.StatusBar = "Sending request to: " & Left(url, 50) & "..."

    ' Primary transport: WinHTTP. Fallback: MSXML transport if 403 persists.
    GetHTTP = ExecuteSingleRequest(url, True, errCode, errMsg, outStatusCode, outRetryAfterSec)

    If errCode <> "" And outStatusCode = 403 Then
        Dim altErrCode As String, altErrMsg As String
        Dim altStatusCode As Long, altRetryAfterSec As Long
        Dim altResponse As String

        Application.StatusBar = "SEC EDGAR: HTTP 403 on primary transport, trying fallback transport..."
        altResponse = ExecuteSingleRequest(url, False, altErrCode, altErrMsg, altStatusCode, altRetryAfterSec)

        If altErrCode = "" Then
            GetHTTP = altResponse
            errCode = ""
            errMsg = ""
            outStatusCode = altStatusCode
            outRetryAfterSec = altRetryAfterSec
            Exit Function
        End If

        ' If fallback returned a more actionable status, surface that result.
        If altStatusCode <> 403 And altStatusCode <> 0 Then
            errCode = altErrCode
            errMsg = altErrMsg
            outStatusCode = altStatusCode
            outRetryAfterSec = altRetryAfterSec
        End If
    End If
End Function

Private Function ExecuteSingleRequest(ByVal url As String, _
                                      ByVal preferWinHttp As Boolean, _
                                      ByRef errCode As String, _
                                      ByRef errMsg As String, _
                                      ByRef outStatusCode As Long, _
                                      ByRef outRetryAfterSec As Long) As String
    ExecuteSingleRequest = ""
    errCode = ""
    errMsg = ""
    outStatusCode = 0
    outRetryAfterSec = 0

    Dim http As Object
    Dim usingWinHttp As Boolean
    Set http = CreateHttpClient(preferWinHttp, usingWinHttp)

    If http Is Nothing Then
        errCode = ERR_NO_NETWORK
        errMsg = "Cannot create HTTP client. Check your network and Excel security settings."
        Exit Function
    End If

    On Error GoTo RequestError

    http.Open "GET", url, False

    If usingWinHttp Then
        http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, _
                         HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
    End If

    If Not ApplySECHeaders(http, errCode, errMsg) Then
        Set http = Nothing
        Exit Function
    End If

    http.Send

    outStatusCode = CLng(http.Status)
    Application.StatusBar = "SEC response: HTTP " & outStatusCode

    Select Case outStatusCode
        Case 200
            ExecuteSingleRequest = http.ResponseText

        Case 403, 429
            outRetryAfterSec = TryGetRetryAfterSeconds(http)
            errCode = ERR_HTTP_RATE_LIMITED

            If outStatusCode = 403 Then
                errMsg = "SEC blocked request (HTTP 403). The add-in will retry automatically."
            Else
                errMsg = "SEC rate-limited (HTTP 429). The add-in will retry automatically."
            End If

        Case 404
            errCode = ERR_TICKER_NOT_FOUND
            errMsg = "Resource not found on SEC servers (HTTP 404)."

        Case Else
            errCode = ERR_NO_NETWORK
            errMsg = "SEC server returned HTTP " & outStatusCode & ". Please try again later."
    End Select

    Set http = Nothing
    Exit Function

RequestError:
    Application.StatusBar = False
    errCode = ERR_NO_NETWORK
    errMsg = "Cannot connect to SEC servers. Check your internet connection."
    Set http = Nothing
End Function

Private Function CreateHttpClient(ByVal preferWinHttp As Boolean, ByRef usingWinHttp As Boolean) As Object
    Set CreateHttpClient = Nothing
    usingWinHttp = False

    On Error Resume Next

    If preferWinHttp Then
        Set CreateHttpClient = CreateObject("WinHttp.WinHttpRequest.5.1")
        If Not CreateHttpClient Is Nothing Then
            usingWinHttp = True
            On Error GoTo 0
            Exit Function
        End If

        Err.Clear
        Set CreateHttpClient = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        usingWinHttp = False
    Else
        Set CreateHttpClient = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        If Not CreateHttpClient Is Nothing Then
            usingWinHttp = False
            On Error GoTo 0
            Exit Function
        End If

        Err.Clear
        Set CreateHttpClient = CreateObject("WinHttp.WinHttpRequest.5.1")
        usingWinHttp = Not CreateHttpClient Is Nothing
    End If

    On Error GoTo 0
End Function

Private Function ApplySECHeaders(ByVal http As Object, ByRef errCode As String, ByRef errMsg As String) As Boolean
    ApplySECHeaders = False
    errCode = ""
    errMsg = ""

    On Error Resume Next
    http.SetRequestHeader "User-Agent", HTTP_USER_AGENT
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        errCode = ERR_NO_NETWORK
        errMsg = "Failed to set SEC User-Agent header. Restart Excel and retry."
        Exit Function
    End If

    http.SetRequestHeader "Accept", "application/json"
    'Accept-Encoding omitted: SEC returns plain JSON w/o it

    Dim contactEmail As String
    contactEmail = ExtractContactEmail(HTTP_USER_AGENT)
    If Len(contactEmail) > 0 Then
        http.SetRequestHeader "From", contactEmail
    End If

    On Error GoTo 0
    ApplySECHeaders = True
End Function

Private Function ExtractContactEmail(ByVal userAgent As String) As String
    ExtractContactEmail = ""

    Dim parts() As String
    parts = Split(Trim$(userAgent), " ")

    If UBound(parts) >= 0 Then
        Dim candidate As String
        candidate = parts(UBound(parts))
        If InStr(1, candidate, "@", vbTextCompare) > 0 Then
            ExtractContactEmail = candidate
        End If
    End If
End Function

'------------------------------------------------------------------------------
' RateLimitedGet
' Smart rate limiting: only sleep if the last request was recent.
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
        TryGetRetryAfterSeconds = 0
    End If
End Function

'------------------------------------------------------------------------------
' FetchCompanyFacts
' Fetches the full companyfacts JSON for a 10-digit CIK.
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

    If errCode <> "" Then
        Application.StatusBar = False
        Exit Function
    End If

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
'------------------------------------------------------------------------------
Public Function GetResponseSize(ByVal jsonText As String) As String
    GetResponseSize = Format(Len(jsonText) / 1024, "#,##0") & " KB"
End Function
