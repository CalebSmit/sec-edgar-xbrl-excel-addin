Attribute VB_Name = "modProgress"
'==============================================================================
' modProgress — Status bar progress messages and user-facing feedback
' PRD §4.1 FR-3, FR-4 | Phase 5
'
' All status bar messages are centralised here so they are consistent and
' easy to update. No MsgBox calls in this module — those live in modMain.
'==============================================================================
Option Explicit

' Progress step identifiers — passed to ShowProgress
Public Const PROG_RESOLVING As String = "Resolving ticker..."
Public Const PROG_FETCHING As String = "Fetching SEC data..."
Public Const PROG_PARSING As String = "Parsing JSON..."
Public Const PROG_CLASSIFYING As String = "Classifying concepts..."
Public Const PROG_WRITING_IS As String = "Writing Income Statement..."
Public Const PROG_WRITING_BS As String = "Writing Balance Sheet..."
Public Const PROG_WRITING_CFS As String = "Writing Cash Flow Statement..."
Public Const PROG_DONE As String = "SEC EDGAR data load complete."

'------------------------------------------------------------------------------
' ShowProgress
' Writes a status bar message (PRD FR-3). Adds the add-in name as prefix.
'------------------------------------------------------------------------------
Public Sub ShowProgress(ByVal message As String)
    Application.StatusBar = "SEC EDGAR: " & message
    DoEvents   ' Allow Excel to repaint the status bar immediately
End Sub

'------------------------------------------------------------------------------
' ClearProgress
' Resets the status bar to Excel's default (clears our message).
'------------------------------------------------------------------------------
Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

'==============================================================================
' Error Message Functions
' Returns the correct user-facing message for each PRD §4.5 error code.
' Centralised here so Phase 5 test can verify exact strings.
'==============================================================================

'------------------------------------------------------------------------------
' GetErrorMessage
' Returns {title, message} for a given error code + optional context string.
' context: e.g. the ticker symbol for E1, the HTTP status code for others.
'------------------------------------------------------------------------------
Public Sub GetErrorMessage(ByVal errCode As String, _
                           ByVal context As String, _
                           ByRef msgTitle As String, _
                           ByRef msgBody As String)
    Select Case errCode
        Case ERR_TICKER_NOT_FOUND   ' E1
            msgTitle = "Ticker Not Found (E1)"
            If Len(context) > 0 Then
                msgBody = "Ticker '" & UCase(context) & "' not found in SEC database."
            Else
                msgBody = "Ticker not found in SEC database."
            End If

        Case ERR_HTTP_RATE_LIMITED  ' E2
            msgTitle = "SEC Rate Limit (E2)"
            msgBody = "SEC rate-limited. Please wait 30 seconds and try again."

        Case ERR_NO_NETWORK         ' E3
            msgTitle = "No Connection (E3)"
            msgBody = "Cannot connect to SEC servers. Check your internet connection."

        Case ERR_NO_USGAAP          ' E4
            msgTitle = "No XBRL Data (E4)"
            msgBody = "No US-GAAP XBRL data found for this company."

        Case ERR_JSON_PARSE         ' E5
            msgTitle = "Parse Error (E5)"
            msgBody = "Failed to parse SEC response. The data format may have changed."

        Case Else
            msgTitle = "Error"
            msgBody = "An unexpected error occurred (code: " & errCode & ")."
    End Select
End Sub

'------------------------------------------------------------------------------
' ShowError
' Displays a MsgBox with the correct error message for the given error code.
' This is the single call site for ALL error display (PRD FR-4).
'------------------------------------------------------------------------------
Public Sub ShowError(ByVal errCode As String, Optional ByVal context As String = "")
    ClearProgress

    Dim msgTitle As String, msgBody As String
    GetErrorMessage errCode, context, msgTitle, msgBody

    MsgBox msgBody, vbExclamation, "SEC EDGAR Add-in — " & msgTitle
End Sub
