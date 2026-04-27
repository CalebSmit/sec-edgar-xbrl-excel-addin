Attribute VB_Name = "modJSONParser"
'==============================================================================
' modJSONParser — JSON parsing and navigation for SEC companyfacts structure
' PRD §5.1, §5.3 | Phase 2
'
' Dependency: VBA-JSON (JsonConverter.bas) — MIT license
'   https://github.com/VBA-tools/VBA-JSON
'   Bundled in dependencies/JsonConverter.bas (pre-baked into dist/SEC_XBRL_Addin.xlam).
'
' VERIFIED JSON structure (live AAPL test 2026-04-27):
'   {
'     "cik": 320193,
'     "entityName": "Apple Inc.",
'     "facts": {
'       "dei": { ... },          ← SKIP per PRD §4.3
'       "us-gaap": {
'         "ConceptName": {
'           "label": "...",
'           "description": "...",
'           "units": {
'             "USD": [           ← or "USD/shares", "shares", "pure"
'               {
'                 "end":   "2024-09-28",   ← always present
'                 "val":   93736000000,    ← always present (0 facts missing val verified)
'                 "accn":  "0000320193-...",
'                 "fy":    2025,
'                 "fp":    "FY",           ← "FY","Q1","Q2","Q3"
'                 "form":  "10-K",         ← "10-K","10-Q","10-K/A","8-K"
'                 "filed": "2025-10-31",
'                 "start": "2024-09-29",   ← OPTIONAL: absent for instant (BS) facts
'                 "frame": "..."           ← OPTIONAL
'               }
'             ]
'           }
'         }
'       }
'     }
'   }
'
' Key filtering rules (empirically verified):
'   ANNUAL:    form="10-K" AND fp="FY" AND (no start → instant) OR (duration >= 300 days)
'   QUARTERLY: form="10-Q" AND fp IN ("Q1","Q2","Q3") AND duration 60-110 days (duration facts)
'              OR form="10-Q" AND no start (instant BS facts)
'   DEDUP:     group by end_date, keep latest filed date
'   8-K facts: excluded (1,118 found in AAPL — not from 10-K/10-Q filings)
'==============================================================================
Option Explicit

'==============================================================================
' SECTION 1 — Core JSON Parsing
'==============================================================================

'------------------------------------------------------------------------------
' ParseSECJson
' Wraps JsonConverter.ParseJson. Returns Nothing on failure.
' VBA-JSON returns Scripting.Dictionary for objects, Collection for arrays.
' NOTE: Named ParseSECJson (not ParseJSON) to avoid name clash with
' JsonConverter.bas which also exports a ParseJSON symbol.
'------------------------------------------------------------------------------
Public Function ParseSECJson(ByVal jsonText As String) As Object
    On Error GoTo ParseFailed
    Set ParseSECJson = JsonConverter.ParseJson(jsonText)
    Exit Function
ParseFailed:
    Set ParseSECJson = Nothing
End Function

'------------------------------------------------------------------------------
' GetDictKey
' Safely retrieves a child dictionary from a parent dictionary by key.
' Returns Nothing if key absent or value is not an object.
'------------------------------------------------------------------------------
Public Function GetDictKey(ByVal dict As Object, ByVal key As String) As Object
    Set GetDictKey = Nothing
    If dict Is Nothing Then Exit Function
    On Error Resume Next
    If dict.Exists(key) Then Set GetDictKey = dict(key)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' SafeString
' Returns dict(key) as String, or "" if absent or non-string.
'------------------------------------------------------------------------------
Public Function SafeString(ByVal dict As Object, ByVal key As String) As String
    SafeString = ""
    If dict Is Nothing Then Exit Function
    On Error Resume Next
    If dict.Exists(key) Then SafeString = CStr(dict(key))
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' SafeLong
' Returns dict(key) as Long, or 0 if absent.
'------------------------------------------------------------------------------
Public Function SafeLong(ByVal dict As Object, ByVal key As String) As Long
    SafeLong = 0
    If dict Is Nothing Then Exit Function
    On Error Resume Next
    If dict.Exists(key) Then SafeLong = CLng(dict(key))
    On Error GoTo 0
End Function

'==============================================================================
' SECTION 2 — Companyfacts Navigation
'==============================================================================

'------------------------------------------------------------------------------
' GetUSGAAP
' Navigates facts → us-gaap and returns the us-gaap Dictionary.
' Returns Nothing and sets errCode=ERR_NO_USGAAP if absent.
'
' VERIFIED path: parsed("facts")("us-gaap")
' Note: "us-gaap" key contains a hyphen — VBA-JSON handles this correctly
' as a string key in Scripting.Dictionary.
'------------------------------------------------------------------------------
Public Function GetUSGAAP(ByVal parsed As Object, _
                          ByRef errCode As String, _
                          ByRef errMsg As String) As Object
    Set GetUSGAAP = Nothing
    errCode = ""
    errMsg = ""

    If parsed Is Nothing Then
        errCode = ERR_JSON_PARSE
        errMsg = "Failed to parse SEC response. The data format may have changed."
        Exit Function
    End If

    ' Navigate: root → "facts"
    Dim facts As Object
    Set facts = GetDictKey(parsed, "facts")
    If facts Is Nothing Then
        errCode = ERR_NO_USGAAP
        errMsg = "No US-GAAP XBRL data found for this company."
        Exit Function
    End If

    ' Navigate: "facts" → "us-gaap"
    Dim usGaap As Object
    Set usGaap = GetDictKey(facts, "us-gaap")
    If usGaap Is Nothing Then
        errCode = ERR_NO_USGAAP
        errMsg = "No US-GAAP XBRL data found for this company."
        Exit Function
    End If

    ' Sanity: must have at least one key
    If usGaap.Count = 0 Then
        errCode = ERR_NO_USGAAP
        errMsg = "No US-GAAP XBRL data found for this company."
        Exit Function
    End If

    Set GetUSGAAP = usGaap
End Function

'==============================================================================
' SECTION 3 — Fact Filtering and Deduplication
'==============================================================================
' These functions operate on VBA-JSON Collections (arrays of fact objects).
' Each fact object is a Scripting.Dictionary with keys verified above.

'------------------------------------------------------------------------------
' IsAnnualFact
' Returns True if a fact belongs in the ANNUAL (10-K) section.
'
' Rule (empirically verified):
'   form = "10-K" AND fp = "FY"
'   AND either:
'     (a) no "start" key (instant/point-in-time — for BS concepts)
'     (b) duration (end - start) >= 300 days (full-year — for IS/CFS concepts)
'
' This correctly excludes the ~90-day sub-period facts that AAPL's 10-K
' also files under fp="FY" (restated quarterly breakdowns within the annual).
'------------------------------------------------------------------------------
Public Function IsAnnualFact(ByVal fact As Object) As Boolean
    IsAnnualFact = False
    If fact Is Nothing Then Exit Function

    ' Must be 10-K with fp=FY
    If SafeString(fact, "form") <> "10-K" Then Exit Function
    If SafeString(fact, "fp") <> "FY" Then Exit Function

    ' Check for instant vs duration
    If Not fact.Exists("start") Then
        ' Instant fact (no start date) — valid for BS annual
        IsAnnualFact = True
    Else
        ' Duration fact — require >= 300 days for full-year
        Dim dur As Long
        dur = DateDiff("d", CDate(SafeString(fact, "start")), _
                           CDate(SafeString(fact, "end")))
        IsAnnualFact = (dur >= 300)
    End If
End Function

'------------------------------------------------------------------------------
' IsQuarterlyFact
' Returns True if a fact belongs in the QUARTERLY (10-Q) section.
'
' Rule (empirically verified):
'   form = "10-Q" AND fp IN ("Q1","Q2","Q3")
'   AND either:
'     (a) no "start" key (instant — BS quarterly snapshot)
'     (b) duration 60-110 days (single-quarter duration fact for IS/CFS)
'
' Note: Q4 is NOT reported in 10-Q; it derives from 10-K only.
' We do NOT synthesize Q4. Per PRD §4.6 FR-14: no estimation.
'------------------------------------------------------------------------------
Public Function IsQuarterlyFact(ByVal fact As Object) As Boolean
    IsQuarterlyFact = False
    If fact Is Nothing Then Exit Function

    If SafeString(fact, "form") <> "10-Q" Then Exit Function

    Dim fp As String
    fp = SafeString(fact, "fp")
    If fp <> "Q1" And fp <> "Q2" And fp <> "Q3" Then Exit Function

    If Not fact.Exists("start") Then
        ' Instant fact — valid for BS quarterly
        IsQuarterlyFact = True
    Else
        Dim dur As Long
        dur = DateDiff("d", CDate(SafeString(fact, "start")), _
                           CDate(SafeString(fact, "end")))
        IsQuarterlyFact = (dur >= 60 And dur <= 110)
    End If
End Function

'------------------------------------------------------------------------------
' DeduplicateFacts
' Given a Collection of fact Dictionaries (all from the same unit array),
' pre-filtered by IsAnnualFact or IsQuarterlyFact, returns a Scripting.Dictionary
' keyed by end_date, holding the single best fact per period.
'
' Dedup rule: keep the fact with the latest "filed" date per end_date.
' This picks the most recent restated/amended value — same strategy the
' SEC EDGAR viewer itself uses.
'
' Returns: Scripting.Dictionary{ end_date_string → fact_Dictionary }
'------------------------------------------------------------------------------
Public Function DeduplicateFacts(ByVal factsCol As Collection) As Object
    ' result: end_date (String) → best fact (Dictionary)
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    If factsCol Is Nothing Then
        Set DeduplicateFacts = result
        Exit Function
    End If

    Dim fact As Object
    For Each fact In factsCol
        Dim endDate As String
        endDate = SafeString(fact, "end")
        If Len(endDate) = 0 Then GoTo NextFact  ' skip malformed

        Dim filedDate As String
        filedDate = SafeString(fact, "filed")

        If Not result.Exists(endDate) Then
            result(endDate) = fact
        Else
            ' Keep the later filed date
            Dim existingFiled As String
            existingFiled = SafeString(result(endDate), "filed")
            If filedDate > existingFiled Then
                result(endDate) = fact
            End If
        End If

NextFact:
    Next fact

    Set DeduplicateFacts = result
End Function

'------------------------------------------------------------------------------
' FilterAndDedup
' Convenience: given a VBA-JSON Collection (the raw units array for one concept+unit),
' filters by annual or quarterly rule, then deduplicates.
'
' Parameters:
'   unitsArray  — VBA-JSON Collection of fact Dicts
'   isAnnual    — True → filter with IsAnnualFact; False → IsQuarterlyFact
'
' Returns: Scripting.Dictionary{ end_date → fact }
'------------------------------------------------------------------------------
Public Function FilterAndDedup(ByVal unitsArray As Object, _
                               ByVal isAnnualSection As Boolean) As Object
    Dim filtered As New Collection
    Dim fact As Object

    On Error Resume Next
    For Each fact In unitsArray
        If isAnnualSection Then
            If IsAnnualFact(fact) Then filtered.Add fact
        Else
            If IsQuarterlyFact(fact) Then filtered.Add fact
        End If
    Next fact
    On Error GoTo 0

    Set FilterAndDedup = DeduplicateFacts(filtered)
End Function

'------------------------------------------------------------------------------
' GetSortedEndDates
' Returns a sorted array of end_date strings from a Scripting.Dictionary
' (as produced by DeduplicateFacts / FilterAndDedup).
' Sorted ascending (oldest → newest) per PRD §4.4 "columns sorted ascending."
'------------------------------------------------------------------------------
Public Function GetSortedEndDates(ByVal dedupDict As Object) As String()
    Dim n As Long
    n = dedupDict.Count

    Dim arr() As String
    ReDim arr(0 To n - 1)

    Dim i As Long
    i = 0
    Dim k As Variant
    For Each k In dedupDict.Keys
        arr(i) = CStr(k)
        i = i + 1
    Next k

    ' Bubble sort ascending (n is small — typically < 30 periods)
    Dim j As Long
    Dim tmp As String
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If arr(j) > arr(j + 1) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i

    GetSortedEndDates = arr
End Function

'==============================================================================
' SECTION 4 — Verification Utility (Phase 2 testing)
'==============================================================================

'------------------------------------------------------------------------------
' GetFirst20USGAAPKeys
' Returns the first 20 us-gaap concept names as a newline-delimited string.
' Used in Phase 2 verification MsgBox.
'------------------------------------------------------------------------------
Public Function GetFirst20USGAAPKeys(ByVal usGaap As Object) As String
    Dim result As String
    Dim i As Integer
    i = 1
    Dim k As Variant
    For Each k In usGaap.Keys
        result = result & i & ". " & CStr(k) & vbCrLf
        i = i + 1
        If i > 20 Then Exit For
    Next k
    GetFirst20USGAAPKeys = result
End Function
