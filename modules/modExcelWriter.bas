Attribute VB_Name = "modExcelWriter"
'==============================================================================
' modExcelWriter — Create / clear three worksheets and write XBRL data
' PRD §4.4 | Phase 4
'
' LAYOUT (same pattern for IS, BS, CFS tabs — PRD §4.4):
'
'   Row 1   Col A: "=== ANNUAL (10-K) ==="
'   Row 2   Col A: "XBRL Tag"  Col B: "Unit"  Col C+: period end-dates (ISO, sorted asc)
'   Row 3+  Col A: conceptName  Col B: unit  Col C+: val (Double) or empty
'   [blank row]
'   Row N   Col A: "=== QUARTERLY (10-Q) ==="
'   Row N+1 Col A: "XBRL Tag"  Col B: "Unit"  Col C+: quarterly end-dates (sorted asc)
'   Row N+2+ same concept rows, quarterly values
'
' DATA INTEGRITY (PRD §4.6):
'   - Every numeric cell value = CDbl(fact["val"]) from SEC JSON — no transformation
'   - Empty cells STAY EMPTY — no zero-fill, no interpolation (FR-14)
'   - Concept names written exactly as XBRL tag — no renaming (FR-13)
'   - String column headers = ISO end-date strings from JSON "end" field
'
' NUMERIC HANDLING:
'   - VBA Long max = 2,147,483,647. AAPL values reach 416,161,000,000 → overflow
'   - SOLUTION: ALL numeric vals written as CDbl() → VBA Double (53-bit mantissa)
'   - Double can represent integers up to 2^53 = 9,007,199,254,740,992 exactly
'   - EPS floats (e.g. 4.04) also handled correctly by Double
'==============================================================================
Option Explicit

'==============================================================================
' SECTION 1 — Worksheet Management
'==============================================================================

'------------------------------------------------------------------------------
' InitWorksheets
' Creates or clears the three output worksheets in the target workbook.
' Sheet order: Income Statement | Balance Sheet | Cash Flow
' Pre-existing sheets are cleared completely (content + formatting).
'------------------------------------------------------------------------------
Public Sub InitWorksheets(ByVal wb As Workbook)
    CreateOrClearSheet wb, WS_INCOME_STMT
    CreateOrClearSheet wb, WS_BALANCE_SHEET
    CreateOrClearSheet wb, WS_CASH_FLOW
End Sub

'------------------------------------------------------------------------------
' CreateOrClearSheet
' If a sheet with the given name exists → clear all cells and formatting.
' If it doesn't exist → create it.
'------------------------------------------------------------------------------
Private Sub CreateOrClearSheet(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.ClearContents
        ws.Cells.ClearFormats
        ws.Cells.Interior.ColorIndex = xlNone
    End If
End Sub

'==============================================================================
' SECTION 2 — Main Write Orchestrator
'==============================================================================

'------------------------------------------------------------------------------
' WriteAllSheets
' Top-level writer called from modMain.PullSECFinancials after classification.
' Writes data to all three sheets.
'
' Parameters:
'   wb            — target workbook (ThisWorkbook for .xlam callers)
'   isCollection  — Collection of IS concept Dicts from modClassifier
'   bsCollection  — Collection of BS concept Dicts
'   cfsCollection — Collection of CFS concept Dicts
'   ticker        — for sheet tab metadata (not written to cells — future use)
'------------------------------------------------------------------------------
Public Sub WriteAllSheets(ByVal wb As Workbook, _
                          ByVal isCollection As Collection, _
                          ByVal bsCollection As Collection, _
                          ByVal cfsCollection As Collection, _
                          ByVal ticker As String)
    ' Create / clear sheets
    ShowProgress "Initializing worksheets..."
    InitWorksheets wb

    ' Write each sheet — use modProgress constants (PRD FR-3)
    ShowProgress PROG_WRITING_IS
    WriteSheet wb.Worksheets(WS_INCOME_STMT), isCollection

    ShowProgress PROG_WRITING_BS
    WriteSheet wb.Worksheets(WS_BALANCE_SHEET), bsCollection

    ShowProgress PROG_WRITING_CFS
    WriteSheet wb.Worksheets(WS_CASH_FLOW), cfsCollection

    ClearProgress

    ' Activate Income Statement as landing tab
    wb.Worksheets(WS_INCOME_STMT).Activate
End Sub

'==============================================================================
' SECTION 3 — Single-Sheet Writer
'==============================================================================

'------------------------------------------------------------------------------
' WriteSheet
' Writes annual + quarterly sections to one worksheet.
'
' Algorithm:
'   1. Collect ALL unique annual end-dates across all concepts → sort → col headers
'   2. Write annual section (rows 1..N)
'   3. Write blank gap row
'   4. Collect ALL unique quarterly end-dates → sort → col headers
'   5. Write quarterly section
'
' The same concept Collection is used for both sections; a concept that has
' no annual data simply has all empty cells in the annual section, and vice versa.
'------------------------------------------------------------------------------
Private Sub WriteSheet(ByVal ws As Worksheet, ByVal concepts As Collection)
    If concepts Is Nothing Then Exit Sub
    If concepts.Count = 0 Then Exit Sub

    ' Turn off screen updating for speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- 1. Collect all unique end-dates -----------------------------------
    Dim allAnnualDates As Object   ' Scripting.Dictionary: date → True (used as set)
    Dim allQtrDates As Object
    Set allAnnualDates = CreateObject("Scripting.Dictionary")
    Set allQtrDates = CreateObject("Scripting.Dictionary")

    Dim rec As Object
    For Each rec In concepts
        Dim annDict As Object
        Dim qtrDict As Object
        Set annDict = rec("AnnualData")
        Set qtrDict = rec("QuarterlyData")

        Dim d As Variant
        For Each d In annDict.Keys
            If Not allAnnualDates.Exists(CStr(d)) Then allAnnualDates(CStr(d)) = True
        Next d
        For Each d In qtrDict.Keys
            If Not allQtrDates.Exists(CStr(d)) Then allQtrDates(CStr(d)) = True
        Next d
    Next rec

    ' Sort dates ascending
    Dim annDates() As String
    Dim qtrDates() As String
    annDates = SortDictKeys(allAnnualDates)
    qtrDates = SortDictKeys(allQtrDates)

    ' --- 2. Write annual section -------------------------------------------
    Dim currentRow As Long
    currentRow = 1

    currentRow = WriteSection(ws, concepts, annDates, True, currentRow)

    ' --- 3. Blank gap row --------------------------------------------------
    currentRow = currentRow + 1   ' one blank row between sections

    ' --- 4. Write quarterly section ----------------------------------------
    currentRow = WriteSection(ws, concepts, qtrDates, False, currentRow)

    ' --- 5. Auto-fit columns A and B only (date cols can be narrow) --------
    ws.Columns(COL_TAG).ColumnWidth = 55
    ws.Columns(COL_UNIT).ColumnWidth = 12

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'------------------------------------------------------------------------------
' WriteSection
' Writes one section (annual or quarterly) to the worksheet starting at startRow.
' Returns the next available row after this section.
'
' Parameters:
'   ws           — target worksheet
'   concepts     — Collection of concept Dicts
'   periodDates  — sorted array of end-date strings (column headers)
'   isAnnual     — True → write annual data; False → quarterly
'   startRow     — row to begin writing at
'------------------------------------------------------------------------------
Private Function WriteSection(ByVal ws As Worksheet, _
                               ByVal concepts As Collection, _
                               ByVal periodDates() As String, _
                               ByVal isAnnual As Boolean, _
                               ByVal startRow As Long) As Long
    Dim r As Long
    r = startRow

    Dim nDates As Long
    nDates = 0
    On Error Resume Next
    nDates = UBound(periodDates) - LBound(periodDates) + 1
    On Error GoTo 0

    ' Row 1 of section: section header
    With ws.Cells(r, COL_TAG)
        .Value = IIf(isAnnual, HDR_ANNUAL, HDR_QUARTERLY)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)   ' Blue header
        .Font.Color = RGB(255, 255, 255)       ' White text
    End With
    r = r + 1

    ' Row 2 of section: column headers
    ws.Cells(r, COL_TAG).Value = "XBRL Tag"
    ws.Cells(r, COL_UNIT).Value = "Unit"
    ws.Cells(r, COL_TAG).Font.Bold = True
    ws.Cells(r, COL_UNIT).Font.Bold = True

    ' Date column headers — ISO end-date strings directly from JSON "end" field
    Dim col As Long
    Dim i As Long
    For i = 0 To nDates - 1
        col = COL_DATA_START + i
        ws.Cells(r, col).Value = periodDates(LBound(periodDates) + i)
        ws.Cells(r, col).Font.Bold = True
        ws.Cells(r, col).HorizontalAlignment = xlRight
    Next i
    r = r + 1

    ' Row 3+: one concept per row
    If concepts.Count = 0 Or nDates = 0 Then
        WriteSection = r
        Exit Function
    End If

    ' Build a lookup: date → column index (0-based offset from COL_DATA_START)
    Dim dateToCol As Object
    Set dateToCol = CreateObject("Scripting.Dictionary")
    For i = 0 To nDates - 1
        dateToCol(periodDates(LBound(periodDates) + i)) = i
    Next i

    Dim rec As Object
    For Each rec In concepts
        ' Col A: XBRL tag — exact concept name, no renaming (PRD FR-13)
        ws.Cells(r, COL_TAG).Value = rec("ConceptName")

        ' Col B: unit string (USD, USD/shares, shares, pure, etc.)
        ws.Cells(r, COL_UNIT).Value = rec("Units")

        ' Col C+: values — ONLY write where data exists (empty = empty per FR-14)
        Dim dataDict As Object
        If isAnnual Then
            Set dataDict = rec("AnnualData")
        Else
            Set dataDict = rec("QuarterlyData")
        End If

        Dim endDate As Variant
        For Each endDate In dataDict.Keys
            Dim dateStr As String
            dateStr = CStr(endDate)

            If dateToCol.Exists(dateStr) Then
                Dim colOffset As Long
                colOffset = dateToCol(dateStr)
                col = COL_DATA_START + colOffset

                ' Get the fact Dictionary for this period
                Dim factObj As Object
                Set factObj = dataDict(dateStr)

                ' Extract val — MUST come from the "val" field (PRD FR-13)
                ' Use CDbl to handle values > Long max (AAPL reaches 416B) and EPS floats
                Dim rawVal As Variant
                rawVal = factObj("val")

                ' Write as Double — covers integers up to 2^53 and all float EPS values
                ws.Cells(r, col).Value = CDbl(rawVal)
            End If
            ' If date not in column headers (shouldn't happen), skip silently
        Next endDate

        r = r + 1
    Next rec

    WriteSection = r
End Function

'==============================================================================
' SECTION 4 — Utility: Sort Dictionary Keys
'==============================================================================

'------------------------------------------------------------------------------
' SortDictKeys
' Returns a String array of Scripting.Dictionary keys sorted ascending.
' Uses bubble sort — key count is small (< 100 periods max).
' Input dict keys must be sortable ISO date strings (YYYY-MM-DD).
'------------------------------------------------------------------------------
Private Function SortDictKeys(ByVal dict As Object) As String()
    Dim n As Long
    n = dict.Count

    Dim arr() As String
    ReDim arr(0 To n - 1)

    Dim i As Long
    i = 0
    Dim k As Variant
    For Each k In dict.Keys
        arr(i) = CStr(k)
        i = i + 1
    Next k

    ' Bubble sort ascending (ISO dates sort lexicographically = chronologically)
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

    SortDictKeys = arr
End Function

'==============================================================================
' SECTION 5 — Spot-Check Audit Utility (Phase 4 / Phase 6 verification)
'==============================================================================

'------------------------------------------------------------------------------
' GetCellAuditInfo
' Returns a string describing a specific cell's value and its expected JSON source.
' Used in TestPhase4 to verify byte-equality between cell value and SEC JSON val.
'
' Parameters:
'   wb          — workbook containing the output sheets
'   sheetName   — one of WS_INCOME_STMT, WS_BALANCE_SHEET, WS_CASH_FLOW
'   rowNum      — 1-based row
'   colNum      — 1-based column
'------------------------------------------------------------------------------
Public Function GetCellAuditInfo(ByVal wb As Workbook, _
                                 ByVal sheetName As String, _
                                 ByVal rowNum As Long, _
                                 ByVal colNum As Long) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    If ws Is Nothing Then
        GetCellAuditInfo = "Sheet not found: " & sheetName
        Exit Function
    End If

    Dim cellVal As Variant
    cellVal = ws.Cells(rowNum, colNum).Value

    Dim tagVal As Variant
    tagVal = ws.Cells(rowNum, COL_TAG).Value

    Dim unitVal As Variant
    unitVal = ws.Cells(rowNum, COL_UNIT).Value

    Dim dateHdr As Variant
    dateHdr = ws.Cells(2, colNum).Value   ' period date from row 2

    GetCellAuditInfo = "Sheet=" & sheetName & _
                       " R" & rowNum & "C" & colNum & _
                       " Tag=" & CStr(tagVal) & _
                       " Period=" & CStr(dateHdr) & _
                       " Unit=" & CStr(unitVal) & _
                       " CellValue=" & CStr(cellVal)
    On Error GoTo 0
End Function
