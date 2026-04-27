Attribute VB_Name = "modExcelWriter"
'==============================================================================
' modExcelWriter  -  Create / clear three worksheets and write XBRL data
' PRD S4.4 | Phase 4
'
' LAYOUT (same pattern for IS, BS, CFS tabs  -  PRD S4.4):
'
'   Row 1   Col A: "=== ANNUAL (10-K) ==="
'   Row 2   Col A: "XBRL Tag"  Col B: "Unit"  Col C+: period end-dates (ISO, sorted asc)
'   Row 3+  Col A: conceptName  Col B: unit  Col C+: val (Double) or empty
'   [blank row]
'   Row N   Col A: "=== QUARTERLY (10-Q) ==="
'   Row N+1 Col A: "XBRL Tag"  Col B: "Unit"  Col C+: quarterly end-dates (sorted asc)
'   Row N+2+ same concept rows, quarterly values
'
' DATA INTEGRITY (PRD S4.6):
'   - Every numeric cell value = CDbl(fact["val"]) from SEC JSON  -  no transformation
'   - Empty cells STAY EMPTY  -  no zero-fill, no interpolation (FR-14)
'   - Concept names written exactly as XBRL tag  -  no renaming (FR-13)
'   - String column headers = ISO end-date strings from JSON "end" field
'
' NUMERIC HANDLING:
'   - VBA Long max = 2,147,483,647. AAPL values reach 416,161,000,000 -> overflow
'   - SOLUTION: ALL numeric vals written as CDbl() -> VBA Double (53-bit mantissa)
'   - Double can represent integers up to 2^53 = 9,007,199,254,740,992 exactly
'   - EPS floats (e.g. 4.04) also handled correctly by Double
'==============================================================================
Option Explicit

'==============================================================================
' SECTION 1  -  Worksheet Management
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
' If a sheet with the given name exists -> clear all cells and formatting.
' If it doesn't exist -> create it.
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
' SECTION 2  -  Main Write Orchestrator
'==============================================================================

'------------------------------------------------------------------------------
' WriteAllSheets
' Top-level writer called from modMain.PullSECFinancials after classification.
' Writes data to all three sheets.
'
' Parameters:
'   wb             -  target workbook (ThisWorkbook for .xlam callers)
'   isCollection   -  Collection of IS concept Dicts from modClassifier
'   bsCollection   -  Collection of BS concept Dicts
'   cfsCollection  -  Collection of CFS concept Dicts
'   ticker         -  for sheet tab metadata (not written to cells  -  future use)
'------------------------------------------------------------------------------
Public Sub WriteAllSheets(ByVal wb As Workbook, _
                          ByVal isCollection As Collection, _
                          ByVal bsCollection As Collection, _
                          ByVal cfsCollection As Collection, _
                          ByVal ticker As String)
    ' Create / clear sheets
    ShowProgress "Initializing worksheets..."
    InitWorksheets wb

    ' Write each sheet  -  use modProgress constants (PRD FR-3)
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
' SECTION 3  -  Single-Sheet Writer
'==============================================================================

'------------------------------------------------------------------------------
' WriteSheet
' Writes annual + quarterly sections to one worksheet.
'
' Algorithm:
'   1. Collect ALL unique annual end-dates across all concepts -> sort -> col headers
'   2. Write annual section (rows 1..N)
'   3. Write blank gap row
'   4. Collect ALL unique quarterly end-dates -> sort -> col headers
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
    Dim allAnnualDates As Object   ' Scripting.Dictionary: date -> True (used as set)
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
'   ws            -  target worksheet
'   concepts      -  Collection of concept Dicts
'   periodDates   -  sorted array of end-date strings (column headers)
'   isAnnual      -  True -> write annual data; False -> quarterly
'   startRow      -  row to begin writing at
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

    ' Date column headers  -  ISO end-date strings directly from JSON "end" field
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

    ' Build a lookup: date -> column index (0-based offset from COL_DATA_START)
    Dim dateToCol As Object
    Set dateToCol = CreateObject("Scripting.Dictionary")
    For i = 0 To nDates - 1
        dateToCol(periodDates(LBound(periodDates) + i)) = i
    Next i

    Dim rec As Object
    For Each rec In concepts
        ' Col A: XBRL tag  -  exact concept name, no renaming (PRD FR-13)
        ws.Cells(r, COL_TAG).Value = rec("ConceptName")

        ' Col B: unit string (USD, USD/shares, shares, pure, etc.)
        ws.Cells(r, COL_UNIT).Value = rec("Units")

        ' Col C+: values  -  ONLY write where data exists (empty = empty per FR-14)
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

                ' Extract val  -  MUST come from the "val" field (PRD FR-13)
                ' Use CDbl to handle values > Long max (AAPL reaches 416B) and EPS floats
                Dim rawVal As Variant
                rawVal = factObj("val")

                ' Write as Double  -  covers integers up to 2^53 and all float EPS values
                ws.Cells(r, col).Value = CDbl(rawVal)
            End If
            ' If date not in column headers (shouldn't happen), skip silently
        Next endDate

        r = r + 1
    Next rec

    WriteSection = r
End Function

'==============================================================================
' SECTION 4  -  Utility: Sort Dictionary Keys
'==============================================================================

'------------------------------------------------------------------------------
' SortDictKeys
' Returns a String array of Scripting.Dictionary keys sorted ascending.
' Uses bubble sort  -  key count is small (< 100 periods max).
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
' SECTION 5  -  Spot-Check Audit Utility (Phase 4 / Phase 6 verification)
'==============================================================================

'------------------------------------------------------------------------------
' GetCellAuditInfo  (Phase 6 / TestPhase6 version)
' Searches the given worksheet for a concept row and period column, and
' returns the cell address and value at their intersection.
'
' The worksheet has this layout (written by WriteSection):
'   Section header row: "=== ANNUAL (10-K) ===" or "=== QUARTERLY (10-Q) ==="
'   Column-header row:  "XBRL Tag" | "Unit" | date1 | date2 | ...
'   Data rows:          conceptName | unit | val1 | val2 | ...
'
' Parameters:
'   wb           -  workbook containing the output sheets
'   sheetName    -  WS_INCOME_STMT / WS_BALANCE_SHEET / WS_CASH_FLOW
'   isAnnual     -  True -> search the ANNUAL section; False -> QUARTERLY section
'   conceptName  -  exact XBRL tag string (e.g. "NetIncomeLoss")
'   periodEnd    -  ISO end-date string (e.g. "2025-09-27")
'   cellAddr     -  OUTPUT: Excel cell address (e.g. "D5") or "" if not found
'   cellVal      -  OUTPUT: variant cell value or Empty if not found
'------------------------------------------------------------------------------
Public Sub GetCellAuditInfo(ByVal wb As Workbook, _
                             ByVal sheetName As String, _
                             ByVal isAnnual As Boolean, _
                             ByVal conceptName As String, _
                             ByVal periodEnd As String, _
                             ByRef cellAddr As String, _
                             ByRef cellVal As Variant)
    cellAddr = ""
    cellVal = Empty

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Determine which section header to look for
    Dim targetHeader As String
    targetHeader = IIf(isAnnual, HDR_ANNUAL, HDR_QUARTERLY)

    ' --- 1. Find the section header row ------------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TAG).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    Dim sectionHeaderRow As Long
    sectionHeaderRow = 0
    Dim r As Long
    For r = 1 To lastRow
        If CStr(ws.Cells(r, COL_TAG).Value) = targetHeader Then
            sectionHeaderRow = r
            Exit For
        End If
    Next r
    If sectionHeaderRow = 0 Then Exit Sub   ' section not found

    ' --- 2. Find the column-header row (immediately after section header) --
    Dim colHdrRow As Long
    colHdrRow = sectionHeaderRow + 1

    ' --- 3. Find the column for periodEnd in the header row ----------------
    Dim targetCol As Long
    targetCol = 0
    Dim lastCol As Long
    lastCol = ws.Cells(colHdrRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = COL_DATA_START To lastCol
        If CStr(ws.Cells(colHdrRow, c).Value) = periodEnd Then
            targetCol = c
            Exit For
        End If
    Next c
    If targetCol = 0 Then Exit Sub   ' period column not found

    ' --- 4. Find the row for conceptName in Col A, within this section -----
    ' Data rows start at colHdrRow + 1 and run until the next section header
    ' or the last used row.
    Dim dataStartRow As Long
    dataStartRow = colHdrRow + 1

    Dim targetRow As Long
    targetRow = 0
    For r = dataStartRow To lastRow
        Dim cellContent As String
        cellContent = CStr(ws.Cells(r, COL_TAG).Value)
        ' Stop at the next section header or a blank row that precedes it
        If cellContent = HDR_ANNUAL Or cellContent = HDR_QUARTERLY Then
            If r > dataStartRow Then Exit For   ' hit next section
        End If
        If cellContent = conceptName Then
            targetRow = r
            Exit For
        End If
    Next r
    If targetRow = 0 Then Exit Sub   ' concept row not found

    ' --- 5. Return the address and value ------------------------------------
    cellAddr = ws.Cells(targetRow, targetCol).Address(False, False)
    cellVal = ws.Cells(targetRow, targetCol).Value
End Sub
