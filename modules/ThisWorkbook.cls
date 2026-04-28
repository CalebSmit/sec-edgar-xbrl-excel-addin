VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' ThisWorkbook  -  Workbook-level event handlers for SEC EDGAR XBRL Add-in
' PRD S4.1 FR-1 | Phase 5
'
' In a .xlam add-in, ThisWorkbook events fire when the add-in loads/unloads.
' Workbook_Open  -> registers keyboard shortcut
' Workbook_BeforeClose -> unregisters shortcut
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    ' Register Ctrl+Shift+S shortcut as fallback for users without Ribbon UI.
    ' On Error swallow to ensure the add-in NEVER pops a runtime-error dialog
    ' on Excel startup; a registration failure should not block the user.
    On Error Resume Next
    RegisterShortcut
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clean up shortcut to avoid leaving a dead key binding
    On Error Resume Next
    UnregisterShortcut
    On Error GoTo 0
End Sub
