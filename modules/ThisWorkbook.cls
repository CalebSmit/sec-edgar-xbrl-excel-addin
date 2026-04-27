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
    ' Register Ctrl+Shift+S shortcut as fallback for users without Ribbon UI
    RegisterShortcut
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clean up shortcut to avoid leaving a dead key binding
    UnregisterShortcut
End Sub
