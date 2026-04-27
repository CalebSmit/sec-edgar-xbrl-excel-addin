Attribute VB_Name = "modRibbon"
'==============================================================================
' modRibbon  -  Custom Ribbon callback and keyboard shortcut registration
' PRD S4.1 FR-1 | Phase 5
'
' The Ribbon button (defined in customUI/customUI14.xml) calls
' RibbonCallback_PullSECFinancials. The IRibbonControl parameter is
' required by the Office Ribbon callback contract  -  it is accepted but unused.
'
' For users who have not installed the Custom UI XML, the add-in also
' registers Ctrl+Shift+S as a keyboard shortcut via Auto_Open.
'
' VERIFIED: IRibbonControl is the correct parameter type for onAction callbacks
' in customUI14.xml (Office 2010+). The parameter must be declared as Object
' (not As IRibbonControl) to avoid requiring the Office object library reference.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' RibbonCallback_PullSECFinancials
' Called by the Ribbon button defined in customUI14.xml.
' The ribbonControl As Object parameter is required by the callback contract.
'------------------------------------------------------------------------------
Public Sub RibbonCallback_PullSECFinancials(ByVal ribbonControl As Object)
    ' Delegate to the main entry point  -  same code path as manual macro run
    PullSECFinancials
End Sub

'------------------------------------------------------------------------------
' RegisterShortcut
' Registers Ctrl+Shift+S as a fallback keyboard shortcut.
' Called from ThisWorkbook.Workbook_Open (see ThisWorkbook.cls).
' Unregistered on Workbook_BeforeClose to avoid conflict.
'------------------------------------------------------------------------------
Public Sub RegisterShortcut()
    ' Ctrl+Shift+S -> PullSECFinancials
    Dim escapedName As String
    escapedName = Replace(ThisWorkbook.Name, "'", "''")
    Application.OnKey "^+S", "'" & escapedName & "'!PullSECFinancials"
End Sub

'------------------------------------------------------------------------------
' UnregisterShortcut
' Clears the Ctrl+Shift+S shortcut registration on close.
'------------------------------------------------------------------------------
Public Sub UnregisterShortcut()
    Application.OnKey "^+S"    ' No second argument -> resets to default behaviour
End Sub
