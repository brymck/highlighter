VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
 
Private Sub Workbook_Activate()
    ShowToolbar True
End Sub
 
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DeleteToolbar
End Sub
 
Private Sub Workbook_Deactivate()
    ShowToolbar False
End Sub
 
Private Sub Workbook_Open()
    CreateToolbar
    InitializeColors
End Sub
