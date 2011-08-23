Attribute VB_Name = "Toolbar"
Option Explicit

'Name of package and toolbar
Public Const BarName As String = "Highlighter"

Public Sub CreateToolbar()
    Application.ScreenUpdating = False

    Dim Bar As CommandBar
    Dim Button As CommandBarButton
    Dim ButtonProperties As Variant
    Dim BtnProp As Variant

    'Set buttons
    'Example using "Import .wk1" array:
    'Array(
    '   "Import .wk1"   text displayed on the toolbar
    '   "ImportWk1"     name of the target function (see the Import module)
    '   2105            face ID (see T:\_FCSTGRP\BLM\Resources\FaceIDs.htm), set to 0 for nothing
    '   False           setting this value to True puts a separator before the button
    ')
    ButtonProperties = Array( _
        Array("Highlight", "ShowEditable", 0, True), _
        Array("Unhighlight", "HideEditable", 0, False), _
        Array("Mark editable", "MarkEditable", 0, False), _
        Array("Mark uneditable", "MarkUneditable", 0, False), _
        Array("Set color", "SetEditableColor", 0, False), _
        Array("Prepare sheet", "PrepareSheet", 0, False))

    'Delete pre-existing any toolbar with the same name
    On Error Resume Next
    CommandBars(BarName).Delete
    On Error GoTo ErrHandler

    'Add a toolbar to the top section of the user interface
    Set Bar = CommandBars.Add(BarName)
    With Bar
        .Position = msoBarTop
        .RowIndex = msoBarRowLast
        .Visible = True
    End With

    For Each BtnProp In ButtonProperties
        Set Button = Bar.Controls.Add(msoControlButton)
        With Button
            .BeginGroup = BtnProp(3)
            .Caption = BtnProp(0)
            .FaceId = BtnProp(2)
            .OnAction = BtnProp(1)
            .Style = 2 - (BtnProp(2) <> 0)
            .Visible = True
        End With
    Next BtnProp

ExitProcesses:
    'Clear object variable memory
    Set Bar = Nothing
    Set Button = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error creating toolbar!"
    Resume ExitProcesses
End Sub

Public Sub DeleteToolbar()
    'Delete toolbar if it already exists
    On Error Resume Next
    CommandBars(BarName).Delete
End Sub

Public Sub ShowToolbar(IsVisible As Boolean)
    'Show or hide toolbar if it already exists
    On Error Resume Next
    CommandBars(BarName).Visible = IsVisible
End Sub
