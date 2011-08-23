Attribute VB_Name = "Highlight"
'Bryan McKelvey
'***************************************************************************************************
Option Explicit

Private Const FIRST_SHEET As Integer = 1
Private OnColor As Integer
Private OffColor As Integer
Enum ToggleMode
    Toggle = 0
    ForceOn = 1
    ForceOff = 2
End Enum

Public Sub InitializeColors()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    OnColor = CInt(wb.Sheets(FIRST_SHEET).Range("OnColor").Value)
    OffColor = CInt(wb.Sheets(FIRST_SHEET).Range("OffColor").Value)
    
    ' Clear object variable memory
    Set wb = Nothing
End Sub

Public Sub ToggleEditability(Optional ByVal Mode As ToggleMode = ToggleMode.Toggle)
    Dim r As Range
    Set r = ActiveCell
    
    Select Case Mode
    Case ToggleMode.Toggle
        With r.Font
            If .ColorIndex = OnColor Or .ColorIndex = OffColor Then
                .ColorIndex = 0
            Else
                .ColorIndex = OnColor
            End If
        End With
    Case ToggleMode.ForceOn
        r.Font.ColorIndex = OnColor
    Case ToggleMode.ForceOff
        r.Font.ColorIndex = 0
    Case Else
        ' Do nothing
    End Select
    
    ' Clear object variable memory
    Set r = Nothing
End Sub

Public Sub MarkEditable()
    ToggleEditability ToggleMode.ForceOn
End Sub

Public Sub MarkUneditable()
    ToggleEditability ToggleMode.ForceOff
End Sub
Private Sub SetSearchReplace(ByRef FoundFirst As Boolean, ByRef SearchColor As Integer, _
                             ByRef ReplaceColor As Integer, ByVal Mode As ToggleMode)
    Select Case Mode
    Case ToggleMode.Toggle
        FoundFirst = False
    Case ToggleMode.ForceOn
        FoundFirst = True
        SearchColor = OffColor
        ReplaceColor = OnColor
    Case ToggleMode.ForceOff
        FoundFirst = True
        SearchColor = OnColor
        ReplaceColor = OffColor
    Case Else
        ' Do nothing
    End Select
End Sub

Public Sub ToggleEditable(Optional ByVal Mode As ToggleMode = ToggleMode.Toggle)
    Dim ws As Worksheet
    Dim r As Range
    Dim FoundFirst As Boolean
    Dim SearchColor As Integer
    Dim ReplaceColor As Integer
    
    Set ws = ActiveSheet

    SetSearchReplace FoundFirst, SearchColor, ReplaceColor, Mode
    Debug.Print SearchColor
    Debug.Print ReplaceColor
    
    For Each r In ws.UsedRange.Cells
        With r.Font
            If FoundFirst Then
                If .ColorIndex = SearchColor Then
                    .ColorIndex = ReplaceColor
                End If
            Else
                If .ColorIndex = OnColor Then
                    SetSearchReplace FoundFirst, SearchColor, ReplaceColor, _
                        ToggleMode.ForceOff
                    .ColorIndex = ReplaceColor
                ElseIf .ColorIndex = OffColor Then
                    SetSearchReplace FoundFirst, SearchColor, ReplaceColor, _
                        ToggleMode.ForceOn
                    .ColorIndex = ReplaceColor
                End If
            End If
        End With
    Next
    
    ' Clear object variable memory
    Set ws = Nothing
    Set r = Nothing
End Sub

Public Sub ShowEditable()
    ToggleEditable ToggleMode.ForceOn
End Sub

Public Sub HideEditable()
    ToggleEditable ToggleMode.ForceOff
End Sub

Public Sub SetEditableColor()
    Dim wb As Workbook
    Dim r As Range
    Dim CurrentIndex As Integer
    Dim Response As Variant
    
    Set wb = ThisWorkbook
    Set r = ActiveCell
    CurrentIndex = r.Font.ColorIndex
    
    ' Allow user to manually set color index, defaulting to current cell's color index
    Response = InputBox("Current cell's color index is " & CurrentIndex & ". Enter a color index " _
        & "to use globally for highlighting." & vbCrLf & vbCrLf & _
        "(Range: 1-32, Recommended: 23, Current: " & OnColor & ")", _
        "Set Editable Color", CurrentIndex)
    
    ' Set valid
    If Response >= 1 And Response <= 32 Then
        wb.Sheets(FIRST_SHEET).Range("OnColor").Value = CInt(Response)
    End If
    
    ' Clear object variable memory
    Set wb = Nothing
    Set r = Nothing
End Sub
