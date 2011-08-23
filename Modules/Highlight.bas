Attribute VB_Name = "Highlight"
Option Explicit

Private Const FIRST_SHEET As Integer = 1
Private OnColor As Integer
Private OffColor As Integer
Private Const MIN_COLOR_INDEX As Integer = 1
Private Const MAX_COLOR_INDEX As Integer = 32
Enum ToggleMode
    Toggle = 0
    ForceOn = 1
    ForceOff = 2
End Enum

Private Function IsValidColor(ByVal Color As Variant)
    On Error GoTo ErrHandler
    Color = CInt(Color)
    If Color >= MIN_COLOR_INDEX And Color <= MAX_COLOR_INDEX Then
        IsValidColor = True
    Else
        IsValidColor = False
    End If

ExitProcesses:
    Exit Function
    
ErrHandler:
    IsValidColor = False
End Function

' Sets colors to contents of OnColor and OffColor (A1 and B1, respectively, on Sheet 1),
' which by default are set to 23 and 1
Public Sub InitializeColors()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    OnColor = CInt(wb.Sheets(FIRST_SHEET).Range("OnColor").Value)
    OffColor = CInt(wb.Sheets(FIRST_SHEET).Range("OffColor").Value)
    
    ' Clear object variable memory
    Set wb = Nothing
End Sub

' Toggles editability of a given cell
Public Sub ToggleEditability(Optional ByVal Mode As ToggleMode = ToggleMode.Toggle)
    Dim r As Range
    Set r = ActiveCell
    
    Select Case Mode
    Case ToggleMode.Toggle
        ' Sets to 0 (i.e. automatic color) if cell uses either on or off color,
        ' otherwise sets to on color
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

Private Sub SetSearchReplace(ByRef FoundFirst As Boolean, ByRef SearchColor As Integer, _
                             ByRef ReplaceColor As Integer, ByVal Mode As ToggleMode)
    Select Case Mode
    Case ToggleMode.Toggle
        ' Allow script to determine whether to turn on or off
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
    
    ' Prevent screen flickering from repainting
    Application.ScreenUpdating = False
    
    SetSearchReplace FoundFirst, SearchColor, ReplaceColor, Mode
    
    ' Run through active sheet's used range of cells looking for cells we might replace
    For Each r In ws.UsedRange.Cells
        With r.Font
            If FoundFirst Then
                ' Replace search colors if we've found an initial on/off color
                If .ColorIndex = SearchColor Then
                    .ColorIndex = ReplaceColor
                End If
            Else
                ' Set search and replace colors when we find the first on/off color
                ' This only executes if Mode is set to Toggle
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
    
    ' Turn repainting back on
    Application.ScreenUpdating = False
    
    ' Clear object variable memory
    Set ws = Nothing
    Set r = Nothing
End Sub

Public Sub SetEditableColor()
    Dim wb As Workbook
    Dim r As Range
    Dim CurrentIndex As Integer
    Dim DefaultIndex As Integer
    Dim Response As Variant
    
    Set wb = ThisWorkbook
    Set r = ActiveCell
    CurrentIndex = r.Font.ColorIndex
    
    ' Use current index only if it's valid
    If IsValidColor(CurrentIndex) Then
        DefaultIndex = CurrentIndex
    Else
        DefaultIndex = OnColor
    End If
    
    ' Allow user to manually set color index, defaulting to current cell's color index
    Response = InputBox(printf("Current cell's color index is %s. Enter a color index " _
                   & "to use globally for highlighting." & vbCrLf & vbCrLf & _
                   "(Range: %s-%s, Recommended: 23, Current: %s)", _
                   CurrentIndex, MIN_COLOR_INDEX, MAX_COLOR_INDEX, OnColor), _
               "Set Editable Color", DefaultIndex)
    
    ' Set OnColor if user's response is valid
    If IsValidColor(Response) Then
        wb.Sheets(FIRST_SHEET).Range("OnColor").Value = CInt(Response)
    End If
    
    ' Clear object variable memory
    Set wb = Nothing
    Set r = Nothing
End Sub

' Force editability
Public Sub MarkEditable()
    ToggleEditability ToggleMode.ForceOn
End Sub
Public Sub MarkUneditable()
    ToggleEditability ToggleMode.ForceOff
End Sub

' Force editable cell display
Public Sub ShowEditable()
    ToggleEditable ToggleMode.ForceOn
End Sub
Public Sub HideEditable()
    ToggleEditable ToggleMode.ForceOff
End Sub
