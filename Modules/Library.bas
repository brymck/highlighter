Attribute VB_Name = "Library"
Option Explicit

Private Const REPLACE_COUNT As Integer = 1

' Mimics printf for %s
Public Function printf(ByVal Str As String, ParamArray Args()) As String
    Dim Arg As Variant
    
    For Each Arg In Args
        Str = Replace(Str, "%s", Arg, , REPLACE_COUNT)
    Next Arg
    
    printf = Str
End Function
