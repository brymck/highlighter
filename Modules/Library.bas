Attribute VB_Name = "Library"
Option Explicit

' Mimics printf for %s
Public Function printf(ByVal Str As String, ParamArray Args()) As String
    Dim Arg As Variant
    
    For Each Arg In Args
        Str = Replace(Str, "%s", Arg, 1, 1)
    Next Arg
    
    printf = Str
End Function
