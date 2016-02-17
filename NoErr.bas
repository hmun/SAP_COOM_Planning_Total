Attribute VB_Name = "NoErr"
Function NE(pVal) As Double
On Error GoTo NE_Error
    If IsError(pVal) Then
        NE = 0
    Else
        If IsError(pVal * 1) Then
            NE = 0
        Else
            NE = CDbl(pVal)
        End If
    End If
    Exit Function
NE_Error:
    NE = 0
End Function

Function NNA(pVal) As Variant
On Error GoTo NNA_Error
    If IsError(pVal) Then
        NNA = "-"
    Else
        NNA = pVal
    End If
    Exit Function
NNA_Error:
    NNA = ""
End Function
