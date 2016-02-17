VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPGetCOObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Private Sub Class_Initialize()
    IsConnected = SAPCheck()
End Sub

Public Function GetCoObjects(pType As String, pFiscy As String, pVersn As String, _
    pKokrs As String, pBukrs As String, pObjects As Collection) As String
    On Error GoTo GetCoObjects_Error
    Dim lSAPFormat As New SAPFormat
    Dim foFunction As Object
    Dim oObjects As Object
    Dim oObjectsRow As Object
    Dim lSAPCOObject As New SAPCOObject
    Dim Ret

    Set foFunction = MySAPCon.SAPFunctions.Add("ZCOPC_GET_COOBJ")
    Set oObjects = foFunction.Tables("T_OBJECTS")
    foFunction.Exports("I_TYPE") = pType
    foFunction.Exports("I_VERSN") = lSAPFormat.unpack(pVersn, 3)
    foFunction.Exports("I_GJAHR") = pFiscy
    foFunction.Exports("I_KOKRS") = pKokrs
    If Not IsNull(pBukrs) And pBukrs <> "" Then
        foFunction.Exports("I_BURKS") = pBukrs
    End If

    Ret = foFunction.call
    If Ret Then
        GetCoObjects = "Success"
        For Each oObjectsRow In oObjects.Rows
            Set lSAPCOObject = New SAPCOObject
            If Not (pType = "I" And oObjectsRow("SKOSTL") = "") Then
                lSAPCOObject.create oObjectsRow("KOSTL"), oObjectsRow("LSTAR"), oObjectsRow("KSTAR"), _
                oObjectsRow("SKOSTL"), oObjectsRow("SLSTAR"), oObjectsRow("WBS_ELEMENT")
                pObjects.Add lSAPCOObject
            End If
        Next oObjectsRow
    Else
        GetCoObjects = "Failed"
    End If
    Set foFunction = Nothing
    Exit Function
GetCoObjects_Error:
    Application.Cursor = xlDefault
    GetCoObjects = "Error"
    MySAPErr.MSGProt "SAPGetCoObject", "GetCoObjects", "", err.Number, err.Description
    Exit Function
End Function

