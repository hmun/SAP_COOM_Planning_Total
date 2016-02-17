Attribute VB_Name = "SAPMakro"
Public sPlanning As Integer

Sub SAP_COOM_ReadActivityInput()
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPGetCOObject As New SAPGetCOObject
    Dim aSAPCOObject As New SAPCOObject

    Dim aData As New Collection
    Dim aContrl As New Collection
    Dim aObjects As New Collection

    Dim aSAPUser As New SAPUser
    Dim bRetStr As String

    Dim i As Integer
    Dim aRetStr As String

    Dim aCoAre As String
    Dim aFiscy As String
    Dim aPfrom As String
    Dim aPto As String
    Dim aSVers As String
    Dim aTVers As String
    Dim aCurt As String
    Dim aCompCodes As String
    Dim aCompCodeSplit
    Dim aCompCode As Variant

    Dim aCostcenter As String
    Dim aActtype As String
    Dim aSCostcenter As String
    Dim aSActtype As String
    Dim aCostelem As String

    Application.StatusBar = ""
    Worksheets("Parameter").Activate
    aCoAre = Cells(2, 2)
    aFiscy = Cells(3, 2)
    aPfrom = Cells(4, 2)
    aPto = Cells(5, 2)
    aSVers = Cells(6, 2)
    aTVers = Cells(7, 2)
    aCurt = Cells(8, 2)
    aCompCodes = Cells(9, 2)
    If IsNull(aCoAre) Or aCoAre = "" Or _
        IsNull(aFiscy) Or aFiscy = "" Or _
        IsNull(aPfrom) Or aPfrom = "" Or _
        IsNull(aPto) Or aPto = "" Or _
        IsNull(aSVers) Or aSVers = "" Or _
        IsNull(aTVers) Or aTVers = "" Or _
        IsNull(aCurt) Or aCurt = "" Then
        MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connection to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If

    ' Read the Objects
    aCompCodeSplit = Split(aCompCodes, ";")
    For Each aCompCode In aCompCodeSplit
        aSAPGetCOObject.GetCoObjects "I", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects
        If Ret = "Failed" Then
            MsgBox "Failed to get CO-Object list!", vbCritical + vbOKOnly
            Exit Sub
        End If
    Next aCompCode
    If aObjects.Count = 0 Then
        Exit Sub
    End If
    Worksheets("AIData").Activate
    i = 1
    aRetStr = aSAPCostActivityPlanning.ReadActivityInputTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
    If aRetStr = "Success" Then
        Application.Cursor = xlWait
        Dim aSapDataRow As Object
        Do
            Set aSapDataRow = aData(i)
            Application.StatusBar = "Line: " & i & ", " & aObjects(i).Costcenter & ", " & aObjects(i).Acttype & ", " & aObjects(i).Costelem
            Cells(i + 1, 1) = aObjects(i).Costcenter
            Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
            Cells(i + 1, 3) = aObjects(i).Acttype
            Cells(i + 1, 4) = aObjects(i).SCostcenter
            Cells(i + 1, 5) = aObjects(i).SActtype

            Cells(i + 1, 6) = aSapDataRow("UNIT_OF_MEASURE")
            Cells(i + 1, 7) = CDbl(aSapDataRow("QUANTITY_FIX"))
            Cells(i + 1, 8) = aSapDataRow("DIST_KEY_QUAN_FIX")
            Cells(i + 1, 9) = CDbl(aSapDataRow("QUANTITY_VAR"))
            Cells(i + 1, 10) = aSapDataRow("DIST_KEY_QUAN_VAR")

            i = i + 1
        Loop While i <= aObjects.Count
    End If
    Cells(i + 1, 2) = aRetStr
    Application.Cursor = xlDefault
End Sub

Sub SAP_COOM_ReadActivityOutput()
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPGetCOObject As New SAPGetCOObject
    Dim aSAPCOObject As New SAPCOObject

    Dim aData As New Collection
    Dim aContrl As New Collection
    Dim aObjects As New Collection

    Dim i As Integer
    Dim aRetStr As String

    Dim aCoAre As String
    Dim aFiscy As String
    Dim aPfrom As String
    Dim aPto As String
    Dim aSVers As String
    Dim aTVers As String
    Dim aCurt As String

    Dim aCostcenter As String
    Dim aActtype As String
    Dim aSCostcenter As String
    Dim aSActtype As String
    Dim aCostelem As String
    Dim aCompCodes As String
    Dim aCompCodeSplit
    Dim aCompCode As Variant

    Application.StatusBar = ""
    Worksheets("Parameter").Activate
    aCoAre = Cells(2, 2)
    aFiscy = Cells(3, 2)
    aPfrom = Cells(4, 2)
    aPto = Cells(5, 2)
    aSVers = Cells(6, 2)
    aTVers = Cells(7, 2)
    aCurt = Cells(8, 2)
    aCompCodes = Cells(9, 2)
    If IsNull(aCoAre) Or aCoAre = "" Or _
        IsNull(aFiscy) Or aFiscy = "" Or _
        IsNull(aPfrom) Or aPfrom = "" Or _
        IsNull(aPto) Or aPto = "" Or _
        IsNull(aSVers) Or aSVers = "" Or _
        IsNull(aTVers) Or aTVers = "" Or _
        IsNull(aCurt) Or aCurt = "" Then
        MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Objects
    aCompCodeSplit = Split(aCompCodes, ";")
    For Each aCompCode In aCompCodeSplit
        Ret = aSAPGetCOObject.GetCoObjects("O", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        If Ret = "Failed" Then
            MsgBox "Failed to get CO-Object list!", vbCritical + vbOKOnly
            Exit Sub
        End If
    Next aCompCode
    If aObjects.Count = 0 Then
        Exit Sub
    End If
    Worksheets("AOData").Activate
    i = 1
    aRetStr = aSAPCostActivityPlanning.ReadActivityOutputTotS(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData, aContrl)
    ' If aRetStr = "Success" Then
    Application.Cursor = xlWait
    Dim aSapDataRow As Object
    Dim aSapContrlRow As Object
    Do
        Set aSapDataRow = aData(i)
        Set aSapContrlRow = aContrl(i)
        Application.StatusBar = "Line: " & i & ", " & aObjects(i).Costcenter & ", " & aObjects(i).Acttype & ", " & aObjects(i).Costelem
        Cells(i + 1, 1) = aObjects(i).Costcenter
        Cells(i + 1, 2) = aObjects(i).Acttype
        Cells(i + 1, 3) = aSapDataRow.aUNIT_OF_MEASURE
        Cells(i + 1, 4) = aSapDataRow.aCURRENCY

        Cells(i + 1, 5) = CDbl(aSapDataRow.aACTVTY_QTY)
        Cells(i + 1, 6) = aSapDataRow.aDIST_KEY_QUAN
        Cells(i + 1, 7) = CDbl(aSapDataRow.aACTVTY_CAPACTY)
        Cells(i + 1, 8) = aSapDataRow.aDIST_KEY_CAPCTY
        Cells(i + 1, 9) = CDbl(aSapDataRow.aPRICE_FIX)
        Cells(i + 1, 10) = aSapDataRow.aDIST_KEY_PRICE_FIX
        Cells(i + 1, 11) = CDbl(aSapDataRow.aPRICE_VAR)
        Cells(i + 1, 12) = aSapDataRow.aDIST_KEY_PRICE_VAR
        Cells(i + 1, 13) = CInt(aSapDataRow.aPRICE_UNIT)
        Cells(i + 1, 14) = aSapDataRow.aEQUIVALENCE_NO

        Cells(i + 1, 15) = aSapContrlRow.aPRICE_INDICATOR
        Cells(i + 1, 16) = aSapContrlRow.aSWITCH_LAYOUT
        Cells(i + 1, 17) = aSapContrlRow.aATTRIB_INDEX
        Cells(i + 1, 18) = aSapDataRow.aVALUE_INDEX
        i = i + 1
    Loop While i <= aObjects.Count
    ' End If
    Cells(i + 1, 2) = aRetStr
    Application.Cursor = xlDefault
End Sub

Sub SAP_COOM_ReadPrimCost()
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPGetCOObject As New SAPGetCOObject
    Dim aSAPCOObject As New SAPCOObject

    Dim aData As New Collection
    Dim aObjects As New Collection

    Dim i As Integer
    Dim aRetStr As String

    Dim aCoAre As String
    Dim aFiscy As String
    Dim aPfrom As String
    Dim aPto As String
    Dim aSVers As String
    Dim aTVers As String
    Dim aCurt As String
    Dim aCompCodes As String
    Dim aCompCodeSplit
    Dim aCompCode As Variant

    Dim aCostcenter As String
    Dim aActtype As String
    Dim aCostelem As String

    Application.StatusBar = ""
    Worksheets("Parameter").Activate
    aCoAre = Cells(2, 2)
    aFiscy = Cells(3, 2)
    aPfrom = Cells(4, 2)
    aPto = Cells(5, 2)
    aSVers = Cells(6, 2)
    aTVers = Cells(7, 2)
    aCurt = Cells(8, 2)
    aCompCodes = Cells(9, 2)
    If IsNull(aCoAre) Or aCoAre = "" Or _
        IsNull(aFiscy) Or aFiscy = "" Or _
        IsNull(aPfrom) Or aPfrom = "" Or _
        IsNull(aPto) Or aPto = "" Or _
        IsNull(aSVers) Or aSVers = "" Or _
        IsNull(aTVers) Or aTVers = "" Or _
        IsNull(aCurt) Or aCurt = "" Then
        MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Objects
    aCompCodeSplit = Split(aCompCodes, ";")
    For Each aCompCode In aCompCodeSplit
        aSAPGetCOObject.GetCoObjects "P", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects
    Next aCompCode
    If aObjects.Count = 0 Then
        Exit Sub
    End If
    Worksheets("PData").Activate
    i = 1
    aRetStr = aSAPCostActivityPlanning.ReadPrimCostTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
    If aRetStr = "Success" Then
        Application.Cursor = xlWait
        Dim aSapDataRow As Object
        Do
            Set aSapDataRow = aData(i)
            Application.StatusBar = "Line: " & i & ", " & aObjects(i).Costcenter & ", " & aObjects(i).WBS_ELEMENT & ", " & aObjects(i).Acttype & ", " & aObjects(i).Costelem
            Cells(i + 1, 1) = aObjects(i).Costcenter
            Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
            Cells(i + 1, 3) = aObjects(i).Acttype
            Cells(i + 1, 4) = aObjects(i).Costelem
            Cells(i + 1, 5) = aSapDataRow("TRANS_CURR")
            Cells(i + 1, 6) = CDbl(aSapDataRow("FIX_VALUE"))
            Cells(i + 1, 7) = aSapDataRow("DIST_KEY_FIX_VAL")
            Cells(i + 1, 8) = CDbl(aSapDataRow("VAR_VALUE"))
            Cells(i + 1, 9) = aSapDataRow("DIST_KEY_VAR_VAL")
            Cells(i + 1, 10) = CDbl(aSapDataRow("FIX_QUAN"))
            Cells(i + 1, 11) = aSapDataRow("DIST_KEY_FIX_QUAN")
            Cells(i + 1, 12) = CDbl(aSapDataRow("VAR_QUAN"))
            Cells(i + 1, 13) = aSapDataRow("DIST_KEY_VAR_QUAN")
            i = i + 1
        Loop While i <= aObjects.Count
    End If
    Cells(i + 1, 2) = aRetStr
    Application.Cursor = xlDefault
End Sub

Sub SAP_COOM_PostPrimCost()
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPCOObject As New SAPCOObject

    Dim aData As New Collection
    Dim aDataRow As New Collection
    Dim aObjects As New Collection

    Dim i As Integer
    Dim aRetStr As String

    Dim aCoAre As String
    Dim aFiscy As String
    Dim aPfrom As String
    Dim aPto As String
    Dim aSVers As String
    Dim aTVers As String
    Dim aCurt As String

    Dim aCostcenter As String
    Dim aActtype As String
    Dim aCostelem As String
    Dim aWBSElement As String
    Dim aVal

    Application.StatusBar = ""
    Worksheets("Parameter").Activate
    aCoAre = Cells(2, 2)
    aFiscy = Cells(3, 2)
    aPfrom = Cells(4, 2)
    aPto = Cells(5, 2)
    aSVers = Cells(6, 2)
    aTVers = Cells(7, 2)
    aCurt = Cells(8, 2)
    If IsNull(aCoAre) Or aCoAre = "" Or _
        IsNull(aFiscy) Or aFiscy = "" Or _
        IsNull(aPfrom) Or aPfrom = "" Or _
        IsNull(aPto) Or aPto = "" Or _
        IsNull(aSVers) Or aSVers = "" Or _
        IsNull(aTVers) Or aTVers = "" Or _
        IsNull(aCurt) Or aCurt = "" Then
        MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Data
    Worksheets("PData").Activate
    i = 2
    Do
        Set aSAPCOObject = New SAPCOObject
        aCostcenter = Cells(i, 1)
        aWBSElement = Cells(i, 2).Value
        aActtype = Cells(i, 3)
        aCostelem = Cells(i, 4)
        aSAPCOObject.create aCostcenter, aActtype, aCostelem, "", "", aWBSElement
        aObjects.Add aSAPCOObject
        Set aDataRow = New Collection
        For J = 6 To 13
            aVal = Cells(i, J)
            aDataRow.Add aVal
        Next J
        aData.Add aDataRow
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
    aRetStr = aSAPCostActivityPlanning.PostPrimCostTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
    Cells(i, 2) = aRetStr
    If aRetStr = "Success" Then
    End If
    Application.Cursor = xlDefault
End Sub

Sub SAP_COOM_PostActivityOutput()
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPCOObject As New SAPCOObject

    Dim aData As New Collection
    Dim aContrl As New Collection
    Dim aDataRow As New Collection
    Dim aContrlRow As New Collection
    Dim aObjects As New Collection

    Dim i As Integer
    Dim aRetStr As String

    Dim aCoAre As String
    Dim aFiscy As String
    Dim aPfrom As String
    Dim aPto As String
    Dim aSVers As String
    Dim aTVers As String
    Dim aCurt As String

    Dim aCostcenter As String
    Dim aActtype As String
    Dim aCostelem As String
    Dim aVal

    Application.StatusBar = ""
    Worksheets("Parameter").Activate
    aCoAre = Cells(2, 2)
    aFiscy = Cells(3, 2)
    aPfrom = Cells(4, 2)
    aPto = Cells(5, 2)
    aSVers = Cells(6, 2)
    aTVers = Cells(7, 2)
    aCurt = Cells(8, 2)
    If IsNull(aCoAre) Or aCoAre = "" Or _
        IsNull(aFiscy) Or aFiscy = "" Or _
        IsNull(aPfrom) Or aPfrom = "" Or _
        IsNull(aPto) Or aPto = "" Or _
        IsNull(aSVers) Or aSVers = "" Or _
        IsNull(aTVers) Or aTVers = "" Or _
        IsNull(aCurt) Or aCurt = "" Then
        MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Data
    Worksheets("AOData").Activate
    i = 2
    Do
        Set aSAPCOObject = New SAPCOObject
        aCostcenter = Cells(i, 1)
        aActtype = Cells(i, 2)
        aSAPCOObject.create aCostcenter, aActtype, ""
        aObjects.Add aSAPCOObject
        Set aDataRow = New Collection
        For J = 3 To 14
            aVal = Cells(i, J)
            aDataRow.Add aVal
        Next J
        aData.Add aDataRow
        Set aContrlRow = New Collection
        For J = 15 To 16
            aVal = Cells(i, J)
            aContrlRow.Add aVal
        Next J
        aContrl.Add aContrlRow
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
    aRetStr = aSAPCostActivityPlanning.PostActivityOutputTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData, aContrl)
    Cells(i, 2) = aRetStr
    If aRetStr = "Success" Then
    End If
    Application.Cursor = xlDefault
End Sub

Sub SAP_COOM_PostActivityInput()
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPCOObject As New SAPCOObject

    Dim aData As New Collection
    Dim aDataRow As New Collection
    Dim aObjects As New Collection

    Dim i As Integer
    Dim aRetStr As String

    Dim aCoAre As String
    Dim aFiscy As String
    Dim aPfrom As String
    Dim aPto As String
    Dim aSVers As String
    Dim aTVers As String
    Dim aCurt As String

    Dim aCostcenter As String
    Dim aWBSElement As String
    Dim aActtype As String
    Dim aCostelem As String
    Dim aSCostcenter As String
    Dim aSActtype As String
    Dim aVal

    Application.StatusBar = ""
    Worksheets("Parameter").Activate
    aCoAre = Cells(2, 2)
    aFiscy = Cells(3, 2)
    aPfrom = Cells(4, 2)
    aPto = Cells(5, 2)
    aSVers = Cells(6, 2)
    aTVers = Cells(7, 2)
    aCurt = Cells(8, 2)
    If IsNull(aCoAre) Or aCoAre = "" Or _
        IsNull(aFiscy) Or aFiscy = "" Or _
        IsNull(aPfrom) Or aPfrom = "" Or _
        IsNull(aPto) Or aPto = "" Or _
        IsNull(aSVers) Or aSVers = "" Or _
        IsNull(aTVers) Or aTVers = "" Or _
        IsNull(aCurt) Or aCurt = "" Then
        MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Data
    Worksheets("AIData").Activate
    i = 2
    Do
        Set aSAPCOObject = New SAPCOObject
        aCostcenter = Cells(i, 1)
        aWBSElement = Cells(i, 2).Value
        aActtype = Cells(i, 3)
        aSCostcenter = Cells(i, 4)
        aSActtype = Cells(i, 5)
        aSAPCOObject.create aCostcenter, aActtype, "", aSCostcenter, aSActtype, aWBSElement
        aObjects.Add aSAPCOObject
        Set aDataRow = New Collection
        For J = 6 To 10
            aVal = Cells(i, J)
            aDataRow.Add aVal
        Next J
        aData.Add aDataRow
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
    aRetStr = aSAPCostActivityPlanning.PostActivityInputTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
    Cells(i, 2) = aRetStr
    If aRetStr = "Success" Then
    End If
    Application.Cursor = xlDefault
End Sub

Sub SAP_COOM_ReadKeyFigure()
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPGetCOObject As New SAPGetCOObject
    Dim aSAPCOObject As New SAPCOObject

    Dim aData As New Collection
    Dim aObjects As New Collection

    Dim i As Integer
    Dim aRetStr As String

    Dim aCoAre As String
    Dim aFiscy As String
    Dim aPfrom As String
    Dim aPto As String
    Dim aSVers As String
    Dim aTVers As String
    Dim aCurt As String
    Dim aCompCodes As String
    Dim aCompCodeSplit
    Dim aCompCode As Variant

    Dim aCostcenter As String
    Dim aActtype As String
    Dim aCostelem As String

    Application.StatusBar = ""
    Worksheets("Parameter").Activate
    aCoAre = Cells(2, 2)
    aFiscy = Cells(3, 2)
    aPfrom = Cells(4, 2)
    aPto = Cells(5, 2)
    aSVers = Cells(6, 2)
    aTVers = Cells(7, 2)
    aCurt = Cells(8, 2)
    aCompCodes = Cells(9, 2)
    If IsNull(aCoAre) Or aCoAre = "" Or _
        IsNull(aFiscy) Or aFiscy = "" Or _
        IsNull(aPfrom) Or aPfrom = "" Or _
        IsNull(aPto) Or aPto = "" Or _
        IsNull(aSVers) Or aSVers = "" Or _
        IsNull(aTVers) Or aTVers = "" Or _
        IsNull(aCurt) Or aCurt = "" Then
        MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Objects
    aCompCodeSplit = Split(aCompCodes, ";")
    For Each aCompCode In aCompCodeSplit
        '   TODO change that to read the objects with key-figure plan
        aSAPGetCOObject.GetCoObjects "O", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects
    Next aCompCode
    If aObjects.Count = 0 Then
        Exit Sub
    End If
    Worksheets("SKData").Activate
    i = 1
    aRetStr = aSAPCostActivityPlanning.ReadKeyFigureTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
    If aRetStr = "Success" Then
        Application.Cursor = xlWait
        Dim aSapDataRow As Object
        Do
            Set aSapDataRow = aData(i)
            Application.StatusBar = "Line: " & i & ", " & aObjects(i).Costcenter & ", " & aObjects(i).WBS_ELEMENT & ", " & aObjects(i).Acttype & ", " & aObjects(i).Costelem
            Cells(i + 1, 1) = aObjects(i).Costcenter
            Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
            Cells(i + 1, 3) = aObjects(i).Acttype
            Cells(i + 1, 4) = aSapDataRow("STATKEYFIG")
            Cells(i + 1, 5) = aSapDataRow("UNIT_OF_MEASURE")
            Cells(i + 1, 6) = CDbl(aSapDataRow("QUANTITY_PER01"))
            Cells(i + 1, 7) = CDbl(aSapDataRow("QUANTITY_PER02"))
            Cells(i + 1, 8) = CDbl(aSapDataRow("QUANTITY_PER03"))
            Cells(i + 1, 9) = CDbl(aSapDataRow("QUANTITY_PER04"))
            Cells(i + 1, 10) = CDbl(aSapDataRow("QUANTITY_PER05"))
            Cells(i + 1, 11) = CDbl(aSapDataRow("QUANTITY_PER06"))
            Cells(i + 1, 12) = CDbl(aSapDataRow("QUANTITY_PER07"))
            Cells(i + 1, 13) = CDbl(aSapDataRow("QUANTITY_PER08"))
            Cells(i + 1, 14) = CDbl(aSapDataRow("QUANTITY_PER09"))
            Cells(i + 1, 15) = CDbl(aSapDataRow("QUANTITY_PER010"))
            Cells(i + 1, 16) = CDbl(aSapDataRow("QUANTITY_PER011"))
            Cells(i + 1, 17) = CDbl(aSapDataRow("QUANTITY_PER012"))
            i = i + 1
        Loop While i <= aObjects.Count
    End If
    Cells(i + 1, 2) = aRetStr
    Application.Cursor = xlDefault
End Sub

Sub SAP_COOM_PostKeyFigure()
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPCOObject As New SAPCOObject

    Dim aData As New Collection
    Dim aDataRow As New Collection
    Dim aObjects As New Collection

    Dim i As Integer
    Dim aRetStr As String

    Dim aCoAre As String
    Dim aFiscy As String
    Dim aPfrom As String
    Dim aPto As String
    Dim aSVers As String
    Dim aTVers As String
    Dim aCurt As String

    Dim aCostcenter As String
    Dim aActtype As String
    Dim aCostelem As String
    Dim aWBSElement As String
    Dim aVal

    Application.StatusBar = ""
    Worksheets("Parameter").Activate
    aCoAre = Cells(2, 2)
    aFiscy = Cells(3, 2)
    aPfrom = Cells(4, 2)
    aPto = Cells(5, 2)
    aSVers = Cells(6, 2)
    aTVers = Cells(7, 2)
    aCurt = Cells(8, 2)
    If IsNull(aCoAre) Or aCoAre = "" Or _
        IsNull(aFiscy) Or aFiscy = "" Or _
        IsNull(aPfrom) Or aPfrom = "" Or _
        IsNull(aPto) Or aPto = "" Or _
        IsNull(aSVers) Or aSVers = "" Or _
        IsNull(aTVers) Or aTVers = "" Or _
        IsNull(aCurt) Or aCurt = "" Then
        MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Data
    Worksheets("SKData").Activate
    i = 2
    Do
        Set aSAPCOObject = New SAPCOObject
        aCostcenter = Cells(i, 1)
        aWBSElement = Cells(i, 2).Value
        aActtype = Cells(i, 3)
        aSAPCOObject.create aCostcenter, aActtype, "", "", "", aWBSElement
        aObjects.Add aSAPCOObject
        Set aDataRow = New Collection
        aDataRow.Add Cells(i, 4).Value
        For J = 6 To 17
            aVal = Cells(i, J)
            aDataRow.Add aVal
        Next J
        aData.Add aDataRow
        i = i + 1
    Loop While (Not IsNull(Cells(i, 1)) And Cells(i, 1) <> "") Or (Not IsNull(Cells(i, 2)) And Cells(i, 2) <> "")
    aRetStr = aSAPCostActivityPlanning.PostKeyFigure(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
    Cells(i, 2) = aRetStr
    If aRetStr = "Success" Then
    End If
    Application.Cursor = xlDefault
End Sub
