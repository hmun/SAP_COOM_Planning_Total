VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCostActivityPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Private Sub Class_Initialize()
  IsConnected = SAPCheck()
End Sub

Public Function ReadPrimCost(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection) As String
On Error GoTo ReadPrimCost_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oPervalue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Coobject")
  Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Pervalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aPValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
      aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    End If
    If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
      aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
    End If
    If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
      aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
'      aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
    End If
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aPValRow = oPervalue.Rows.Add
    aPValRow.Value("VALUE_INDEX") = lCnt
    aPValRow.Value("COST_ELEM") = lSAPFormat.unpack(aObjRow.Costelem, 10)
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.ReadPrimCost Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        Pervalue:=oPervalue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    ReadPrimCost = "Success"
    For Each aPValRow In oPervalue.Rows
      pData.Add aPValRow
    Next aPValRow
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      ReadPrimCost = ReadPrimCost & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
ReadPrimCost_Error:
  Application.Cursor = xlDefault
  ReadPrimCost = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadPrimCost", "", err.Number, err.Description
  Exit Function
End Function

Public Function ReadPrimCostTot(pCoAre As String, pFiscy As String, pPfrom As String, _
                                pPto As String, pVers As String, pCurt As String, _
                                pObjects As Collection, pData As Collection) As String
On Error GoTo ReadPrimCostTot_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oTotValue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Coobject")
  Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadPrimCost", "Totvalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aTValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
      aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    End If
    If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
      aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
    End If
    If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
      aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
'      aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
    End If
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aTValRow = oTotValue.Rows.Add
    aTValRow.Value("VALUE_INDEX") = lCnt
    aTValRow.Value("COST_ELEM") = lSAPFormat.unpack(aObjRow.Costelem, 10)
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.ReadPrimCost Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        TotValue:=oTotValue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    ReadPrimCostTot = "Success"
    For Each aTValRow In oTotValue.Rows
      pData.Add aTValRow
    Next aTValRow
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      ReadPrimCostTot = ReadPrimCostTot & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
ReadPrimCostTot_Error:
  Application.Cursor = xlDefault
  ReadPrimCostTot = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadPrimCostTot", "", err.Number, err.Description
  Exit Function
End Function

Public Function ReadActivityInput(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection) As String
On Error GoTo ReadActivityInput_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oPervalue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Coobject")
  Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Pervalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aPValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
      aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    End If
    If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
      aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
    End If
    If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
'      aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
      aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
    End If
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aPValRow = oPervalue.Rows.Add
    aPValRow.Value("VALUE_INDEX") = lCnt
    aPValRow.Value("SEND_CCTR") = lSAPFormat.unpack(aObjRow.SCostcenter, 10)
'    aPValRow.Value("SEND_ACTIVITY") = lSAPFormat.unpack(aObjRow.SActtype, 6)
    aPValRow.Value("SEND_ACTIVITY") = aObjRow.SActtype
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.ReadActivityInput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        Pervalue:=oPervalue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    ReadActivityInput = "Success"
    For Each aPValRow In oPervalue.Rows
      pData.Add aPValRow
    Next aPValRow
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      ReadActivityInput = ReadActivityInput & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
ReadActivityInput_Error:
  Application.Cursor = xlDefault
  ReadActivityInput = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadActivityInput", "", err.Number, err.Description
  Exit Function
End Function

Public Function ReadActivityInputTot(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection) As String
On Error GoTo ReadActivityInputTot_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oTotValue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Coobject")
  Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityInput", "Totvalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aTValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
      aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    End If
    If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
      aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
    End If
    If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
'      aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
      aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
    End If
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aTValRow = oTotValue.Rows.Add
    aTValRow.Value("VALUE_INDEX") = lCnt
    aTValRow.Value("SEND_CCTR") = lSAPFormat.unpack(aObjRow.SCostcenter, 10)
'    aPValRow.Value("SEND_ACTIVITY") = lSAPFormat.unpack(aObjRow.SActtype, 6)
    aTValRow.Value("SEND_ACTIVITY") = aObjRow.SActtype
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.ReadActivityInput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        TotValue:=oTotValue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    ReadActivityInputTot = "Success"
    For Each aTValRow In oTotValue.Rows
      pData.Add aTValRow
    Next aTValRow
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      ReadActivityInputTot = ReadActivityInputTot & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
ReadActivityInputTot_Error:
  Application.Cursor = xlDefault
  ReadActivityInputTot = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadActivityInputTot", "", err.Number, err.Description
  Exit Function
End Function

Public Function ReadActivityOutput(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection) As String
On Error GoTo ReadActivityOutput_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oPervalue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Coobject")
  Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Pervalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aPValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
'    aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aPValRow = oPervalue.Rows.Add
    aPValRow.Value("VALUE_INDEX") = lCnt
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.ReadActivityOutput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        Contrl:=oContrl, _
                        Pervalue:=oPervalue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    ReadActivityOutput = "Success"
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      ReadActivityOutput = ReadActivityOutput & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  For Each aPValRow In oPervalue.Rows
    pData.Add aPValRow
  Next aPValRow
  Set boCostActivityPlanning = Nothing
  Exit Function
ReadActivityOutput_Error:
  Application.Cursor = xlDefault
  ReadActivityOutput = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadActivityOutput", "", err.Number, err.Description
  Exit Function
End Function

Public Function ReadActivityOutputTotS(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection, pContrl As Collection) As String
On Error GoTo ReadActivityOutputTotS_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oTotValue As Object
Dim oContrl As Object
Dim oRETURN As Object
Dim lError As Integer
Dim lKey As String
Dim lSAP_BAPIACPTOT As New SAP_BAPIACPTOT
Dim lSAP_BAPIACPCTRL As New SAP_BAPIACPCTRL
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Coobject")
  Set oContrl = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Contrl")
  Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Totvalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aTValRow As Object
  Dim aContrlRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  lCnt = 0
  lError = False
  Application.Cursor = xlWait
  For Each aObjRow In pObjects
    oCoobject.Rows.RemoveAll
    oIndexstructure.Rows.RemoveAll
    oTotValue.Rows.RemoveAll
    oContrl.Rows.RemoveAll
    oRETURN.Rows.RemoveAll
    lCnt = lCnt + 1
    lKey = aObjRow.Costcenter & "-" & aObjRow.Acttype
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
'    aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    aIndRow.Value("ATTRIB_INDEX") = lCnt
    Set aTValRow = oTotValue.Rows.Add
    aTValRow.Value("VALUE_INDEX") = lCnt
    Set aContrlRow = oContrl.Rows.Add
    aContrlRow.Value("ATTRIB_INDEX") = lCnt
    boCostActivityPlanning.ReadActivityOutput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        TotValue:=oTotValue, _
                        Contrl:=oContrl, _
                        RETURN:=oRETURN
    If oRETURN.RowCount <> 0 Then
        Dim RetRow As Object
        For Each RetRow In oRETURN.Rows
          ReadActivityOutputTotS = ReadActivityOutputTotS & ";" & RetRow("MESSAGE")
        Next RetRow
        lError = True
    Else
        For Each aTValRow In oTotValue.Rows
          Set lSAP_BAPIACPTOT = New SAP_BAPIACPTOT
          lSAP_BAPIACPTOT.create aTValRow
          pData.Add lSAP_BAPIACPTOT, lKey
        Next aTValRow
        For Each aContrlRow In oContrl.Rows
          Set lSAP_BAPIACPCTRL = New SAP_BAPIACPCTRL
          lSAP_BAPIACPCTRL.create aContrlRow
          pContrl.Add lSAP_BAPIACPCTRL, lKey
        Next aContrlRow
    End If
  Next aObjRow
  Application.Cursor = xlDefault
  If lError = False Then
    ReadActivityOutputTotS = "Success"
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
ReadActivityOutputTotS_Error:
  Application.Cursor = xlDefault
  ReadActivityOutputTotS = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadActivityOutputTotS", "", err.Number, err.Description
  Exit Function
End Function

Public Function ReadActivityOutputTot(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection, pContrl As Collection) As String
On Error GoTo ReadActivityOutputTot_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oTotValue As Object
Dim oContrl As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Coobject")
  Set oContrl = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Contrl")
  Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Totvalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aTValRow As Object
  Dim aContrlRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
'    aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    aIndRow.Value("ATTRIB_INDEX") = lCnt
    Set aTValRow = oTotValue.Rows.Add
    aTValRow.Value("VALUE_INDEX") = lCnt
    Set aContrlRow = oContrl.Rows.Add
    aContrlRow.Value("ATTRIB_INDEX") = lCnt
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.ReadActivityOutput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        TotValue:=oTotValue, _
                        Contrl:=oContrl, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    ReadActivityOutputTot = "Success"
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      ReadActivityOutputTot = ReadActivityOutputTot & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  For Each aTValRow In oTotValue.Rows
    pData.Add aTValRow
  Next aTValRow
  For Each aContrlRow In oContrl.Rows
    pContrl.Add aContrlRow
  Next aContrlRow
  Set boCostActivityPlanning = Nothing
  Exit Function
ReadActivityOutputTot_Error:
  Application.Cursor = xlDefault
  ReadActivityOutputTot = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadActivityOutputTot", "", err.Number, err.Description
  Exit Function
End Function

Public Function PostActivityOutput(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection) As String
On Error GoTo PostActivityOutput_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oPervalue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Coobject")
  Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Pervalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aPValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  Dim aVal
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aPValRow = oPervalue.Rows.Add
    aPValRow.Value("VALUE_INDEX") = lCnt
'   move the values from the data
    Set aDataRow = pData(lCnt)
    For J = 2 To 97
      aVal = aDataRow(J - 1)
      aPValRow.Value(J) = aVal
    Next J
    aPValRow.Value(99) = aDataRow(97)
    aPValRow.Value(101) = aDataRow(98)
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.PostActivityOutput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        Pervalue:=oPervalue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    PostActivityOutput = "Success"
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    aSAPBapiTranctionCommit.commit
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      PostActivityOutput = PostActivityOutput & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
PostActivityOutput_Error:
  Application.Cursor = xlDefault
  PostActivityOutput = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "PostActivityOutput", "", err.Number, err.Description
  Exit Function
End Function


Public Function PostActivityOutputTot(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection, pContrl As Collection) As String
On Error GoTo PostActivityOutputTot_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oTotValue As Object
Dim oContrl As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Coobject")
  Set oContrl = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadActivityOutput", "Contrl")
  Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityOutput", "Totvalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aTValRow As Object
  Dim aContrlRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  Dim aVal
  Dim aDataRow As Object
  Dim aCtrlRow As Object
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    aIndRow.Value("ATTRIB_INDEX") = lCnt
    Set aTValRow = oTotValue.Rows.Add
    aTValRow.Value("VALUE_INDEX") = lCnt
    Set aContrlRow = oContrl.Rows.Add
    aContrlRow.Value("ATTRIB_INDEX") = lCnt
'   move the values from the data
    Set aDataRow = pData(lCnt)
    aTValRow.Value("UNIT_OF_MEASURE") = aDataRow(1)
    aTValRow.Value("CURRENCY") = aDataRow(2)
    aTValRow.Value("ACTVTY_QTY") = aDataRow(3)
    aTValRow.Value("DIST_KEY_QUAN") = aDataRow(4)
    aTValRow.Value("ACTVTY_CAPACTY") = aDataRow(5)
    aTValRow.Value("DIST_KEY_CAPCTY") = aDataRow(6)
    aTValRow.Value("PRICE_FIX") = aDataRow(7)
    aTValRow.Value("DIST_KEY_PRICE_FIX") = aDataRow(8)
    aTValRow.Value("PRICE_VAR") = aDataRow(9)
    aTValRow.Value("DIST_KEY_PRICE_VAR") = aDataRow(10)
    aTValRow.Value("PRICE_UNIT") = aDataRow(11)
    aTValRow.Value("EQUIVALENCE_NO") = aDataRow(12)
'   move the values from the contrl
    Set aCtrlRow = pContrl(lCnt)
    aContrlRow.Value("PRICE_INDICATOR") = lSAPFormat.unpack(aCtrlRow(1), 3)
    aContrlRow.Value("SWITCH_LAYOUT") = aCtrlRow(2)
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.PostActivityOutput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        TotValue:=oTotValue, _
                        Contrl:=oContrl, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    PostActivityOutputTot = "Success"
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    aSAPBapiTranctionCommit.commit
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      PostActivityOutputTot = PostActivityOutputTot & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
PostActivityOutputTot_Error:
  Application.Cursor = xlDefault
  PostActivityOutputTot = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "PostActivityOutputTot", "", err.Number, err.Description
  Exit Function
End Function

Public Function PostActivityInput(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection) As String
On Error GoTo PostActivityInput_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oPervalue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Coobject")
  Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Pervalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aPValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  Dim aVal
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
      aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    End If
    If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
      aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
    End If
    If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
      aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
'      aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
    End If
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aPValRow = oPervalue.Rows.Add
    aPValRow.Value("VALUE_INDEX") = lCnt
    aPValRow.Value("SEND_CCTR") = lSAPFormat.unpack(aObjRow.SCostcenter, 10)
    aPValRow.Value("SEND_ACTIVITY") = lSAPFormat.unpack(aObjRow.SActtype, 6)
'   move the values from the data
    Set aDataRow = pData(lCnt)
    For J = 6 To 37
      aVal = aDataRow(J - 5)
      aPValRow.Value(J) = aVal
    Next J
    aPValRow.Value(45) = aDataRow(33)
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.PostActivityInput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        Pervalue:=oPervalue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    PostActivityInput = "Success"
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    aSAPBapiTranctionCommit.commit
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      PostActivityInput = PostActivityInput & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
PostActivityInput_Error:
  Application.Cursor = xlDefault
  PostActivityInput = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "PostActivityInput", "", err.Number, err.Description
  Exit Function
End Function

Public Function PostActivityInputTot(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection) As String
On Error GoTo PostActivityInputTot_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oTotValue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "Coobject")
  Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostActivityInput", "TotValue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aTValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  Dim aVal
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
      aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    End If
    If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
      aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
    End If
    If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
'      aCOObjRow.Value("ACTTYPE") = lSAPFormat.unpack(aObjRow.Acttype, 6)
      aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
    End If
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aTValRow = oTotValue.Rows.Add
    aTValRow.Value("VALUE_INDEX") = lCnt
    aTValRow.Value("SEND_CCTR") = lSAPFormat.unpack(aObjRow.SCostcenter, 10)
    aTValRow.Value("SEND_ACTIVITY") = aObjRow.SActtype
'   move the values from the data
    Set aDataRow = pData(lCnt)
    aTValRow.Value("UNIT_OF_MEASURE") = aDataRow(1)
    aTValRow.Value("QUANTITY_FIX") = aDataRow(2)
    aTValRow.Value("DIST_KEY_QUAN_FIX") = aDataRow(3)
    aTValRow.Value("QUANTITY_VAR") = aDataRow(4)
    aTValRow.Value("DIST_KEY_QUAN_VAR") = aDataRow(5)
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.PostActivityInput Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        TotValue:=oTotValue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    PostActivityInputTot = "Success"
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    aSAPBapiTranctionCommit.commit
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      PostActivityInputTot = PostActivityInputTot & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
PostActivityInputTot_Error:
  Application.Cursor = xlDefault
  PostActivityInputTot = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "PostActivityInputTot", "", err.Number, err.Description
  Exit Function
End Function

Public Function PostPrimCost(pCoAre As String, pFiscy As String, pPfrom As String, _
                             pPto As String, pVers As String, pCurt As String, _
                             pObjects As Collection, pData As Collection) As String
On Error GoTo PostPrimCost_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oPervalue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Coobject")
  Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Pervalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aPValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  Dim aDataRow As Collection
  Dim aVal
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
      aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    End If
    If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
      aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
    End If
    If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
      aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
    End If
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aPValRow = oPervalue.Rows.Add
    aPValRow.Value("VALUE_INDEX") = lCnt
    aPValRow.Value("COST_ELEM") = lSAPFormat.unpack(aObjRow.Costelem, 10)
'   move the values from the data
    Set aDataRow = pData(lCnt)
    J = 8
    For Each aVal In aDataRow
      aPValRow.Value(J) = aVal
      J = J + 1
    Next aVal
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.PostPrimCost Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        Pervalue:=oPervalue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    PostPrimCost = "Success"
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    aSAPBapiTranctionCommit.commit
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      PostPrimCost = PostPrimCost & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
PostPrimCost_Error:
  Application.Cursor = xlDefault
  PostPrimCost = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "PostPrimCost", "", err.Number, err.Description
  Exit Function
End Function

Public Function PostPrimCostTot(pCoAre As String, pFiscy As String, pPfrom As String, _
                                pPto As String, pVers As String, pCurt As String, _
                                pObjects As Collection, pData As Collection) As String
On Error GoTo PostPrimCostTot_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oTotValue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Coobject")
  Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Totvalue")
  oHeaderinfo("CO_AREA") = pCoAre
  oHeaderinfo("FISC_YEAR") = pFiscy
  oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
  oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
  oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
  oHeaderinfo("PLAN_CURRTYPE") = pCurt
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aTValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  Dim aDataRow As Collection
  Dim aVal
  lCnt = 0
  For Each aObjRow In pObjects
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.Add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
      aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
    End If
    If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
      aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
    End If
    If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
      aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
    End If
    Set aIndRow = oIndexstructure.Rows.Add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aTValRow = oTotValue.Rows.Add
    aTValRow.Value("VALUE_INDEX") = lCnt
    aTValRow.Value("COST_ELEM") = lSAPFormat.unpack(aObjRow.Costelem, 10)
'   move the values from the data
    Set aDataRow = pData(lCnt)
  
    aTValRow.Value("FIX_VALUE") = aDataRow(1)
    aTValRow.Value("DIST_KEY_FIX_VAL") = aDataRow(2)
    aTValRow.Value("VAR_VALUE") = aDataRow(3)
    aTValRow.Value("DIST_KEY_VAR_VAL") = aDataRow(4)
    aTValRow.Value("FIX_QUAN") = aDataRow(5)
    aTValRow.Value("DIST_KEY_FIX_QUAN") = aDataRow(6)
    aTValRow.Value("VAR_QUAN") = aDataRow(7)
    aTValRow.Value("DIST_KEY_VAR_QUAN") = aDataRow(8)
  Next aObjRow
  Application.Cursor = xlWait
  boCostActivityPlanning.PostPrimCost Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        TotValue:=oTotValue, _
                        RETURN:=oRETURN
  Application.Cursor = xlDefault
  If oRETURN.RowCount = 0 Then
    PostPrimCostTot = "Success"
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    aSAPBapiTranctionCommit.commit
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      PostPrimCostTot = PostPrimCostTot & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
PostPrimCostTot_Error:
  Application.Cursor = xlDefault
  PostPrimCostTot = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "PostPrimCostTot", "", err.Number, err.Description
  Exit Function
End Function

Public Function PostKeyFigureTot(pCoAre As String, pFiscy As String, pPfrom As String, _
    pPto As String, pVers As String, pCurt As String, _
    pObjects As Collection, pData As Collection) As String
    On Error GoTo PostKeyFigureTot_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCostActivityPlanning As Object
    Dim oHeaderinfo As Object
    Dim oIndexstructure As Object
    Dim oCoobject As Object
    Dim oTotValue As Object
    Dim oRETURN As Object
    Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Return")
    Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Headerinfo")
    Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Indexstructure")
    Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Coobject")
    Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Totvalue")
    oHeaderinfo("CO_AREA") = pCoAre
    oHeaderinfo("FISC_YEAR") = pFiscy
    oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
    oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
    oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
    Dim aObjRow As Object
    Dim aIndRow As Object
    Dim aCOObjRow As Object
    Dim aTValRow As Object
    Dim lRow As Object
    Dim lCnt As Integer
    Dim aDataRow As Collection
    Dim aVal
    lCnt = 0
    For Each aObjRow In pObjects
        lCnt = lCnt + 1
        Set aCOObjRow = oCoobject.Rows.Add
        aCOObjRow.Value("OBJECT_INDEX") = lCnt
        If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
            aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
        End If
        If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
            aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
        End If
        If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
            aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
        End If
        Set aIndRow = oIndexstructure.Rows.Add
        aIndRow.Value("OBJECT_INDEX") = lCnt
        aIndRow.Value("VALUE_INDEX") = lCnt
        Set aTValRow = oTotValue.Rows.Add
        aTValRow.Value("VALUE_INDEX") = lCnt
        '   move the values from the data
        Set aDataRow = pData(lCnt)
        aTValRow.Value("STATKEYFIG") = lSAPFormat.unpack(aDataRow(1), 6)
        aTValRow.Value("STAT_QTY") = aDataRow(2)
        aTValRow.Value("DIST_KEY_QUAN") = aDataRow(3)
        aTValRow.Value("STAT_QTY_MAX") = aDataRow(4)
        aTValRow.Value("DIST_KEY_QUAN_MAX") = aDataRow(5)
    Next aObjRow
    Application.Cursor = xlWait
    boCostActivityPlanning.PostKeyFigure Headerinfo:=oHeaderinfo, _
    Indexstructure:=oIndexstructure, _
    Coobject:=oCoobject, _
    TotValue:=oTotValue, _
    RETURN:=oRETURN
    Application.Cursor = xlDefault
    If oRETURN.RowCount = 0 Then
        PostKeyFigureTot = "Success"
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        aSAPBapiTranctionCommit.commit
    Else
        Dim RetRow As Object
        For Each RetRow In oRETURN.Rows
            PostKeyFigureTot = PostKeyFigureTot & ";" & RetRow("MESSAGE")
        Next RetRow
    End If
    Set boCostActivityPlanning = Nothing
    Exit Function
PostKeyFigureTot_Error:
    Application.Cursor = xlDefault
    PostKeyFigureTot = "Error"
    MySAPErr.MSGProt "SAPCostActivityPlanning", "PostKeyFigureTot", "", err.Number, err.Description
    Exit Function
End Function

Public Function PostKeyFigure(pCoAre As String, pFiscy As String, pPfrom As String, _
    pPto As String, pVers As String, pCurt As String, _
    pObjects As Collection, pData As Collection) As String
    On Error GoTo PostKeyFigureTot_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCostActivityPlanning As Object
    Dim oHeaderinfo As Object
    Dim oIndexstructure As Object
    Dim oCoobject As Object
    Dim oPervalue As Object
    Dim oRETURN As Object
    Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Return")
    Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Headerinfo")
    Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Indexstructure")
    Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Coobject")
    Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostKeyFigure", "Pervalue")
    oHeaderinfo("CO_AREA") = pCoAre
    oHeaderinfo("FISC_YEAR") = pFiscy
    oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
    oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
    oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
    Dim aObjRow As Object
    Dim aIndRow As Object
    Dim aCOObjRow As Object
    Dim aPValRow As Object
    Dim lRow As Object
    Dim lCnt As Integer
    Dim aDataRow As Collection
    Dim aVal
    lCnt = 0
    For Each aObjRow In pObjects
        lCnt = lCnt + 1
        Set aCOObjRow = oCoobject.Rows.Add
        aCOObjRow.Value("OBJECT_INDEX") = lCnt
        If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
            aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
        End If
        If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
            aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
        End If
        If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
            aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
        End If
        Set aIndRow = oIndexstructure.Rows.Add
        aIndRow.Value("OBJECT_INDEX") = lCnt
        aIndRow.Value("VALUE_INDEX") = lCnt
        Set aPValRow = oPervalue.Rows.Add
        aPValRow.Value("VALUE_INDEX") = lCnt
        '   move the values from the data
        Set aDataRow = pData(lCnt)
        aPValRow.Value("STATKEYFIG") = lSAPFormat.unpack(aDataRow(1), 6)
        aPValRow.Value("QUANTITY_PER01") = CDbl(aDataRow(2))
        aPValRow.Value("QUANTITY_PER02") = CDbl(aDataRow(3))
        aPValRow.Value("QUANTITY_PER03") = CDbl(aDataRow(4))
        aPValRow.Value("QUANTITY_PER04") = CDbl(aDataRow(5))
        aPValRow.Value("QUANTITY_PER05") = CDbl(aDataRow(6))
        aPValRow.Value("QUANTITY_PER06") = CDbl(aDataRow(7))
        aPValRow.Value("QUANTITY_PER07") = CDbl(aDataRow(8))
        aPValRow.Value("QUANTITY_PER08") = CDbl(aDataRow(9))
        aPValRow.Value("QUANTITY_PER09") = CDbl(aDataRow(10))
        aPValRow.Value("QUANTITY_PER10") = CDbl(aDataRow(11))
        aPValRow.Value("QUANTITY_PER11") = CDbl(aDataRow(12))
        aPValRow.Value("QUANTITY_PER12") = CDbl(aDataRow(13))
    Next aObjRow
    Application.Cursor = xlWait
    boCostActivityPlanning.PostKeyFigure Headerinfo:=oHeaderinfo, _
    Indexstructure:=oIndexstructure, _
    Coobject:=oCoobject, _
    Pervalue:=oPervalue, _
    RETURN:=oRETURN
    Application.Cursor = xlDefault
    If oRETURN.RowCount = 0 Then
        PostKeyFigure = "Success"
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        aSAPBapiTranctionCommit.commit
    Else
        Dim RetRow As Object
        For Each RetRow In oRETURN.Rows
            PostKeyFigure = PostKeyFigure & ";" & RetRow("MESSAGE")
        Next RetRow
    End If
    Set boCostActivityPlanning = Nothing
    Exit Function
PostKeyFigureTot_Error:
    Application.Cursor = xlDefault
    PostKeyFigure = "Error"
    MySAPErr.MSGProt "SAPCostActivityPlanning", "PostKeyFigure", "", err.Number, err.Description
    Exit Function
End Function


Public Function ReadKeyFigureTot(pCoAre As String, pFiscy As String, pPfrom As String, _
    pPto As String, pVers As String, pCurt As String, _
    pObjects As Collection, pData As Collection) As String
    On Error GoTo ReadKeyFigureTot_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCostActivityPlanning As Object
    Dim oHeaderinfo As Object
    Dim oIndexstructure As Object
    Dim oCoobject As Object
    Dim oTotValue As Object
    Dim oRETURN As Object
    Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Return")
    Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Headerinfo")
    Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Indexstructure")
    Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Coobject")
    Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Totvalue")
    oHeaderinfo("CO_AREA") = pCoAre
    oHeaderinfo("FISC_YEAR") = pFiscy
    oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
    oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
    oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
    Dim aObjRow As Object
    Dim aIndRow As Object
    Dim aCOObjRow As Object
    Dim aTValRow As Object
    Dim lRow As Object
    Dim lCnt As Integer
    lCnt = 0
    For Each aObjRow In pObjects
        lCnt = lCnt + 1
        Set aCOObjRow = oCoobject.Rows.Add
        aCOObjRow.Value("OBJECT_INDEX") = lCnt
        If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
            aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
        End If
        If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
            aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
        End If
        If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
            aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
        End If
        Set aIndRow = oIndexstructure.Rows.Add
        aIndRow.Value("OBJECT_INDEX") = lCnt
        aIndRow.Value("VALUE_INDEX") = lCnt
        Set aTValRow = oTotValue.Rows.Add
        aTValRow.Value("VALUE_INDEX") = lCnt
        '    aTValRow.Value("SEND_CCTR") = lSAPFormat.unpack(aObjRow.SCostcenter, 10)
        '    aPValRow.Value("SEND_ACTIVITY") = lSAPFormat.unpack(aObjRow.SActtype, 6)
        '    aTValRow.Value("SEND_ACTIVITY") = aObjRow.SActtype
    Next aObjRow
    Application.Cursor = xlWait
    boCostActivityPlanning.ReadKeyFigure Headerinfo:=oHeaderinfo, _
    Indexstructure:=oIndexstructure, _
    Coobject:=oCoobject, _
    TotValue:=oTotValue, _
    RETURN:=oRETURN
    Application.Cursor = xlDefault
    If oRETURN.RowCount = 0 Then
        ReadKeyFigureTot = "Success"
        For Each aTValRow In oTotValue.Rows
            pData.Add aTValRow
        Next aTValRow
    Else
        Dim RetRow As Object
        For Each RetRow In oRETURN.Rows
            ReadKeyFigureTot = ReadKeyFigureTot & ";" & RetRow("MESSAGE")
        Next RetRow
    End If
    Set boCostActivityPlanning = Nothing
    Exit Function
ReadKeyFigureTot_Error:
    Application.Cursor = xlDefault
    ReadKeyFigureTot = "Error"
    MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadKeyFigureTot", "", err.Number, err.Description
    Exit Function
End Function

Public Function ReadKeyFigure(pCoAre As String, pFiscy As String, pPfrom As String, _
    pPto As String, pVers As String, pCurt As String, _
    pObjects As Collection, pData As Collection) As String
    On Error GoTo ReadKeyFigure_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCostActivityPlanning As Object
    Dim oHeaderinfo As Object
    Dim oIndexstructure As Object
    Dim oCoobject As Object
    Dim oPervalue As Object
    Dim oRETURN As Object
    Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Return")
    Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Headerinfo")
    Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Indexstructure")
    Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Coobject")
    Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "ReadKeyFigure", "Pervalue")
    oHeaderinfo("CO_AREA") = pCoAre
    oHeaderinfo("FISC_YEAR") = pFiscy
    oHeaderinfo("PERIOD_FROM") = lSAPFormat.unpack(pPfrom, 3)
    oHeaderinfo("PERIOD_TO") = lSAPFormat.unpack(pPto, 3)
    oHeaderinfo("VERSION") = lSAPFormat.unpack(pVers, 3)
    Dim aObjRow As Object
    Dim aIndRow As Object
    Dim aCOObjRow As Object
    Dim aPValRow As Object
    Dim lRow As Object
    Dim lCnt As Integer
    lCnt = 0
    For Each aObjRow In pObjects
        lCnt = lCnt + 1
        Set aCOObjRow = oCoobject.Rows.Add
        aCOObjRow.Value("OBJECT_INDEX") = lCnt
        If Not IsNull(aObjRow.Costcenter) And aObjRow.Costcenter <> "" Then
            aCOObjRow.Value("COSTCENTER") = lSAPFormat.unpack(aObjRow.Costcenter, 10)
        End If
        If Not IsNull(aObjRow.WBS_ELEMENT) And aObjRow.WBS_ELEMENT <> "" Then
            aCOObjRow.Value("WBS_ELEMENT") = aObjRow.WBS_ELEMENT
        End If
        If Not IsNull(aObjRow.Acttype) And aObjRow.Acttype <> "" Then
            aCOObjRow.Value("ACTTYPE") = aObjRow.Acttype
        End If
        Set aIndRow = oIndexstructure.Rows.Add
        aIndRow.Value("OBJECT_INDEX") = lCnt
        aIndRow.Value("VALUE_INDEX") = lCnt
        Set aPValRow = oPervalue.Rows.Add
        aPValRow.Value("VALUE_INDEX") = lCnt
        '    aTValRow.Value("SEND_CCTR") = lSAPFormat.unpack(aObjRow.SCostcenter, 10)
        '    aPValRow.Value("SEND_ACTIVITY") = lSAPFormat.unpack(aObjRow.SActtype, 6)
        '    aTValRow.Value("SEND_ACTIVITY") = aObjRow.SActtype
    Next aObjRow
    Application.Cursor = xlWait
    boCostActivityPlanning.ReadKeyFigure Headerinfo:=oHeaderinfo, _
    Indexstructure:=oIndexstructure, _
    Coobject:=oCoobject, _
    Pervalue:=oPervalue, _
    RETURN:=oRETURN
    Application.Cursor = xlDefault
    If oRETURN.RowCount = 0 Then
        ReadKeyFigure = "Success"
        For Each aPValRow In oPervalue.Rows
            pData.Add aPValRow
        Next aPValRow
    Else
        Dim RetRow As Object
        For Each RetRow In oRETURN.Rows
            ReadKeyFigure = ReadKeyFigure & ";" & RetRow("MESSAGE")
        Next RetRow
    End If
    Set boCostActivityPlanning = Nothing
    Exit Function
ReadKeyFigure_Error:
    Application.Cursor = xlDefault
    ReadKeyFigure = "Error"
    MySAPErr.MSGProt "SAPCostActivityPlanning", "ReadKeyFigure", "", err.Number, err.Description
    Exit Function
End Function

