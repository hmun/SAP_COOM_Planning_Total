VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCOObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Costcenter As String
Public Acttype As String
Public Costelem As String
Public SCostcenter As String
Public SActtype As String
Public WBS_ELEMENT As String
Public STATKEYFIG As String

Public Function create(pCostcenter As String, pActtype As String, pCostelem As String, _
    Optional pSCostcenter As String = "", _
    Optional pSActtype As String = "", _
    Optional pWBS_ELEMENT As String = "", _
    Optional pSTATKEYFIG As String = "")
    
    Costcenter = pCostcenter
    Acttype = pActtype
    Costelem = pCostelem
    SCostcenter = pSCostcenter
    SActtype = pSActtype
    WBS_ELEMENT = pWBS_ELEMENT
    STATKEYFIG = pSTATKEYFIG
End Function

