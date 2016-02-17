VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP_BAPIACPCTRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public aATTRIB_INDEX As Variant
Public aACTIVITY_CAT As Variant
Public aACT_CAT_ACTUAL As Variant
Public aPRICE_INDICATOR As Variant
Public aACT_PRICE_IND As Variant
Public aALLOC_COST_ELEM As Variant
Public aPREDIS_FXD_COST As Variant
Public aMANUAL_ACT_QTY As Variant
Public aAVERAGE_PRICE_IND As Variant
Public aSWITCH_LAYOUT As Variant

Public Function create(oRow As Object)
    aATTRIB_INDEX = oRow("ATTRIB_INDEX")
    aACTIVITY_CAT = oRow("ACTIVITY_CAT")
    aACT_CAT_ACTUAL = oRow("ACT_CAT_ACTUAL")
    aPRICE_INDICATOR = oRow("PRICE_INDICATOR")
    aACT_PRICE_IND = oRow("ACT_PRICE_IND")
    aALLOC_COST_ELEM = oRow("ALLOC_COST_ELEM")
    aPREDIS_FXD_COST = oRow("PREDIS_FXD_COST")
    aMANUAL_ACT_QTY = oRow("MANUAL_ACT_QTY")
    aAVERAGE_PRICE_IND = oRow("AVERAGE_PRICE_IND")
    aSWITCH_LAYOUT = oRow("SWITCH_LAYOUT")
End Function

