VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP_BAPIACPTOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public aVALUE_INDEX As Variant
Public aACTVTY_QTY As Variant
Public aDIST_KEY_QUAN As Variant
Public aACTVTY_CAPACTY As Variant
Public aDIST_KEY_CAPCTY As Variant
Public aPRICE_FIX As Variant
Public aDIST_KEY_PRICE_FIX As Variant
Public aPRICE_VAR As Variant
Public aDIST_KEY_PRICE_VAR As Variant
Public aPRICE_UNIT As Variant
Public aEQUIVALENCE_NO As Variant
Public aUNIT_ISO As Variant
Public aUNIT_OF_MEASURE As Variant
Public aCURRENCY_ISO As Variant
Public aCURRENCY As Variant
Public aOUTPUT_UNIT_ISO As Variant
Public aOUTPUT_UNIT As Variant
Public aACT_OUTPUT As Variant
Public aDIST_KEY_OUTPUT As Variant

Public Function create(oRow As Object)
    aVALUE_INDEX = oRow("VALUE_INDEX")
    aACTVTY_QTY = oRow("ACTVTY_QTY")
    aDIST_KEY_QUAN = oRow("DIST_KEY_QUAN")
    aACTVTY_CAPACTY = oRow("ACTVTY_CAPACTY")
    aDIST_KEY_CAPCTY = oRow("DIST_KEY_CAPCTY")
    aPRICE_FIX = oRow("PRICE_FIX")
    aDIST_KEY_PRICE_FIX = oRow("DIST_KEY_PRICE_FIX")
    aPRICE_VAR = oRow("PRICE_VAR")
    aDIST_KEY_PRICE_VAR = oRow("DIST_KEY_PRICE_VAR")
    aPRICE_UNIT = oRow("PRICE_UNIT")
    aEQUIVALENCE_NO = oRow("EQUIVALENCE_NO")
    aUNIT_ISO = oRow("UNIT_ISO")
    aUNIT_OF_MEASURE = oRow("UNIT_OF_MEASURE")
    aCURRENCY_ISO = oRow("CURRENCY_ISO")
    aCURRENCY = oRow("CURRENCY")
    aOUTPUT_UNIT_ISO = oRow("OUTPUT_UNIT_ISO")
    aOUTPUT_UNIT = oRow("OUTPUT_UNIT")
    aACT_OUTPUT = oRow("ACT_OUTPUT")
    aDIST_KEY_OUTPUT = oRow("DIST_KEY_OUTPUT")
End Function
