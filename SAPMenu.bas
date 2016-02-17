Attribute VB_Name = "SAPMenu"
Function delSAPCommandbar()
    Dim aCmdBars As CommandBars
    Dim aCmdBar As CommandBar
    Dim aCmdBarExists As Boolean
    Set aCmdBars = Application.CommandBars
    For Each aCmdBar In aCmdBars
      If aCmdBar.Name = "SAPCOOMPlanning" Then
        aCmdBarExists = True
        Exit For
      End If
    Next
    If aCmdBarExists Then
      aCmdBar.Delete
    End If
End Function

Function addSAPCommandbar()
Attribute addSAPCommandbar.VB_Description = "Makro am 8/12/2008 von Hermann Mundprecht aufgezeichnet"
Attribute addSAPCommandbar.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim aCmdBars As CommandBars
    Dim aCmdBar As CommandBar
    Dim aCmdBarExists As Boolean
    Set aCmdBars = Application.CommandBars
    For Each aCmdBar In aCmdBars
      If aCmdBar.Name = "SAPCOOMPlanning" Then
        aCmdBarExists = True
        Exit For
      End If
    Next
    If aCmdBarExists Then
      aCmdBar.Visible = True
    Else
      Set aCmdBar = aCmdBars.Add("SAPCOOMPlanning", msoBarTop, , True)
        Dim aButton As CommandBarControl
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .Style = msoButtonCaption
            .Caption = "AO Read"
            .TooltipText = "Read Activity Output Planning from SAP"
            .OnAction = "SAP_COOM_ReadActivityOutput"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .Style = msoButtonCaption
            .Caption = "PC Read"
            .TooltipText = "Read Primary Cost Planning from SAP"
            .OnAction = "SAP_COOM_ReadPrimCost"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .Style = msoButtonCaption
            .Caption = "AI Read"
            .TooltipText = "Read Activity Input Planning from SAP"
            .OnAction = "SAP_COOM_ReadActivityInput"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .Style = msoButtonCaption
            .Caption = "SK Read"
            .TooltipText = "Read Statistical Key Figure Planning from SAP"
            .OnAction = "SAP_COOM_ReadKeyFigure"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .BeginGroup = True
            .Style = msoButtonCaption
            .Caption = "AO Post"
            .TooltipText = "Post Activity Output Planning to SAP"
            .OnAction = "SAP_COOM_PostActivityOutput"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .Style = msoButtonCaption
            .Caption = "PC Post"
            .TooltipText = "Post Primary Cost Planning to SAP"
            .OnAction = "SAP_COOM_PostPrimCost"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .Style = msoButtonCaption
            .Caption = "AI Post"
            .TooltipText = "Post Activity Input Planning to SAP"
            .OnAction = "SAP_COOM_PostActivityInput"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .Style = msoButtonCaption
            .Caption = "SK Post"
            .TooltipText = "Post Statistical Key Figure Planning to SAP"
            .OnAction = "SAP_COOM_PostKeyFigure"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .BeginGroup = True
            .Style = msoButtonCaption
            .Caption = "Logoff"
            .TooltipText = "Logoff from SAP"
            .OnAction = "SAPLogoff"
        End With
        aCmdBar.Visible = True
    End If
End Function
