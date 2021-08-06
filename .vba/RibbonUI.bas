Attribute VB_Name = "RibbonUI"
Option Explicit

Dim Rib        As IRibbonUI
Public MyTag   As String

'Callback for customUI.onLoad
Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set Rib = ribbon
End Sub

Sub RefreshRibbon(Tag As String)
    MyTag = Tag
    If Rib Is Nothing Then
        MsgBox "Error, Save/Restart your workbook"
    Else
        Rib.Invalidate
    End If
End Sub

Sub GetVisible(control As IRibbonControl, ByRef visible)
    If MyTag = "show" Then
        visible = True
    Else
        If control.Tag Like MyTag Then
            visible = True
        Else
            visible = False
        End If
    End If
End Sub

Sub HideRibbonControls()
'Hide every Tab, Group or Control with a Tag(we use Tag:="")
    Call RefreshRibbon(Tag:="")
End Sub

Sub ShowRibbonControls()
'Show every Tab, Group or Control(we use the wildcard "*")
    Call RefreshRibbon(Tag:="*")
End Sub

'Callback for btAbout onAction
Sub uiAbout(control As IRibbonControl)
End Sub

'Callback for btSwitchControls onAction
Sub uiSwitchControls(control As IRibbonControl)

    SwitchUI

End Sub

'Callback for btSave onAction
Sub uiSave(control As IRibbonControl)
End Sub

'Callback for btLoad onAction
Sub uiLoad(control As IRibbonControl)
End Sub

'Callback for btLineCartesian onAction
Sub uiLineCartesian(control As IRibbonControl)
End Sub

'Callback for btLinePolar onAction
Sub uiLinePolar(control As IRibbonControl)
End Sub

'Callback for btLineEquation onAction
Sub uiLineEquation(control As IRibbonControl)
End Sub

'Callback for btLineEquationPolar onAction
Sub uiLineEquationPolar(control As IRibbonControl)
End Sub

'Callback for btCircleArc onAction
Sub uiCircleArc(control As IRibbonControl)
End Sub

'Callback for btRectangle onAction
Sub uiRectangle(control As IRibbonControl)
End Sub

'Callback for btPolygon onAction
Sub uiPolygon(control As IRibbonControl)
End Sub

'Callback for btCartesianRepeat onAction
Sub uiCartesianRepeat(control As IRibbonControl)
End Sub

'Callback for btPolarRepeat onAction
Sub uiPolarRepeat(control As IRibbonControl)
End Sub

'Callback for btReflectXY onAction
Sub uiReflectXY(control As IRibbonControl)
End Sub

'Callback for btReflectPolar onAction
Sub uiReflectPolar(control As IRibbonControl)
End Sub

'Callback for btReflectZ onAction
Sub uiReflectZ(control As IRibbonControl)
End Sub

'Callback for btConcentricRepeat onAction
Sub uiConcentricRepeat(control As IRibbonControl)
End Sub

'Callback for btReproduceRecalculate onAction
Sub uiReproduceRecalculate(control As IRibbonControl)
End Sub

'Callback for btRetraction onAction
Sub uiRetraction(control As IRibbonControl)
End Sub

'Callback for btCustomGCODE onAction
Sub uiCustomGCODE(control As IRibbonControl)
End Sub

'Callback for btRepeatRule onAction
Sub uiRepeatRule(control As IRibbonControl)
End Sub

'Callback for btPostprocess onAction
Sub uiPostprocess(control As IRibbonControl)
End Sub

'Callback for btSkipStopUse onAction
Sub uiSkipStopUse(control As IRibbonControl)

    SkipStopUse

End Sub

'Callback for btParameters onAction
Sub uiAssignParameters(control As IRibbonControl)

    SetParameterCellNames

End Sub

'Callback for btGenerateGCODE onAction
Sub uiGenerateGCODE(control As IRibbonControl)



End Sub

Private Sub HideControls()

    Dim btn As Shape
    Dim btns As Shapes
    Dim mainSheet As Worksheet

    Set mainSheet = ThisWorkbook.Sheets("Main Sheet")
    
    With mainSheet
        .Shapes("Flowchart: Alternate Process 15").visible = msoFalse
        .Shapes("Flowchart: Alternate Process 16").visible = msoFalse
        .Shapes("Flowchart: Alternate Process 17").visible = msoFalse
        .Shapes("Flowchart: Alternate Process 18").visible = msoFalse
        .Shapes("Flowchart: Alternate Process 19").visible = msoFalse
        .Shapes("Flowchart: Alternate Process 27").visible = msoFalse
        .Shapes("ParameterButton").visible = msoFalse
        .Shapes("Oval 9").visible = msoFalse
        .Shapes("Oval 21").visible = msoFalse
        .Shapes("Oval 22").visible = msoFalse
        .Shapes("Oval 23").visible = msoFalse
        .Shapes("Oval 24").visible = msoFalse
        .Shapes("Oval 25").visible = msoFalse
        .Shapes("Group 4").visible = msoFalse

        .Rows(1).EntireRow.Hidden = True
    End With

End Sub


Private Sub ShowControls()
    Dim btn As Shape
    Dim btns As Shapes
    Dim mainSheet As Worksheet

    Set mainSheet = ThisWorkbook.Sheets("Main Sheet")
    
    With mainSheet
        .Rows(1).EntireRow.Hidden = False

        .Shapes("Flowchart: Alternate Process 15").visible = msoTrue
        .Shapes("Flowchart: Alternate Process 16").visible = msoTrue
        .Shapes("Flowchart: Alternate Process 17").visible = msoTrue
        .Shapes("Flowchart: Alternate Process 18").visible = msoTrue
        .Shapes("Flowchart: Alternate Process 19").visible = msoTrue
        .Shapes("Flowchart: Alternate Process 27").visible = msoTrue
        .Shapes("ParameterButton").visible = msoTrue
        .Shapes("Oval 9").visible = msoTrue
        .Shapes("Oval 21").visible = msoTrue
        .Shapes("Oval 22").visible = msoTrue
        .Shapes("Oval 23").visible = msoTrue
        .Shapes("Oval 24").visible = msoTrue
        .Shapes("Oval 25").visible = msoTrue
        .Shapes("Group 4").visible = msoTrue
    End With

End Sub
Sub SwitchUI()

    Dim UIstored As String

    UIstored = GetSetting("FullControl", "User", "UIMode", "")

    Select Case UIstored
        Case "" 'Ribbon is default on first use
            SaveSetting "FullControl", "User", "UIMode", "Ribbon"
        Case "Original"
            SaveSetting "FullControl", "User", "UIMode", "Ribbon"
        Case "Ribbon"
            SaveSetting "FullControl", "User", "UIMode", "Original"
    End Select

    SetUI

End Sub

Sub SetUI()

    Dim UIstored As String
    Dim UIMode As String
    
    UIstored = GetSetting("FullControl", "User", "UIMode", "")
    
    'Check which UI we're already using
    On Error Resume Next
    If ThisWorkbook.Sheets("Main Sheet").Shapes("Group 4").visible = msoTrue Then
        UIMode = "Original"
    Else
        UIMode = "Ribbon"
    End If
    On Error GoTo 0
    
    'If it's different to the stored value then change it
    If UIMode <> UIstored Then
    
        Select Case UIstored
            Case "" 'Ribbon is default on first use
                ShowRibbonControls
                HideControls
                SaveSetting "FullControl", "User", "UIMode", "Ribbon"
                Rib.ActivateTab "tbFullControl"
            Case "Original"
                HideRibbonControls
                ShowControls
                SaveSetting "FullControl", "User", "UIMode", "Original"
            Case "Ribbon"
                ShowRibbonControls
                HideControls
                SaveSetting "FullControl", "User", "UIMode", "Ribbon"
                Rib.ActivateTab "tbFullControl"
        End Select

        UIMode = UIstored
        
    End If
    
End Sub
