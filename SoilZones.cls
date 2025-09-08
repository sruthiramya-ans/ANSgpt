VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SoilZones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    If Target.Address = Range("soilZonesCt").Address Then
    
        Application.EnableEvents = False
        
        Dim lstRow As Integer: lstRow = 20
        Dim dropListG As String: dropListG = "G1"
        
        If Cells(Rows.Count, "A").End(xlUp).row > 20 Then lstRow = Cells(Rows.Count, "A").End(xlUp).row
        Rows("20:" & lstRow).Delete
        
        If Target.Value > 1 Then
            For i = 20 To 14 * (Target.Value - 2) + 20 Step 14
                Range("firstSoilZn").Copy Range("A" & i)
                Cells(i, "B") = "G" & (i - 20) / 14 + 2
                dropListG = dropListG & "," & "G" & (i - 20) / 14 + 2
            Next i
        End If
        
        With Sheets("Dashboard").Range("Soil.Zone").Validation
            .Delete ' remove existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:=dropListG
        End With
        
        Sheets("Dashboard").Range("Soil.Zone") = "G1"
        Range("soilZonesCt").Select
        
        Range("soilZonesCt").Select
        
        Application.EnableEvents = True
        
    End If
    
    If Target.Address = Range("scourZonesCt").Address Then
        
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        Dim lstCol As Integer: lstCol = 3
        Dim dropListS As String: dropListS = "S1"
        
        If Cells(2, Columns.Count).End(xlToLeft).Column > 3 Then lstCol = Cells(2, Columns.Count).End(xlToLeft).Column
        
        Range(Cells(2, 3), Cells(3, lstCol)).ClearContents
        Range("C1").Copy
        Range(Cells(2, 3), Cells(3, lstCol)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        
        If Target.Value > 1 Then
            For i = 3 To Target.Value + 1
                Range("firstScourZn").Copy Cells(2, i)
                Range("firstScourD").Copy Cells(3, i)
                Cells(3, i) = Cells(3, i - 1) + 0.25
                Cells(2, i) = "S" & i - 1
                dropListS = dropListS & "," & "S" & i - 1
            Next i
        End If
        
        With Sheets("Dashboard").Range("Scour.Zone").Validation
            .Delete ' remove existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:=dropListS
        End With
        
        Sheets("Dashboard").Range("Scour.Zone") = "S1"
        Range("scourZonesCt").Select
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
    End If
End Sub

