Attribute VB_Name = "PCS_Fixity"
Public Sub CreateFixityLPILE()

'''''''''''''''''''''''''''''''''''
''Determine worst soil parameters''
'''''''''''''''''''''''''''''''''''
'Goal is to limit the number of LPILE files to check fixity

    Dim r As Integer, b As Integer, g As Integer
    r = 104
    g = 126
    b = 103
        
    'Get worst scour
    UpdateProgressBar 5, 100, "Determining worst case Scour Zone", r, g, b
    With SoilZones
        scourZone = ""
        scourDepth = -1
        i = 0
        For i = .Range("firstScourD").Column To .Range("firstScourD").Column + .Range("scourZonesCt") - 1
            If .Cells(.Range("firstScourD").row, i) > scourDepth Then
                scourZone = .Cells(.Range("firstScourZn").row, i)
                scourDepth = .Cells(.Range("firstScourD").row, i)
            End If
            '''MsgBox scourZone & " - " & scourDepth
        Next i
    End With

    'Find worst lateral soil zone
    UpdateProgressBar 10, 100, "Determining worst case Soil Zone", r, g, b
    With SoilZones
        soilZone = ""
        worstSoilZone = ""
        lateralFactor = 10
        j = 1
        zoneRow = 1
        For j = 1 To .Range("soilZonesCt")

            'Find start of next zone table
            k = zoneRow
            Do Until .Cells(k, 1) = "Zone"
                k = k + 1
            Loop
            soilZone = .Cells(k, 2)
            zoneRow = k + 1

            'Get average lateral values across first 15 ft
            avgWt = 0
            avgCu = 0
            avgPhi = 0
            runWt = 0
            runCu = 0
            runPhi = 0
            For m = 1 To 10
                topD = .Cells(zoneRow + m, 2)
                botD = .Cells(zoneRow + m, 3)
                    If botD > 15 Or .Cells(zoneRow + m + 1, 3) = "" Then botD = 15 'check for over 15 ft or last stratum less than 15 ft
                unitWt = .Cells(zoneRow + m, 8)
                cohesion = .Cells(zoneRow + m, 9)
                fricAng = .Cells(zoneRow + m, 10)

                runWt = runWt + unitWt * (botD - topD)
                runCu = runCu + cohesion * (botD - topD)
                runPhi = runPhi + fricAng * (botD - topD)

                If botD = 15 Then Exit For
            Next m
            avgWt = runWt / 15
            avgCu = runCu / 15
            avgPhi = runPhi / 15
            If avgCu / 1000 + avgPhi / 30 < lateralFactor Then
                lateralFactor = avgCu / 1000 + avgPhi / 30
                worstSoilZone = soilZone
            End If
            '''MsgBox worstSoilZone & " - " & lateralFactor
        Next j
    End With

    'Find TOPL with max shear
    Dim cel As Range
    maxShear = WorksheetFunction.Max(TOPLs.Range("TOPL.data").Columns(3), Abs(WorksheetFunction.Min(TOPLs.Range("TOPL.data").Columns(3))))
    For Each cel In TOPLs.Range("TOPL.data").Columns(3).Cells
        If Abs(cel.Value) = maxShear Then maxTOPL = TOPLs.Cells(cel.row, 1)
    Next cel

    '''MsgBox "Worst scour is Zone " & scourZone & " at " & scourDepth & "ft." & vbCrLf & vbCrLf & "Worst soil is Zone " & worstSoilZone & " with an Lateral Factor (normalized equivalent cohesion and friction angle) of " & Format(lateralFactor, "#0.000") & ".", vbInformation, "Worst Scour and Soil Zones"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Create LPILE files and store inputs for fixity checks''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Create files for W6 thru W8 with lbft less than 36
    Dim recorder As Range
    Set recorder = FixityResults.Cells(Range("Fixity.Results").row, Range("Fixity.Results").Column)
    FixityResults.Range("Fixity.Results").ClearContents
    
    n = 0
    shapeCt = WorksheetFunction.CountIf(ShapesDB.Columns("C:C"), "<36")
    
    Dashboard.Range("Pile.Type") = maxTOPL
    Application.Calculate
    
    For Each cel In ShapesDB.Range("AISC.wShapes")
    UpdateProgressBar 10 + 210 * (n / shapeCt), 100, "Creating Lpile files for AISC Shapes W6 - W8 less than 36 plf", r, g, b
    
        'Check AISC shapes
        weightW = Right(cel.Text, Len(cel.Text) - InStr(cel.Text, "X"))
        sizeW = Mid(cel.Text, 2, InStr(1, cel.Text, "X") - 2)
        If sizeW >= 6 And sizeW <= 8 And weightW < 36 Then
            lpileName = cel.Text & " - Fixity Check - Soil Zone " & worstSoilZone & " - Scour Zone " & scourZone
            
            'Set Dashboard to use ANSgpt module for LPile creation
            With Dashboard
            
                .Range("Pile.Embed") = Settings.Range("Settings.FixityDepth")
                .Range("Soil.Zone") = worstSoilZone
                .Range("Scour.Zone") = scourZone
                .Range("Pile.Shape") = cel.Text
                .Range("Lpile.Name") = lpileName
                .Range("Soil.Ignore") = scourDepth * 12
                
                Application.Calculate
                
                Call ANSgptCreator(True, True, False, True, , 0)
                
            End With
                                        
            'Add data to Fixity Results tab
            recorder = lpileName
            recorder.Offset(0, 1) = Settings.Range("Settings.FixityDepth")
            recorder.Offset(0, 2) = Dashboard.Range("Pile.Reveal")
            recorder.Offset(0, 3) = cel.Text
            recorder.Offset(0, 4) = worstSoilZone
            recorder.Offset(0, 5) = scourZone
            
            Set recorder = recorder.Offset(1, 0)
            
            n = n + 1
            
        End If
        
    Next cel
    
    UpdateProgressBar 100, 100, "Creating Lpile files for AISC Shapes W6 - W8 less than 36 plf", r, g, b
    'Message box - notification that routine is complete
    strMsg = "LPile files have been created to check pile Point of Fixity." & vbCrLf & _
    "1. Review LPile files and complete a Batch Run analysis." & vbCrLf & _
    "2. After running, click 'Import LPILE Fixity Results'"
    MsgBox strMsg, vbInformation, "Finished"

End Sub

Public Sub ImportFixityLPILE()

'''''''''''''''''''''''''''''''''''''''''''
''Import and log results from LPILE files''
'''''''''''''''''''''''''''''''''''''''''''
    Dim record As Range, i As Long, recNum As Long, r As Integer, b As Integer, g As Integer
    

    recNum = FixityResults.Range("Fixity.Results").Columns(1).Cells.SpecialCells(xlCellTypeConstants).Count + 1
    i = 1
    r = 104
    g = 126
    b = 103
    
    For Each record In FixityResults.Range("Fixity.Results").Columns(1).Cells
    
        If record = "" Then Exit For
        
        UpdateProgressBar i, recNum, "Importing Lpile Output Results for Point of Fixity", r, g, b
'        r = 255 - 255 * i / recNum
'        g = 255 * i / recNum
        
        'Set Dashboard to use ANSgpt module for LPile creation
        With Dashboard
        
            deflectPof = 0
            slopePof = 0
        
            .Range("Lpile.Name") = record.Text
            .Range("Pile.Embed") = record.Offset(0, 1).Value2
            .Range("Pile.Reveal") = record.Offset(0, 2).Value2
            .Range("Pile.Shape") = record.Offset(0, 3).Text
            .Range("Soil.Zone") = record.Offset(0, 4).Text
            .Range("Scour.Zone") = record.Offset(0, 5).Text
            .Range("Soil.Ignore") = .Range("Soil.Scour").Value2
            
            Application.Calculate
            
            fileName = Settings.Range("LPILE.Folder") & "\" & .Range("Project.Name") & "\Fixity\" & .Range("Lpile.Name") & ".lp12o"
        
            Call ANSgptOutput(fileName)
            
            record.Offset(0, 6) = Dashboard.Range("Steel.AGresult")
            record.Offset(0, 7) = Dashboard.Range("Steel.AMresult")
            
            'Find fixity by delfection and slope
            Dim depth As Range
            For Each depth In .Range("lpile.output2").Columns(1).Cells
                If depth = "" Then Exit For
                
                'Check head deflection
                If depth = 0 Then record.Offset(0, 9) = depth.Offset(0, 1).Value2
                
                'Check grade deflection
                If depth = .Range("Pile.Reveal") Then
                    record.Offset(0, 8) = depth.Offset(0, 1).Value2
                ElseIf depth < .Range("Pile.Reveal") And depth.Offset(1, 0) > .Range("Pile.Reveal") Then
                    record.Offset(0, 8) = depth.Offset(0, 1).Value2 + (depth.Offset(1, 1).Value2 - depth.Offset(0, 1).Value2) / (depth.Offset(1, 0).Value2 - depth.Value2) * (.Range("Pile.Reveal") - depth.Value2)
                End If
            
                'Check deflection fixity
                If depth.Offset(0, 1) * depth.Offset(1, 1) < 0 And deflectPof = 0 Then 'Negative product = sign change
                    deflectPof = depth.Value2 + (depth.Offset(1, 0).Value2 - depth.Value2) / (depth.Offset(1, 1).Value2 - depth.Offset(0, 1).Value2) * (0 - depth.Offset(0, 1).Value2)
                End If
                            
                'Check slope fixity
                If depth.Offset(0, 4) * depth.Offset(1, 4) < 0 And slopePof = 0 Then  'Negative product = sign change
                    slopePof = depth.Value2 + (depth.Offset(1, 0).Value2 - depth.Value2) / (depth.Offset(1, 4).Value2 - depth.Offset(0, 4).Value2) * (0 - depth.Offset(0, 4).Value2)
                End If
                
                
                If deflectPof > 0 And slopePof > 0 Then Exit For
                
            Next depth
                            
            If depth <> "" Then record.Offset(0, 10) = WorksheetFunction.Max(deflectPof, slopePof) - .Range("Pile.Reveal")
            
        End With
        
        i = i + 1
    
    Next record
    
    UpdateProgressBar recNum, recNum, "Importing Lpile Output Results for Point of Fixity", r, g, b
    MsgBox "Lpile output results imported for Point of Fixity analyses.", vbInformation, "Complete"
    FixityResults.Activate
    
End Sub

Sub Fixit()

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

