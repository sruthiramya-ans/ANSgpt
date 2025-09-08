Attribute VB_Name = "PCS_Batcher"
Public Sub generateBatchList()

    Dim shp As Range, geoZone As Range, scourZone As Range, glv As Range, embed As Double, batchDat As Range, fileName As String, reveal As Double, revArray() As Variant, rev As Variant, pileType As String, typ As Range
    Dim brcFiles() As String
    Dim formAGM As String, formAGS As String, formAGMweak As String, formAGSweak As String
    Dim r As Integer, b As Integer, g As Integer, axis_id As Integer, n As Long, fileCount As Long, fldr As String, lps As Double, o As Integer, m As Long
    r = 104
    g = 126
    b = 103
    lps = 6
    axis_id = 0
    
    formAGM = Dashboard.Range("Load.AGM").Formula '=INDEX(Lpile.Moment,MATCH(Pile.Reveal,Lpile.Depth,1)+1)
    formAGS = Dashboard.Range("Load.AGS").Formula '=INDEX(Lpile.Shear,MATCH(Pile.Reveal,Lpile.Depth,1))
    formAGMweak = Dashboard.Range("Load.AGM.Weak").Formula '=INDEX(LPile.Moment.Weak,MATCH(Pile.Reveal,Lpile.Depth,1)+1)
    formAGSweak = Dashboard.Range("Load.AGS.Weak").Formula '=INDEX(Lpile.Shear.Weak,MATCH(Pile.Reveal,Lpile.Depth,1))
    
    BatchResults.Range("Batch.Data").ClearContents
    Set batchDat = BatchResults.Cells(BatchResults.Range("Batch.data").row, BatchResults.Range("Batch.data").Column)

    'Get reveal heights from TOPLs
    'revArray = WorksheetFunction.Unique(TOPLs.Range("TOPL.data").Columns(2).Cells)

    n = 0
    pilect = 0
    ct = 0
    fileCount = WorksheetFunction.CountA(Settings.Range("Settings.TypesList")) * WorksheetFunction.CountA(Settings.Range("Settings.ShapesList")) * _
                WorksheetFunction.CountA(Settings.Range("Settings.GalvList")) * WorksheetFunction.CountA(Settings.Range("Settings.GeoList")) * _
                WorksheetFunction.CountA(Settings.Range("Settings.ScourList")) * ((Settings.Range("Settings.maxEmbed") - Settings.Range("Settings.minEmbed")) / _
                Settings.Range("Settings.intEmbed") + 1) * 2 ' * (UBound(revArray) - 1)
                
    If Settings.Range("Settings.axis") = "Strong" Then
        axis_id = 0
    Else
        axis_id = 1
    End If

    ReDim Preserve brcFiles(1 To fileCount)
    secRemain = fileCount / lps
    
    For Each typ In Settings.Range("Settings.TypesList")
        If typ = "" Then Exit For
        Dashboard.Range("Pile.Type") = typ
        Application.Calculate
        Dashboard.Range("Pile.Reveal") = WorksheetFunction.VLookup(typ, TOPLs.Range("TOPL.data"), 2, False)
        Dashboard.Range("Load.AGM") = Dashboard.Range("TOPL.Moment") + Dashboard.Range("TOPL.Shear") * (Dashboard.Range("Pile.Reveal") * 12 + Dashboard.Range("Soil.Scour"))
        Dashboard.Range("Load.AGS") = Dashboard.Range("TOPL.Shear")
        Dashboard.Range("Load.AGM.Weak") = Dashboard.Range("TOPL.Moment.Weak") + Dashboard.Range("TOPL.Shear.Weak") * (Dashboard.Range("Pile.Reveal") * 12 + Dashboard.Range("Soil.Scour"))
        Dashboard.Range("Load.AGS.Weak") = Dashboard.Range("TOPL.Shear.Weak")
        
        For Each shp In Settings.Range("Settings.ShapesList")
            If shp = "" Then Exit For
            Dashboard.Range("Pile.Shape") = shp
                
            For Each glv In Settings.Range("Settings.GalvList")
                If glv = "" Then Exit For
                Dashboard.Range("Pile.Galv") = glv
                        
                For Each geoZone In Settings.Range("Settings.GeoList")
                    If geoZone = "" Then Exit For
                    Dashboard.Range("Soil.Zone") = geoZone
                
                    For Each scourZone In Settings.Range("Settings.ScourList")
                        If scourZone = "" Then Exit For
                        Dashboard.Range("Scour.Zone") = scourZone

                        'Limit embed depth by soil axial checks
                        Application.Calculate
                        Dashboard.Range("Soil.AxialResult").GoalSeek Goal:=1, ChangingCell:=Dashboard.Range("Pile.Embed")
                        embed = Application.WorksheetFunction.Max(Settings.Range("Settings.minEmbed"), Application.WorksheetFunction.Ceiling_Math(Dashboard.Range("Pile.Embed"), Settings.Range("Settings.intEmbed")))
                        Dashboard.Range("Pile.Embed") = embed
                        
                        Do Until embed > Settings.Range("Settings.maxEmbed")
                            'Limit results by ignoring failing soil axil or steel checks
                            Application.Calculate
                            'If Dashboard.Range("Soil.AxialResult") <= 1 And Dashboard.Range("Steel.AGresult") <= 1 Then

                                batchDat = typ
                                
                                batchDat.Offset(0, 3) = shp
                                batchDat.Offset(0, 4) = glv
                                batchDat.Offset(0, 5) = Dashboard.Range("Pile.Reveal")
                                batchDat.Offset(0, 6) = embed
                                batchDat.Offset(0, 7) = geoZone
                                batchDat.Offset(0, 8) = scourZone
                                
                                batchDat.Offset(0, 13) = Dashboard.Range("Soil.AxialResult")
                                batchDat.Offset(0, 14) = Dashboard.Range("Steel.AGresult")
                                'batchDat.Offset(0, 15) = Dashboard.Range("Steel.AMresult")

                                batchDat.Offset(0, 16) = (embed + Dashboard.Range("Pile.Reveal")) * Right(shp, Len(shp) - InStr(1, shp, "X"))
                                batchDat.Offset(0, 17) = Dashboard.Range("TOPL.selected.sMu")
                                batchDat.Offset(0, 18) = Dashboard.Range("TOPL.selected.sVu")
                                batchDat.Offset(0, 19) = Dashboard.Range("TOPL.M_external_weak")
                                batchDat.Offset(0, 20) = Dashboard.Range("TOPL.Shear.Weak")
                                batchDat.Offset(0, 21) = Dashboard.Range("TOPL.selected.sPu")
                                batchDat.Offset(0, 22) = Dashboard.Range("TOPL.selected.sTu")
                                
                                'Loop thru steel grades and record results
                                
                                fldr = typ
                                
                                If Dashboard.Range("Soil.AxialResult") <= 1 And Dashboard.Range("Steel.AGresult") <= 1 Then
                                    
                                    fileName = typ & "-" & shp & "-Embed " & embed & "ft-" & glv & " mil-Soil " & geoZone & "-Scour " & scourZone & "Strong"
                                    batchDat.Offset(0, 23) = fileName
                                    Dashboard.Range("Lpile.Name") = fileName
                                    Call ANSgptCreator(True, True, False, False, fldr, 0) ' strong
                                    pilect = pilect + 1
                                    brcFiles(pilect) = Range("LPILE.Folder") & "\" & Range("Project.Name") & "\" & fldr & "\" & fileName & ".lp12d"
'                                    Set batchDat = batchDat.Offset(1, 0)
'                                    batchDat = typ
'
'                                    batchDat.Offset(0, 3) = shp
'                                    batchDat.Offset(0, 4) = glv
'                                    batchDat.Offset(0, 5) = Dashboard.Range("Pile.Reveal")
'                                    batchDat.Offset(0, 6) = embed
'                                    batchDat.Offset(0, 7) = geoZone
'                                    batchDat.Offset(0, 8) = scourZone
'
'                                    batchDat.Offset(0, 13) = Dashboard.Range("Soil.AxialResult")
'                                    batchDat.Offset(0, 14) = Dashboard.Range("Steel.AGresult")
'                                    'batchDat.Offset(0, 15) = Dashboard.Range("Steel.AMresult")
'
'                                    batchDat.Offset(0, 16) = (embed + Dashboard.Range("Pile.Reveal")) * Right(shp, Len(shp) - InStr(1, shp, "X"))
'                                    batchDat.Offset(0, 17) = Dashboard.Range("TOPL.selected.sMu")
'                                    batchDat.Offset(0, 18) = Dashboard.Range("TOPL.selected.sVu")
'                                    batchDat.Offset(0, 19) = "-" ' Dashboard.Range("TOPL.M_external_weak")
'                                    batchDat.Offset(0, 20) = "-" ' Dashboard.Range("TOPL.Shear.Weak")
'                                    batchDat.Offset(0, 21) = Dashboard.Range("TOPL.selected.sPu")
'                                    batchDat.Offset(0, 22) = Dashboard.Range("TOPL.selected.sTu")
                                    
                                    fileName = typ & "-" & shp & "-Embed " & embed & "ft-" & glv & " mil-Soil " & geoZone & "-Scour " & scourZone & "Weak"
                                    batchDat.Offset(0, 24) = fileName
'                                    batchDat.Offset(0, 24) = "Weak"
                                    Dashboard.Range("Lpile.Name") = fileName
                                    Call ANSgptCreator(True, True, False, False, fldr, 1) ' weak
                                    pilect = pilect + 1
                                    brcFiles(pilect) = Range("LPILE.Folder") & "\" & Range("Project.Name") & "\" & fldr & "\" & fileName & ".lp12d"
                                End If

                                Set batchDat = batchDat.Offset(1, 0)
                                
                            'End If
                            
                            embed = embed + Settings.Range("Settings.intEmbed")
                            Dashboard.Range("Pile.Embed") = embed
                            
                            n = n + 1
                            secRemain = secRemain - 1 / lps
                            UpdateProgressBar n, fileCount, "Generating up to " & fileCount & " Lpile files for Batch Analysis (Time Remaining: " & WorksheetFunction.Floor((secRemain / 60), 1) & ":" & Format(WorksheetFunction.RoundDown(secRemain Mod 60, 0), "00") & ")", r, g, b

                        Loop
                            
                    Next scourZone
                        
                Next geoZone
                                    
            Next glv
                    
        Next shp
        
    Next typ
    
    ReDim Preserve brcFiles(1 To pilect)
    BatchBRCfiles (brcFiles)
    
    Dashboard.Range("Load.AGM") = formAGM
    Dashboard.Range("Load.AGS") = formAGS
    Dashboard.Range("Load.AGM.Weak") = formAGMweak
    Dashboard.Range("Load.AGS.Weak") = formAGSweak
    
    UpdateProgressBar fileCount, fileCount, "Generating up to " & fileCount & " Lpile files for Batch Analysis", r, g, b
    
    MsgBox pilect & " Lpile created out of a possible " & fileCount & " scenarios. Cases were excluded based on failing soil axial or steel results.", vbInformation, "Batch files created"
        
End Sub

Sub importBatchList()

    Dim lpileName As Range, lpileNameWk As Range, lpileArray As Variant, i As Long, fileNameSt As String, fileNameWk As String
    Dim formAGM As String, formAGS As String, formAMM As String
    
    ' Retain formulas for cells to import back in later '
    
    Dashboard.Range("lpile.output2").ClearContents
    Dashboard.Range("lpile.output2.weak").ClearContents
    Application.Calculate
    
    formAGM = Dashboard.Range("Load.AGM").Formula '=INDEX(Lpile.Moment,MATCH(Pile.Reveal,Lpile.Depth,1)+1)
    formAGS = Dashboard.Range("Load.AGS").Formula '=INDEX(Lpile.Shear,MATCH(Pile.Reveal,Lpile.Depth,1))
    formAMM = Dashboard.Range("Load.AMM").Formula '=MAX(Lpile.Moment)
    formAGMweak = Dashboard.Range("Load.AGM.Weak").Formula '=INDEX(LPile.Moment.Weak,MATCH(Pile.Reveal,Lpile.Depth,1)+1)
    formAGSweak = Dashboard.Range("Load.AGS.Weak").Formula '=INDEX(Lpile.Shear.Weak,MATCH(Pile.Reveal,Lpile.Depth,1))
    formAMMweak = Dashboard.Range("Load.AMM.Weak").Formula '=MAX(LPile.Moment.Weak)
    Dim r As Integer, b As Integer, g As Integer, n As Long, fileCount As Long, fldr As String
    r = 104
    g = 126
    b = 103
    
    n = 0
    fileCount = WorksheetFunction.CountA(BatchResults.Range("Batch.data").Columns(1).Cells)

    'For Each lpileName In BatchResults.Range("Batch.data").Columns(24).Cells
    For i = 7 To 7 + fileCount
        Set lpileName = BatchResults.Cells(i, 24)
        Set lpileNameWk = BatchResults.Cells(i, 25)
        
        If lpileName.Value = "" Then GoTo NextIteration
        
        folderName = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & BatchResults.Cells(lpileName.row, 1) & "\"
        fileNameSt = folderName & lpileName & ".lp12o"
        fileNameWk = folderName & lpileNameWk & ".lp12o"

        lpileArray = LpileReader(fileNameSt)
        lpileArrayWk = LpileReader(fileNameWk)
        
        n = n + 1
        UpdateProgressBar n, fileCount, "Importing up to " & fileCount & " Lpile results files to Batch Results", r, g, b
        
        On Error Resume Next
        gradeDeflUse = Settings.Range("Settings.GradeDefl") / lpileArray(1, 3)
        headDeflUse = Settings.Range("Settings.HeadDefl") / lpileArray(1, 4)

        gradeDeflUseWk = Settings.Range("Settings.GradeDefl") / lpileArrayWk(1, 3)
        headDeflUseWk = Settings.Range("Settings.HeadDefl") / lpileArrayWk(1, 4)
        
        If Err Then
            errct = errct + 1
            BatchResults.Cells(lpileName.row, 2) = "Not Found"
            GoTo skipfile
        End If
        ' = lpileArray(1, 2) 'pile
        ' = lpileArray(1, 6) 'Vu
        ' = lpileArray(1, 7) 'Axis
        If lpileArray(1, 7) = 0 Then
            BatchResults.Cells(lpileName.row, 10) = lpileArray(1, 3) 'gradeDefl strong
            BatchResults.Cells(lpileName.row, 11) = lpileArray(1, 4) 'headDefl strong
            BatchResults.Cells(lpileName.row, 12) = lpileArrayWk(1, 3) 'gradeDefl weak
            BatchResults.Cells(lpileName.row, 13) = lpileArrayWk(1, 4) 'headDefl weak
            
            Dashboard.Range("Pile.Shape") = BatchResults.Cells(i, 4)
            Dashboard.Range("Load.AGM") = lpileArray(1, 8) 'AGM
            Dashboard.Range("Load.AGS") = lpileArray(1, 9) 'AGS
            Dashboard.Range("Load.AMM") = lpileArray(1, 5) 'AMM
            Dashboard.Range("Load.AGM.Weak") = lpileArrayWk(1, 8) 'AGMWk
            Dashboard.Range("Load.AGS.Weak") = lpileArrayWk(1, 9) 'AGSWk
            Dashboard.Range("Load.AMM.Weak") = lpileArrayWk(1, 5) 'AMMWk
            
            BatchResults.Cells(lpileName.row, 26) = lpileArray(1, 8) 'AGM
            BatchResults.Cells(lpileName.row, 27) = lpileArray(1, 9) 'AGS
            BatchResults.Cells(lpileName.row, 28) = lpileArray(1, 5) 'AMM
            BatchResults.Cells(lpileName.row, 29) = lpileArrayWk(1, 8) 'AGMWk
            BatchResults.Cells(lpileName.row, 30) = lpileArrayWk(1, 9) 'AGSWk
            BatchResults.Cells(lpileName.row, 31) = lpileArrayWk(1, 5) 'AMMWk
            
'            Dashboard.Range("Load.AGM.Weak") = Dashboard.Range("TOPL.Moment.Weak") + Dashboard.Range("TOPL.Shear.Weak") * (Dashboard.Range("Pile.Reveal") * 12 + Dashboard.Range("Soil.Scour"))
'            Dashboard.Range("Load.AGS.Weak") = Dashboard.Range("TOPL.Shear.Weak")
            Application.Calculate
            BatchResults.Cells(lpileName.row, 16) = Dashboard.Range("STEEL.AMresult")
            BatchResults.Cells(lpileName.row, 15) = Dashboard.Range("Steel.AGresult")
        Else
            BatchResults.Cells(lpileName.row, 12) = lpileArray(1, 3) 'gradeDefl
            BatchResults.Cells(lpileName.row, 13) = lpileArray(1, 4) 'headDefl
        End If
        
        'Determine controlling case
        If gradeDeflUse < headDeflUse And gradeDeflUse <> "" Then
            If headDeflUse < BatchResults.Cells(lpileName.row, 14) And headDeflUse <> "" Then
                If BatchResults.Cells(lpileName.row, 14) < BatchResults.Cells(lpileName.row, 15) And BatchResults.Cells(lpileName.row, 14) <> "" Then
                    If BatchResults.Cells(lpileName.row, 15) < BatchResults.Cells(lpileName.row, 16) And BatchResults.Cells(lpileName.row, 15) <> "" Then
                        cCase = "Steel AM"
                    Else
                        cCase = "Steel AG"
                    End If
                Else
                    cCase = "Soil Axial"
                End If
            Else
                cCase = "Head Defl."
            End If
        Else
            cCase = "Grade Defl."
        End If
        
        BatchResults.Cells(lpileName.row, 3) = cCase 'lpileArray(1, 1) 'LC
        
        'Check all results for pass/fail
        If BatchResults.Cells(lpileName.row, 3) = "" Then
            BatchResults.Cells(lpileName.row, 2) = "Not Solved"
        ElseIf BatchResults.Cells(lpileName.row, 10) < Settings.Range("Settings.GradeDefl") And _
                BatchResults.Cells(lpileName.row, 11) < Settings.Range("Settings.HeadDefl") And _
                BatchResults.Cells(lpileName.row, 14) < 1 And _
                BatchResults.Cells(lpileName.row, 15) < 1 And _
                BatchResults.Cells(lpileName.row, 16) < 1 Then
            BatchResults.Cells(lpileName.row, 2) = "Pass"
        Else
            BatchResults.Cells(lpileName.row, 2) = "Fail"
        End If

        
skipfile:
        'On Error GoTo 0
NextIteration:
    Next i
    
    
' Reinsert formlaus for cells that we overwritten

    Dashboard.Range("Load.AGM").Formula = formAGM
    Dashboard.Range("Load.AGS").Formula = formAGS
    Dashboard.Range("Load.AMM").Formula = formAMM
    Dashboard.Range("Load.AGM.Weak").Formula = formAGMweak
    Dashboard.Range("Load.AGS.Weak").Formula = formAGSweak
    Dashboard.Range("Load.AMM.Weak").Formula = formAMMweak
    
    UpdateProgressBar 100, 100, "Importing up to " & fileCount & " Lpile results files to Batch Results", r, g, b
    If errct >= 1 Then MsgBox errct & " solution files not found or imported out of a possible files " & _
                        WorksheetFunction.CountA(BatchResults.Range("Batch.data").Columns(24).Cells) & ".", vbExclamation, "Results files not found"
    
End Sub


Sub BatchBRCfiles(ByVal arr As Variant)

    Const BATCH_SIZE As Long = 100
    Dim totalItems As Long
    Dim i As Long
    Dim startIdx As Long
    Dim endIdx As Long
    Dim batchNum As Long
    Dim stamp As String
    Dim outPath As String
    Dim outFileName As String
    Dim fnum As Integer
        
    ' Ensure arr is a properly initialized 1-D array
    If Not IsArray(arr) Then
        Exit Sub
    End If
    
    totalItems = UBound(arr) - LBound(arr) + 1
    If totalItems < 1 Then
        Exit Sub
    End If
    
    stamp = Format(Now, "yyyymmdd_HHMMSS")
    outPath = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name") & "\" & "Batch " & stamp)
    batchNum = 0
    
    ' Loop through the array in chunks of BATCH_SIZE
    For startIdx = LBound(arr) To UBound(arr) Step BATCH_SIZE
        
        batchNum = batchNum + 1
        endIdx = Application.Min(startIdx + BATCH_SIZE - 1, UBound(arr))
        
        outFileName = "B" & batchNum & _
                      "_" & startIdx & "-" & endIdx & ".brc"
        
        fnum = FreeFile
        Open outPath & outFileName For Output As #fnum
        
        ' Write each item on its own line
        For i = startIdx To endIdx
            Print #fnum, arr(i)
        Next i
        
        Close #fnum
        
    Next startIdx

End Sub




