Attribute VB_Name = "ANSgpt"
Option Explicit
Option Base 1
Const PI = 3.14159265358979
Public fileName As String


'Create LPILE Files
Public Sub ANSgptCreator(batch As Boolean, overwrite As Boolean, openLpile As Boolean, fixityCheck As Boolean, Optional newFolder As String, Optional orientation As Integer) ' Orientation (0=strong, 1=weak) we can cgange to cell if easier -  will eventually need inputs for file naming conventions as we batch runs
    'define subfolder for lpile, folder check variable, lpile file name & its checker variable
    Dim FolderSpec As String, folderName As String, fileSpec As String, fileName As String
    'define strings for file content
    Dim data1a As String, data1b As String, data2 As String, data3 As String, data4a As String, data4b As String, data4c As String, data5 As String
    'define pile length variable, number of soil layers, generic iteration integers
    Dim L1 As Double, L2 As Double, laynum As Integer, i As Integer, j As Integer, k As Integer, depth As Double, pmult As Double, ymult As Double
    'define parameter arrays to import from spreadsheet to VBA
    Dim arrDepthTop() As Variant, arrDepthBot() As Variant, arrGamma() As Variant, arrGammaEff() As Variant, arrCohesion() As Variant, arrPHI() As Variant, arrk() As Variant, arrPYcurve() As Variant, arrPYdepth() As Variant, arrPYPmult() As Variant, arrPYYmult() As Variant
    'define file creation variables
    Dim f As Integer, strMsg As String, b As Integer
    Dim AG_IxIy As Double, BG_IxIy As Double

    Dim rngSoil As Range, bAxis As String


    'Hide alerts
'    Application.DisplayAlerts = False

    'Check to see if LPile is already running - simplifies operations if we can point to one instance of Lpile
    Do While True
        If IsExeRunning("LPile2019.exe") = False Then
            Exit Do
        ElseIf IsExeRunning("LPile2022.exe") = False Then 'Add check for Lpile2022.exe
            Exit Do
        Else
            MsgBox "To continue, please close all instances of LPile that are running.", vbCritical, "Close LPile"
            Exit Sub
        End If
    Loop

    'Calculate total pile length
    L1 = Range("Pile.Reveal").Value2
    L2 = Range("Pile.Embed").Value2


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Create arrays of soil properties to print to LPile file''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Determine number of soil strata needed
'    laynum = WorksheetFunction.Count([Layer.Top])

    Set rngSoil = Range("Layer.Top")
    For i = rngSoil.Rows.Count To 1 Step -1
        ' Check that the cell is not empty and not zero:
        If Nz(rngSoil.Cells(i, 1).Value) <> 0 Then
            laynum = i
            Exit For
        End If
    Next i

    'PLACEHOLDER to Add groundwater adjustments (effective weight) in future development
    'PLACEHOLDER to define batter and slope per install tolerances - for leaner designs, worst case is likely secondary moment about weak axis
    'PLACEHOLDER to adjust elastic stiffness as needed for rotation tolerance - should be close enough to handle with steel check post-processing
    'PLACEHOLDER to Add sorting, remove duplicates, remove empties, or automatic layer injection in future development as needed

    'Redim/Populate arrays per number of soil strata
    ReDim arrDepthTop(1 To laynum)
    ReDim arrDepthBot(1 To laynum)
    ReDim arrGamma(1 To laynum)
    ReDim arrGammaEff(1 To laynum)
    ReDim arrCohesion(1 To laynum)
    ReDim arrPHI(1 To laynum)
    ReDim arrPYcurve(1 To laynum)
    ReDim arrk(1 To laynum)
    ReDim arrE60(1 To laynum)
    'array per py factors
    ReDim arrPYdepth(1 To [py_layer_count])
    ReDim arrPYPmult(1 To [py_layer_count])
    ReDim arrPYYmult(1 To [py_layer_count])


    'populate arrays by row/integer
    'PLACEHOLDER to shift strata down based on depth to ignore, scour, etc
    For i = 1 To laynum
        arrDepthTop(i) = [Layer.Top].Cells(i).Value2
        arrDepthBot(i) = [Layer.Bot].Cells(i).Value2
        arrGamma(i) = [Layer.uWt].Cells(i).Value2
        arrGammaEff(i) = [Layer.uWt].Cells(i).Value2
        arrCohesion(i) = [Layer.Cohesion].Cells(i).Value2
        arrPHI(i) = [Layer.FrAngle].Cells(i).Value2
        arrPYcurve(i) = [Layer.Material].Cells(i).Value2
        arrk(i) = [Layer.k].Cells(i).Value2
        arrE60(i) = [Layer.E60].Cells(i).Value2
    Next i
    
    For i = 1 To [py_layer_count]
        arrPYdepth(i) = [py.depth_below_pile_head].Cells(i).Value2
        arrPYPmult(i) = [py.p_mult].Cells(i).Value2
        arrPYYmult(i) = [py.y_mult].Cells(i).Value2
    Next i


    'modify arrays based on soil properties
    'PLACEHOLDER FOR GROUNDWATER adjustments for soil types - stick to sand and clay for initial execution
    'PLACEHOLDER for soft clays

    'adjust depth at bottom of the array as needed - ref L, top layer, embed, etc
    If arrDepthBot(laynum) < L1 + L2 + 1 Then
        arrDepthBot(laynum) = L1 + L2 + 1
    End If


'   Create array for LPile p-y curve numbers from input
    For i = 1 To laynum
        Select Case arrPYcurve(i)
            Case Is = ""
                If (arrPHI(i) = 0 Or arrPHI(i) = Empty) And (arrCohesion(i) = 0 Or arrCohesion(i) = Empty) Then
                    Resume Next
                Else
                    MsgBox "Enter LPile p-y Curve Type", vbCritical, "Input Error"
                    Exit Sub
                End If
            Case Is = "Soft Clay"
                arrPYcurve(i) = 1
            Case Is = "Stiff Clay with Free Water"
                arrPYcurve(i) = 3
            Case Is = "Stiff Clay w/o Free Water"
                arrPYcurve(i) = 4
            Case Is = "Sand"
                arrPYcurve(i) = 6
            Case Is = "Strong Rock"
                arrPYcurve(i) = 11
            Case Is = "Silt"
                arrPYcurve(i) = 15
        End Select
    Next i

    'PLACEHOLDER FOR PY INPUTS - reference PY place in options generation below as well, table inputs would need to be separate text chunk

'''''''''''''''''''''''''''''''''''''''''''''
''Create LPile files and populate with data''
'''''''''''''''''''''''''''''''''''''''''''''

'   Set LPile Folder and File
    If fixityCheck = True Then
        folderName = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & "Fixity\"
    ElseIf newFolder <> "" Then
        folderName = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & newFolder & "\"
    Else
        folderName = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & "Single Run\"
    End If
    CreateDir (folderName)
    If batch = False Then
        If orientation = 0 Then bAxis = "ST" Else bAxis = "WK"
        fileName = folderName & Dashboard.Range("Lpile.Name") & "(" & bAxis & ").lp12d"
        fileSpec = folderName & Dashboard.Range("Lpile.Name") & "(" & bAxis & ").lp12d"
    Else
        fileName = folderName & Dashboard.Range("Lpile.Name") & ".lp12d"
        fileSpec = folderName & Dashboard.Range("Lpile.Name") & ".lp12d"
    End If

'   Check to see if LPile files already exist
    If Dir(folderName & "*.*") <> "" Then
        If overwrite = False Then
            Exit Sub
        Else
            On Error Resume Next
            Kill fileName
            On Error GoTo -1
        End If
    End If

'    FileSpec = ThisWorkbook.Path & "\LPile\" & [Project.Name].Value2 & "- ANSgpt.lp11d"
'    FileName = Dir(FileSpec)

'''''''''''''''''''''''''''''''''''''''''''
'   Create data strings to write to file  '
'''''''''''''''''''''''''''''''''''''''''''
'   Project information
    data1a = _
    "LPILEP12" & vbCrLf & _
    "TITLE" & vbCrLf & _
    "Project Name: " & [Project.Name].Value2 & vbCrLf & _
    "Job Number: " & vbCrLf & _
    "Client: " & vbCrLf & _
    "Engineer: " & Environ("USERNAME") & vbCrLf

    data1b = _
    "Description: " & [Pile.Type].Value2 & " Pile Design" & vbCrLf

'   Default program options
    'PLACEHOLDER to add PY mult options when ready
    data2 = "OPTIONS" & vbCrLf & "Units USCS" & vbCrLf & "UseLRFD NO" & vbCrLf & "UseLayeringCorrection YES" & vbCrLf & _
    "UseinSoilsofSameType YES" & vbCrLf & "ComputeEIOnly NO" & vbCrLf & _
    "Loading STATIC" & vbCrLf & "UsePYModifiers YES" & vbCrLf & "UseTipShear NO" & vbCrLf & "UseDistributedLoading NO" & vbCrLf & _
    "UseSoilMovement NO" & vbCrLf & "ComputeKmatrix NO" & vbCrLf & "ComputePushover NO" & vbCrLf & "ComputePileBuckling NO" & vbCrLf & _
    "NumberPileIncrements" & vbCrLf & "100" & vbCrLf & "IterationsLimit" & vbCrLf & "500" & vbCrLf & "MaxDeflectionLimit" & vbCrLf & _
    " 1.0000000000000E+0002" & vbCrLf & "ConvergenceTolerance" & vbCrLf & " 1.00000000000000E-0005" & vbCrLf & _
    "PrintPYCurves NO" & vbCrLf & "PrintSummaryOnly NO" & vbCrLf & "1 = Printing Increment" & vbCrLf & _
    "PrintNarrowReport NO" & vbCrLf & "ComputeShearCapacity NO" & vbCrLf & "ComputeInteraction NO" & vbCrLf & "END OPTIONS" & vbCrLf

'   Define pile parameters
'PLACEHOLDER to have multiple pile sections for atmospheric vs soil corrosion, collars, etc -- UPDATE added above grade and below grade sections

    If orientation = 1 Then
        AG_IxIy = [CorrMem.Iy.AG].Value2
        BG_IxIy = [CorrMem.Iy.BG].Value2
    Else
        AG_IxIy = [CorrMem.Ix.AG].Value2
        BG_IxIy = [CorrMem.Ix.BG].Value2
    End If

    data3 = _
    "SECTIONS" & vbCrLf & _
    "2 = Total Number of Sections" & vbCrLf & _
    "1 = Section Number" & vbCrLf & _
    "11 = Section type =  elastic section" & vbCrLf & _
    " " & Format(L1, "0.00000000000000E+0000") & "  = Section length (ft)" & vbCrLf & _
    orientation + 4 & " = Elastic Strong/Weak H section" & vbCrLf & _
    " " & Round(Format(29000000, "0.00000000000000E+0000"), 0) & "  = Elastic modulus (psi)" & vbCrLf & _
    " " & Round(Format([CorrMem.Width.AG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section width (in)" & vbCrLf & _
    " " & Round(Format([CorrMem.Depth.AG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section depth (in)" & vbCrLf & _
    " " & Round(Format([CorrMem.Flange_t.AG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section flange thickness (in)" & vbCrLf & _
    " " & Round(Format([CorrMem.Web_t.AG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section web thickness (in)" & vbCrLf & _
    " " & Round(Format([CorrMem.Area.AG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section area (sq in)" & vbCrLf & _
    " " & Round(Format(AG_IxIy, "0.00000000000000E+0000"), 3) & "  = H Section MOI (in^4)" & vbCrLf & _
    "2 = Section Number" & vbCrLf & _
    "11 = Section type =  elastic section" & vbCrLf & _
    " " & Format(L2, "0.00000000000000E+0000") & "  = Section length (ft)" & vbCrLf & _
    orientation + 4 & " = Elastic Strong/Weak H section" & vbCrLf & _
    " " & Round(Format(29000000, "0.00000000000000E+0000"), 0) & "  = Elastic modulus (psi)" & vbCrLf & _
    " " & Round(Format([CorrMem.Width.BG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section width (in)" & vbCrLf & _
    " " & Round(Format([CorrMem.Depth.BG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section depth (in)" & vbCrLf & _
    " " & Round(Format([CorrMem.Flange_t.BG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section flange thickness (in)" & vbCrLf & _
    " " & Round(Format([CorrMem.Web_t.BG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section web thickness (in)" & vbCrLf & _
    " " & Round(Format([CorrMem.Area.BG].Value2, "0.00000000000000E+0000"), 3) & "  = H Section area (sq in)" & vbCrLf & _
    " " & Round(Format(BG_IxIy, "0.00000000000000E+0000"), 3) & "  = H Section MOI (in^4)" & vbCrLf

'   Define soil parameters
'PLACEHOLDER TO add batter/slope install tolerance

    data4a = _
    "SOIL LAYERS" & vbCrLf & _
    Format(laynum, "0") & " = number of soil layers" & vbCrLf

    For i = 1 To laynum
        Select Case arrPYcurve(i)
            Case Is = "1"
                data4b = data4b + _
                "1        " & Format([Pile.Reveal].Value2 + arrDepthTop(i), "0.00000000000000E+0000") & "   " & Format([Pile.Reveal].Value2 + arrDepthBot(i), "0.00000000000000E+0000") & "  = soil type number for soft clay, Xtop (ft), Xbot(ft)" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "   " & Format(arrE60(i), "0.00000000000000E+0000") & " = top gamma (pcf), c (psf), epsilon_50 for soft clay" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "   " & Format(arrE60(i), "0.00000000000000E+0000") & " = bot gamma (pcf), c (psf), epsilon_50 for soft clay" & vbCrLf
            Case Is = "3"
                data4b = data4b + _
                "3        " & Format([Pile.Reveal].Value2 + arrDepthTop(i), "0.00000000000000E+0000") & "   " & Format([Pile.Reveal].Value2 + arrDepthBot(i), "0.00000000000000E+0000") & "  = soil type number for stiff clay w/free water, Xtop (ft), Xbot(ft)" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "   " & Format(arrE60(i), "0.00000000000000E+0000") & "   " & Format(arrk(i), "0.00000000000000E+0000") & " = top gamma (pcf), c (psf), epsilon_50, k (pci) for stiff clay w/ free water" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "   " & Format(arrE60(i), "0.00000000000000E+0000") & "   " & Format(arrk(i), "0.00000000000000E+0000") & " = bot gamma (pcf), c (psf), epsilon_50, k (pci) for stiff clay w/ free water" & vbCrLf
            Case Is = "4"
                data4b = data4b + _
                "4        " & Format([Pile.Reveal].Value2 + arrDepthTop(i), "0.00000000000000E+0000") & "   " & Format([Pile.Reveal].Value2 + arrDepthBot(i), "0.00000000000000E+0000") & "  = soil type number for stiff clay no water, Xtop (ft), Xbot(ft)" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "   " & Format(arrE60(i), "0.00000000000000E+0000") & " = top gamma (pcf), c (psf), epsilon_50 for stiff clay w/o free water" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "   " & Format(arrE60(i), "0.00000000000000E+0000") & " = bot gamma (pcf), c (psf), epsilon_50 for stiff clay w/o free water" & vbCrLf
            Case Is = "6"
                data4b = data4b + _
                "6        " & Format([Pile.Reveal].Value2 + arrDepthTop(i), "0.00000000000000E+0000") & "   " & Format([Pile.Reveal].Value2 + arrDepthBot(i), "0.00000000000000E+0000") & "  = soil type number for Reese sand, Xtop (ft), Xbot(ft)" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrPHI(i), "0.00000000000000E+0000") & "   " & Format(arrk(i), "0.00000000000000E+0000") & " = top gamma (pcf), phi (deg), k (pci) for sand" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrPHI(i), "0.00000000000000E+0000") & "   " & Format(arrk(i), "0.00000000000000E+0000") & " = bot gamma (pcf), phi (deg), k (pci) for sand" & vbCrLf
            Case Is = "11"
                data4b = data4b + _
                "11        " & Format([Pile.Reveal].Value2 + arrDepthTop(i), "0.00000000000000E+0000") & "   " & Format([Pile.Reveal].Value2 + arrDepthBot(i), "0.00000000000000E+0000") & "  = soil type number for vuggy limestone rock, Xtop (ft), Xbot(ft)" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "  = top gamma (pcf), qu (psi) for vuggy limestone" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "  = bot gamma (pcf), qu (psi) for vuggy limestone" & vbCrLf
            Case Is = "15"
                data4b = data4b + _
                "15       " & Format([Pile.Reveal].Value2 + arrDepthTop(i), "0.00000000000000E+0000") & "   " & Format(Range("Pile.Reveal") + arrDepthBot(i), "0.00000000000000E+0000") & "  = soil type number for cemented c-phi silt, Xtop (ft), Xbot(ft)" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "   " & Format(arrPHI(i), "0.00000000000000E+0000") & "   " & Format(arrE60(i), "0.00000000000000E+0000") & "   " & Format(arrk(i), "0.00000000000000E+0000") & " = top gamma (pcf), c (psf) , phi (deg), epsilon_50, k (pci) for cemented silt" & vbCrLf & _
                " " & Format(arrGammaEff(i), "0.00000000000000E+0000") & "   " & Format(arrCohesion(i), "0.00000000000000E+0000") & "   " & Format(arrPHI(i), "0.00000000000000E+0000") & "   " & Format(arrE60(i), "0.00000000000000E+0000") & "   " & Format(arrk(i), "0.00000000000000E+0000") & " = bot gamma (pcf), c (psf) , phi (deg), epsilon_50, k (pci) for cemented silt" & vbCrLf
        End Select
    Next i

    data4c = _
    "PILE BATTER AND SLOPE" & vbCrLf & _
    " 0.00000000000000E+0000 = Ground Slope (deg)" & vbCrLf & _
    " 0.00000000000000E+0000 = Pile Batter  (deg)" & vbCrLf
    
    'P-MULT
    data4c = data4c & _
    "GROUP EFFECT FACTORS" & vbCrLf & _
    [py_layer_count] & vbCrLf

    ' Loop through each row of data
    For i = 1 To [py_layer_count]
        depth = arrPYdepth(i)
        pmult = arrPYPmult(i)
        ymult = arrPYYmult(i)

        data4c = data4c & _
            (i) & "    " & _
            Format(depth, "0.000000000000000E+0000") & "  " & _
            Format(pmult, "0.00000000000000E+0000") & "  " & _
            Format(ymult, "0.00000000000000E+0000") & " = point " & (i) & ", depth (ft), p-multiplier, y-multiplier" & vbCrLf
    Next i


'   Define loading conditions
    'PLACEHOLDER to add multiple load cases - T&C?
    
    If orientation = 0 Then
        data5 = _
        "LOADING" & vbCrLf & _
        "1 = Number of load cases" & vbCrLf & _
        "1       1        " & Format([TOPL.Shear].Value2, "0.00000000000000E+0000") & "  " & Format([TOPL.Moment].Value2, "0.00000000000000E+0000") & "  " & Format(Range("TOPL.Axial"), "0.00000000000000E+0000") & " 1 : Load 1; BC1:Shear (lb), Moment (in-lb);          Axial Load (lb); Compute Top y vs L: 1=no, 2=yes" & vbCrLf & _
        "END" & vbCrLf
    Else
        data5 = _
        "LOADING" & vbCrLf & _
        "1 = Number of load cases" & vbCrLf & _
        "1       1        " & Format([TOPL.Shear.Weak].Value2, "0.00000000000000E+0000") & "  " & Format([TOPL.Moment.Weak].Value2, "0.00000000000000E+0000") & "  " & Format(Range("TOPL.Axial"), "0.00000000000000E+0000") & " 1 : Load 1; BC1:Shear (lb), Moment (in-lb);          Axial Load (lb); Compute Top y vs L: 1=no, 2=yes" & vbCrLf & _
        "END" & vbCrLf
    End If

'   Print/append data strings to files
    f = FreeFile
    Open fileSpec For Output As #f
    Print #f, data1a; data1b; data2; data3; data4a; data4b; data4c; data5
    Close #f

'   Delete VBA array data to avoid hanging memory issues
'    arrDepthTop() = Nothing
'    arrDepthBot() = Nothing
'    arrGamma() = Nothing
'    arrGammaEff() = Nothing
'    arrCohesion() = Nothing
'    arrPHI() = Nothing
'    arrk() = Nothing
'    arrPYcurve() = Nothing



    If batch = False Then
    '   Message box - notification that routine is complete
        'PLACEHOLDER TO ECHO MULTIPLE FILES
        strMsg = "LPile files have been created:" & vbCrLf & fileSpec & vbCrLf & vbCrLf & _
        "1. Review LPile files and Run Analysis." & vbCrLf & _
        "2. After running, click 'Import LPILE Results'"
        MsgBox strMsg, vbInformation, "Finished"

    '   Open LPile file
        'PLACEHOLDER FOR BATCHING / CLICKER / WRAPPER
        If openLpile = True Then
            b = ShellExecute(0, "Open", fileSpec, "", "", 9)
        End If
    End If

'    Application.DisplayAlerts = True

End Sub

'Process to pull data from Lpile text outputs into excel
Sub ANSgptOutput(fileSpec As String)
    
    'Define subroutine parameters
    'file variables
    Dim f As Integer ', FileSpec As String
    'Define placeholder string for whole data, variant for row data, variant for full row-column array data
    Dim LpileData As String, arrRowData As Variant, arrRowSplitData As Variant, arrRCData As Variant
    'Define string variable as a search query in larger strings
    Dim strSearchStart As String, strSearchEnd As String, line As Long, strLine As Variant
    'Define output variables from string searches
    Dim LineStart As Integer, LineEnd As Integer
    'Define variables for iteration
    Dim row As Integer, i As Integer, j As Integer
    'Define variables for writing to spreadsheet
    Dim startCol As Integer, startRow As Integer, maxWrite As Integer
        
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
    
    'declare our filename - PLACEHOLDER to bring this in automatically by a wrapper/handler function that processes all batch output
    'PLACEHOLDER TO PROCESS lp11p instead
'    If orientation = 0 Then Axis = "ST" Else Axis = "WK"
'    fileName = folderName & Dashboard.Range("Lpile.Name") & "(" & Axis & ").lp12d"
'    fileSpec = folderName & Dashboard.Range("Lpile.Name") & "(" & Axis & ").lp12d"
    
    'Open file
    f = FreeFile
    
    On Error GoTo filenotfound
    Open fileSpec For Input As #f
    On Error GoTo 0
    
    'Read the entire file to the placeholder string variable and close
    LpileData = Input(LOF(f), f)
    Close #f
    
    'Split the string data by row
    arrRowData = Split(LpileData, vbCrLf)
    
    'PLACEHOLDER FOR LPILE ERROR HANDLING - should only really need quick out once we get lpile running automatically.  There should be a reasonable amount of user control at this point in development where they'll manage the inputs to get a design that works in lpile.
    
    'PLACEHOLDER for multiple load cases in one file - depending on simplicity may want to split into separate files if that's our next step in functionality anyway
    
    'Deflection, shear, and moment
    'Find start and end of table in Lpile output
    strSearchStart = "Pile-head conditions are Shear and Moment (Loading Type 1)"
    strSearchEnd = "* The above values of total stress are combined axial and bending stresses."
    For line = 0 To UBound(arrRowData)
        If InStr(arrRowData(line), strSearchStart) > 0 Then
            LineStart = line + 10
        ElseIf InStr(arrRowData(line), strSearchEnd) > 0 Then
            LineEnd = line - 2
        End If
    Next line
    
    'Populate array with lpile output values from text
    ReDim arrRCData(0 To 9, 0 To (LineEnd - LineStart))
    For line = LineStart To LineEnd
        strLine = arrRowData(line)
        arrRowSplitData = Split(strLine, " ")
        row = 0
        For j = LBound(arrRowSplitData) To UBound(arrRowSplitData)
            If arrRowSplitData(j) <> "" Then
                arrRCData(row, line - LineStart) = Val(arrRowSplitData(j))
                row = row + 1
            End If
        Next j
    Next line
    
    'Write array to sheet
    [lpile.output2].ClearContents
    startCol = [lpile.output].Column
    startRow = [lpile.output].row
    maxWrite = UBound(arrRCData, 2)
    i = 0
    j = 0
    For i = 0 To UBound(arrRCData, 1)
        For j = 0 To UBound(arrRCData, 2)
            'ThisWorkbook.Sheets("Sheet2").Cells(startRow + i, startCol + j) = arrRCData(i, j).value
            Cells(startRow + j, startCol + i).Value = arrRCData(i, j)
        Next j
    Next i
'    With Dashboard
'        Set Destination = .Range(.Cells(startRow, startCol), .Cells(maxWrite + startRow - 1, startCol))
'        Destination.value = Application.Transpose(Application.WorksheetFunction.Index(arrRCData, 2, 0))
'    End With

    Application.Calculate
    Exit Sub
    
filenotfound:
    MsgBox fileSpec & " could not be found.", vbExclamation, "Skip File"
    [lpile.output2].ClearContents
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationAutomatic

End Sub
