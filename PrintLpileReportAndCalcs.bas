Attribute VB_Name = "PrintLpileReportAndCalcs"
Sub ExportLPileReportsToPDF()
    Dim fso As Object, folder As Object, file As Object, ts As Object
    Dim folderPath As String, lpileFileName As String
    Dim pdfPath_lpile_report As String, pdfPath_plot As String, objWord As Object, doc As Object
    Dim reportFile As String, plotFile As String
    Dim scriptPath As String, outputPDF As String
    Dim pdfNames() As String, pdfs As Variant, i As Long
    Dim bluebeamPath As String, cmd As String, line As String
    Dim newIndex As Long
    Dim BluebeamApp As Object
    Dim PDFDocument As Object, staplerPath As String, jobFile As String
    
    ' === Read folder path and LPile file name from Excel ===
    folderPath = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & ThisWorkbook.Sheets("Pile Menu").Range("A7").Value & "\"
    lpileFileName = ThisWorkbook.Sheets("Pile Menu").Range("A7").Value & "-" & ThisWorkbook.Sheets("Pile Menu").Range("B7").Value & "-Embed " & ThisWorkbook.Sheets("Pile Menu").Range("E7").Value & "ft-" & ThisWorkbook.Sheets("Pile Menu").Range("C7").Value & " mil-Soil " & ThisWorkbook.Sheets("Pile Menu").Range("J7").Value & "-Scour " & ThisWorkbook.Sheets("Pile Menu").Range("K7").Value & "Strong"
    
    ' Build expected file names
    reportFile = folderPath & lpileFileName & ".lp12o"
    plotFile = folderPath & lpileFileName & ".lp12p"
    
    ' Create Word object
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
    
    ' === Export Report file (.lp12o) ===
    If Dir(reportFile) <> "" Then
        Set doc = objWord.Documents.Open(reportFile, ReadOnly:=True)
        pdfPath_lpile_report = folderPath & lpileFileName & "_Report.pdf"
        doc.ExportAsFixedFormat OutputFileName:=pdfPath_lpile_report, _
                                ExportFormat:=17 ' wdExportFormatPDF
        doc.Close False
    Else
        MsgBox "Report file not found: " & reportFile, vbExclamation
    End If
    
    ' === Export Plot file (.lp12p) ===
    If Dir(plotFile) <> "" Then
        Set doc = objWord.Documents.Open(plotFile, ReadOnly:=True)
        pdfPath_plots = folderPath & lpileFileName & "_Plots.pdf"
        doc.ExportAsFixedFormat OutputFileName:=pdfPath_plots, _
                                ExportFormat:=17 ' wdExportFormatPDF
        doc.Close False
    Else
        MsgBox "Plot file not found: " & plotFile, vbExclamation
    End If
    
    ' === Export Specific Excel Sheets to PDF ===
    ' List sheet names here (adjust as needed)
    sheetNames = Array("AG", "AM", "Soil Axial")
    
     ' Redim array to hold pdf names
    ReDim pdfNames(LBound(sheetNames) To UBound(sheetNames))
    
    ' === Loop through sheets and export each as PDF ===
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            pdfNames(i) = folderPath & lpileFileName & "_" & ws.Name & ".pdf"
            ws.Select
            ws.ExportAsFixedFormat Type:=xlTypePDF, _
                                   fileName:=pdfNames(i), _
                                   Quality:=xlQualityStandard, _
                                   IncludeDocProperties:=True, _
                                   IgnorePrintAreas:=False, _
                                   OpenAfterPublish:=False
        Else
            MsgBox "Sheet not found: " & sheetNames(i), vbExclamation
        End If
        
        Set ws = Nothing
    Next i
    
    ' === List of PDFs to merge (adjust paths as needed) ===
    newIndex = UBound(pdfNames) + 1
    
    ' Resize array (Preserve keeps existing values)
    ReDim Preserve pdfNames(0 To newIndex)
    
    ' Add new value
    pdfNames(newIndex) = pdfPath_lpile_report
    
    
    outputPDF = folderPath & "Merged.pdf"
    jobFile = folderPath & "MergeJob.job"
    
'    ' === Create Stapler job file ===
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set ts = fso.CreateTextFile(jobFile, True)
'
'    ts.WriteLine "[Job]"
'    ts.WriteLine "OutputFile=" & outputPDF
'    ts.WriteLine "Files="
'    For i = LBound(pdfNames) To UBound(pdfNames)
'        ts.WriteLine pdfNames(i)
'    Next i
'    ts.Close
    
    ' === Path to Stapler ===
    staplerPath = "C:\Program Files\Bluebeam Software\Bluebeam Revu\21\Revu\Stapler.exe"
    
    ok = MergeWithBluebeamStapler(pdfNames, outputPDF, staplerPath, jobFile)

    If ok Then
        MsgBox "Merged successfully: " & outputPDF, vbInformation
    Else
        MsgBox "Merge failed or timed out." & vbCrLf & outputPDF, vbExclamation
    End If
    
'    ' === Run Stapler with job file ===
'    cmd = staplerPath & " /a """ & jobFile & """"
'    Shell cmd, vbNormalFocus
    
''    scriptPath = folderPath & "MergedPDFs.txt"
''
''    ' === Build the Bluebeam script ===
''    Set fso = CreateObject("Scripting.FileSystemObject")
''    Set ts = fso.CreateTextFile(scriptPath, True)
''
''
''    ts.WriteLine "CombinePDFs"
''
''    For i = LBound(pdfNames) To UBound(pdfNames)
''        ts.WriteLine """" & pdfNames(i) & """"
''    Next i
''
''    ts.WriteLine """" & outputPDF & """"
''    ts.Close
''
''    ' === Path to Bluebeam Revu executable (adjust version if needed) ===
''    bluebeamPath = """C:\Program Files\Bluebeam Software\Bluebeam Revu\21\Revu\Revu.exe"""
''
''    ' === Run Bluebeam with the script ===
''    cmd = bluebeamPath & " /s " & """" & scriptPath & """"
''    Shell cmd, vbNormalFocus
'
'    ' === Path to Bluebeam Stapler.exe (adjust if needed) ===
'    staplerPath = """C:\Program Files\Bluebeam Software\Bluebeam Revu\21\Revu\Stapler.exe"""
'
'    ' === Build command line ===
'    cmd = staplerPath
'    For i = LBound(pdfNames) To UBound(pdfNames)
'        cmd = cmd & " """ & pdfNames(i) & """"
'    Next i
'    cmd = cmd & " """ & outputPDF & """"   ' output at the end
'
'    ' === Run Stapler to merge PDFs ===
'    Shell cmd, vbNormalFocus

    
    MsgBox "Merge started in Bluebeam. Check output: " & outputPDF, vbInformation
End Sub


' === Helper that does the heavy lifting ======================================
Private Function MergeWithBluebeamStapler(pdfs As Variant, _
                                          outputPDF As String, _
                                          staplerExe As String, jobFile As String) As Boolean
    Dim fso As Object, ts As Object, wsh As Object
    Dim i As Long
    Dim started As Boolean, activated As Boolean
    Dim tEnd As Single

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")

    ' Delete any existing output to avoid overwrite prompts
    On Error Resume Next
    If fso.FileExists(outputPDF) Then fso.DeleteFile outputPDF, True
    On Error GoTo 0

    ' Create a temporary Stapler job file
    Set ts = fso.CreateTextFile(jobFile, True)
    ts.WriteLine "[Job]"
    ts.WriteLine "OutputFile=" & outputPDF
    ts.WriteLine "Files="
    For i = LBound(pdfs) To UBound(pdfs)
        ts.WriteLine CStr(pdfs(i))
    Next i
    ts.Close


    For i = LBound(pdfs) To UBound(pdfs)
        If LCase(fso.GetExtensionName(pdfs(i))) <> "pdf" Then
            MsgBox "Unsupported file type: " & pdfs(i)
            Exit Function
        End If
    Next i

    ' Ensure no previous Stapler instance is running
    wsh.Run "taskkill /IM Stapler.exe /F", 0, True
    WaitSeconds 0.5 ' Give it a moment to close

    ' Launch Stapler with the job file (interactive UI)
    wsh.Run """" & staplerExe & """ """ & jobFile & """", 1, True

    ' Try to activate the Stapler window (title can vary by version)
    tEnd = Timer + 10    ' 10s to find the window
    Do While Timer < tEnd And Not activated
        DoEvents
        activated = wsh.AppActivate("Bluebeam Stapler")
        If Not activated Then activated = wsh.AppActivate("Stapler")
    Loop

    ' Press the Staple button:
    '   - First try Alt+S (common accelerator for Staple)
    '   - Then send Enter in case Staple is the default button
    If activated Then
        ' give it a beat to settle
        WaitSeconds 0.5
        wsh.SendKeys "%s"      ' Alt+S
        WaitSeconds 0.5
        wsh.SendKeys "~"       ' Enter (as a fallback)
    End If

    ' Wait up to 90s for output to appear
    tEnd = Timer + 90
    Do While Timer < tEnd And Not fso.FileExists(outputPDF)
        DoEvents
        WaitSeconds 0.25
    Loop

    MergeWithBluebeamStapler = fso.FileExists(outputPDF)

    ' Optional: close Stapler
    If activated Then
        wsh.AppActivate "Bluebeam Stapler"
        WaitSeconds 0.3
        wsh.SendKeys "%{F4}"   ' Alt+F4
    End If

    ' Clean up temp job
    On Error Resume Next
    If fso.FileExists(jobFile) Then fso.DeleteFile jobFile, True
    On Error GoTo 0
    wsh.Run "taskkill /IM Stapler.exe /F", 0, True

End Function

' Small wait helper that doesn't freeze Excel
Private Sub WaitSeconds(ByVal seconds As Double)
    Dim t As Single: t = Timer + seconds
    Do While Timer < t
        DoEvents
    Loop
End Sub


