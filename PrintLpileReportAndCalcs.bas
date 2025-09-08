Attribute VB_Name = "PrintLpileReportAndCalcs"
Sub GenerateIndividualReportsAndStaplerBSXFiles()
    Dim ws As Worksheet
    Dim outputFolderName As String, folderPath As String, pileType As String, lpileFileNameStrong As String, lpileFileNameWeak As String
    Dim lastRow As Long
    Dim i As Long
    Dim concatValue As String
    
    outputFolderName = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & "Output Reports\"
    CreateDir (outputFolderName)
    ' Set worksheet to work on
    Set ws = ThisWorkbook.Sheets("Pile Menu")
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Loop through each row starting from row 7
    For i = 7 To lastRow
        folderPath = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & ws.Cells(i, 1).Value & "\"
        pileType = ws.Cells(i, 1).Value & "-" & ws.Cells(i, 2).Value & "-Embed " & ws.Cells(i, 5).Value & "ft-" & ws.Cells(i, 3).Value & " mil-Soil " & ws.Cells(i, 10).Value & "-Scour " & ws.Cells(i, 11).Value
        lpileFileNameStrong = ws.Cells(i, 1).Value & "-" & ws.Cells(i, 2).Value & "-Embed " & ws.Cells(i, 5).Value & "ft-" & ws.Cells(i, 3).Value & " mil-Soil " & ws.Cells(i, 10).Value & "-Scour " & ws.Cells(i, 11).Value & "Strong"
        lpileFileNameWeak = ws.Cells(i, 1).Value & "-" & ws.Cells(i, 2).Value & "-Embed " & ws.Cells(i, 5).Value & "ft-" & ws.Cells(i, 3).Value & " mil-Soil " & ws.Cells(i, 10).Value & "-Scour " & ws.Cells(i, 11).Value & "Weak"
        folderName = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & "Output Reports\" & ws.Cells(i, 1).Value & "\"
    
        ' Recalculate the application before printing
        Dashboard.Range("Pile.Type") = ws.Cells(i, 1).Value
        Dashboard.Range("Pile.Shape") = ws.Cells(i, 2).Value
        Dashboard.Range("Pile.Galv") = ws.Cells(i, 3).Value
        Dashboard.Range("Pile.Reveal") = ws.Cells(i, 4).Value
        Dashboard.Range("Pile.Embed") = ws.Cells(i, 5).Value
        Dashboard.Range("Scour.Zone") = ws.Cells(i, 11).Value
        Dashboard.Range("Soil.Zone") = ws.Cells(i, 10).Value
        
        Dashboard.Range("Load.AGM") = ws.Cells(i, 12).Value
        Dashboard.Range("Load.AGS") = ws.Cells(i, 13).Value
        Dashboard.Range("Load.AMM") = ws.Cells(i, 14).Value
        Dashboard.Range("Load.AGM.Weak") = ws.Cells(i, 15).Value
        Dashboard.Range("Load.AGS.Weak") = ws.Cells(i, 16).Value
        Dashboard.Range("Load.AMM.Weak") = ws.Cells(i, 17).Value
        
        Application.Calculate
        Call CreateIndividualMergedReport(outputFolderName, folderPath, pileType, lpileFileNameStrong, lpileFileNameWeak)

    Next i
End Sub
Sub CreateIndividualMergedReport(outputFolderName As String, folderPath As String, pileType As String, lpileFileNameStrong As String, lpileFileNameWeak As String)
    Dim fso As Object, folder As Object, file As Object, ts As Object
    Dim pdfPath_lpile_report_strong As String, pdfPath_lpile_report_weak As String, objWord As Object, doc As Object
    Dim reportFileStrong As String, reportFileWeak As String
    Dim scriptPath As String, outputPDF As String
    Dim pdfNames() As String, i As Long
    Dim bluebeamPath As String, cmd As String, line As String
    Dim newIndex As Long
    Dim BluebeamApp As Object
    Dim PDFDocument As Object, staplerPath As String, jobFile As String
    Dim staplerExe As String
    Dim pdfs(4) As String
       
    ' Build expected file names
    reportFileStrong = folderPath & lpileFileNameStrong & ".lp12o"
    reportFileWeak = folderPath & lpileFileNameWeak & ".lp12o"
    ' Create Word object
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
    
    ' === Export Strong lpile Report file (.lp12o) ===
    If Dir(reportFileStrong) <> "" Then
        Set doc = objWord.Documents.Open(reportFileStrong, ReadOnly:=True)
        pdfPath_lpile_report_strong = folderPath & lpileFileNameStrong & "_Report.pdf"
        doc.ExportAsFixedFormat OutputFileName:=pdfPath_lpile_report_strong, _
                                ExportFormat:=17 ' wdExportFormatPDF
        doc.Close False
    Else
        MsgBox "Report file not found: " & reportFileStrong, vbExclamation
    End If
    
    ' === Export Weak lpile Report file (.lp12o) ===
    If Dir(reportFileWeak) <> "" Then
        Set doc = objWord.Documents.Open(reportFileWeak, ReadOnly:=True)
        pdfPath_lpile_report_weak = folderPath & lpileFileNameWeak & "_Report.pdf"
        doc.ExportAsFixedFormat OutputFileName:=pdfPath_lpile_report_weak, _
                                ExportFormat:=17 ' wdExportFormatPDF
        doc.Close False
    Else
        MsgBox "Report file not found: " & reportFileWeak, vbExclamation
    End If
    
    ' === Export Specific Excel Sheets to PDF ===
    ' List sheet names here
    sheetNames = Array("AG", "AM", "Soil Axial")
    
     ' Redim array to hold pdf names
    ReDim pdfNames(LBound(sheetNames) To UBound(sheetNames))
    
    ' === Loop through sheets and export each as PDF ===
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            pdfNames(i) = folderPath & pileType & "_" & ws.Name & ".pdf"
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
    
    ' === List of PDFs to merge ===
    newIndex = UBound(pdfNames) + 1
    
    ' Resize array (Preserve keeps existing values)
    ReDim Preserve pdfNames(0 To newIndex)
    
    ' Add new value
    pdfNames(newIndex) = pdfPath_lpile_report_strong
    
    ' === List of PDFs to merge ===
    newIndex = UBound(pdfNames) + 1
    
    ' Resize array (Preserve keeps existing values)
    ReDim Preserve pdfNames(0 To newIndex)
    
    ' Add new value
    pdfNames(newIndex) = pdfPath_lpile_report_weak
    
    
    outputPDF = outputFolderName & pileType & "_Merged.pdf"
    jobFile = folderPath & pileType & "_MergeJob.bsx"
    
    ' Path for stapler exe
    staplerExe = "C:\Program Files\Bluebeam Software\Bluebeam Revu\21\Revu\Stapler.exe"
    
    ' Input PDFs
    pdfs(0) = pdfNames(0)  'excel sheet AG
    pdfs(1) = pdfNames(1)  'excel sheet AM
    pdfs(2) = pdfPath_lpile_report_strong  'Lpile report strong axis
    pdfs(3) = pdfPath_lpile_report_weak  'Lpile report weak axis
    pdfs(4) = pdfNames(2)  'excel sheet Soil Axial
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Delete old files
    On Error Resume Next
    If fso.FileExists(outputPDF) Then fso.DeleteFile outputPDF, True
    If fso.FileExists(jobFile) Then fso.DeleteFile jobFile, True
    On Error GoTo 0
    
    ' Create .bsx file
    Set ts = fso.CreateTextFile(jobFile, True)
    ts.WriteLine "<?xml version=""1.0"" encoding=""utf-8""?>"
    ts.WriteLine "<Jobs>"
    ts.WriteLine "  <Job>"
    ts.WriteLine "    <OutputFileName>" & fso.GetFileName(outputPDF) & "</OutputFileName>"
    ts.WriteLine "    <StampsOnAllPages />"
    ts.WriteLine "    <OutputDir>" & fso.GetParentFolderName(outputPDF) & "</OutputDir>"
    ts.WriteLine "    <JobOptions>"
    ts.WriteLine "      <Name>Standard Document.joboptions</Name>"
    ts.WriteLine "      <Width>-1</Width>"
    ts.WriteLine "      <Height>-1</Height>"
    ts.WriteLine "      <Orient>Auto</Orient>"
    ts.WriteLine "      <UserRotation>0</UserRotation>"
    ts.WriteLine "      <ImageCompression>Flate</ImageCompression>"
    ts.WriteLine "      <ImageResolution>300</ImageResolution>"
    ts.WriteLine "      <JpegQuality>75</JpegQuality>"
    ts.WriteLine "      <ImageAliasingText>2</ImageAliasingText>"
    ts.WriteLine "      <ImageAliasingGraphics>2</ImageAliasingGraphics>"
    ts.WriteLine "      <LineMergeOn>False</LineMergeOn>"
    ts.WriteLine "      <BlendMode>Darken</BlendMode>"
    ts.WriteLine "      <BlendAlpha>1</BlendAlpha>"
    ts.WriteLine "      <PDFPostProcess>False</PDFPostProcess>"
    ts.WriteLine "      <PostProcessProcessMasks>False</PostProcessProcessMasks>"
    ts.WriteLine "      <PostProcessFixStripedImageTransparency>False</PostProcessFixStripedImageTransparency>"
    ts.WriteLine "      <PostProcessCombineAdjacentImages>False</PostProcessCombineAdjacentImages>"
    ts.WriteLine "      <PostProcessOptimizeSolidImages>False</PostProcessOptimizeSolidImages>"
    ts.WriteLine "      <PostProcessRemoveTextClipping>False</PostProcessRemoveTextClipping>"
    ts.WriteLine "      <PostProcessSimplifyClippingPaths>False</PostProcessSimplifyClippingPaths>"
    ts.WriteLine "      <PostProcessPDFVersion>Version_1_4</PostProcessPDFVersion>"
    ts.WriteLine "    </JobOptions>"
    ts.WriteLine "    <ColorDepth>4</ColorDepth>"
    ts.WriteLine "    <OpenOutputFileAfter>True</OpenOutputFileAfter>"
    ts.WriteLine "    <DeleteTempPS>False</DeleteTempPS>"
    ts.WriteLine "    <Name />"
    ts.WriteLine "    <Overwrite>0</Overwrite>"
    ts.WriteLine "    <Delete>False</Delete>"
    ts.WriteLine "    <InterpreterType />"
    ts.WriteLine "    <LastError />"
    ts.WriteLine "    <Unfiltered>False</Unfiltered>"
    
    ' Add SubJobs for each PDF
    For i = 0 To UBound(pdfs)
        ts.WriteLine "    <SubJob>"
        ts.WriteLine "      <OriginalFileName>" & pdfs(i) & "</OriginalFileName>"
        ts.WriteLine "      <InputFileName>" & pdfs(i) & "</InputFileName>"
        ts.WriteLine "      <InputFileType>.pdf</InputFileType>"
        ts.WriteLine "      <ExeName>Revu</ExeName>"
        ts.WriteLine "      <ApplicationTitle />"
        ts.WriteLine "      <PageSize />"
        ts.WriteLine "      <Orientation />"
        ts.WriteLine "      <Scale />"
        ts.WriteLine "      <TransferBookmarks>False</TransferBookmarks>"
        ts.WriteLine "      <TransferHyperlinks>False</TransferHyperlinks>"
        ts.WriteLine "      <TransferFileProperties>False</TransferFileProperties>"
        ts.WriteLine "      <Message />"
        ts.WriteLine "      <Stamps />"
        ts.WriteLine "    </SubJob>"
    Next i
    
    ts.WriteLine "  </Job>"
    ts.WriteLine "</Jobs>"
    ts.Close

    'Call CreateStaplerJobBSX(outputPDF, jobFile, pileType, pdfNames(0), pdfNames(1), pdfPath_lpile_report_strong, pdfPath_lpile_report_weak, pdfNames(2))

End Sub
Sub BatchStaplerJobBSX()
    Dim ws As Worksheet
    Dim outputFolderName As String, folderPath As String, pileType As String, pdfPath_lpile_report_strong As String, pdfPath_lpile_report_weak As String, outputPDF As String, jobFile As String
    Dim lastRow As Long
    Dim i As Long
    Dim concatValue As String
    Dim pdfNames(4) As String
    
    outputFolderName = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & "Output Reports\"
    ' Set worksheet to work on
    Set ws = ThisWorkbook.Sheets("Pile Menu")
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Loop through each row starting from row 7
    For i = 7 To lastRow
        folderPath = EnsureFolderExists(Range("LPILE.Folder") & "\" & Range("Project.Name")) & ws.Cells(i, 1).Value & "\"
        pileType = ws.Cells(i, 1).Value & "-" & ws.Cells(i, 2).Value & "-Embed " & ws.Cells(i, 5).Value & "ft-" & ws.Cells(i, 3).Value & " mil-Soil " & ws.Cells(i, 10).Value & "-Scour " & ws.Cells(i, 11).Value
        outputPDF = outputFolderName & pileType & "_Merged.pdf"
        jobFile = folderPath & pileType & "_MergeJob.bsx"
        pdfNames(0) = folderPath & pileType & "_AG.pdf"
        pdfNames(1) = folderPath & pileType & "_AM.pdf"
        pdfPath_lpile_report_strong = folderPath & pileType & "Weak_Report.pdf"
        pdfPath_lpile_report_weak = folderPath & pileType & "Strong_Report.pdf"
        pdfNames(2) = folderPath & pileType & "_Soil Axial.pdf"
        Call CreateStaplerJobBSX(outputPDF, jobFile, pileType, pdfNames(0), pdfNames(1), pdfPath_lpile_report_strong, pdfPath_lpile_report_weak, pdfNames(2))
    Next i

End Sub


