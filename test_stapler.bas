Attribute VB_Name = "test_stapler"
Private Sub CreateStaplerJobBSX()
    Dim staplerExe As String
    Dim outputPDF As String
    Dim jobFile As String
    Dim pdfs(2) As String
    Dim fso As Object, ts As Object, wsh As Object
    Dim i As Integer
    
    ' Paths
    staplerExe = "C:\Program Files\Bluebeam Software\Bluebeam Revu\21\Revu\Stapler.exe"
    outputPDF = "C:\Temp\Merged.pdf"
    jobFile = "C:\Temp\MergeJob.bsx"
    
    ' Input PDFs
    pdfs(0) = "C:\Temp\pdf1.pdf"
    pdfs(1) = "C:\Temp\pdf2.pdf"
    pdfs(2) = "C:\Temp\pdf3.pdf"
    
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
    
    ' Run Stapler with the bsx file
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & staplerExe & """ """ & jobFile & """", 1, False
    
    MsgBox "Stapler job created at " & jobFile, vbInformation
End Sub
