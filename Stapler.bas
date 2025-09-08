Attribute VB_Name = "Stapler"
Public Sub CreateStaplerJobBSX(outputPDF As String, jobFile As String, pileType As String, pdf1 As String, pdf2 As String, pdf3 As String, pdf4 As String, pdf5 As String)
    Dim staplerExe As String
    Dim pdfs(4) As String
    Dim fso As Object, ts As Object
    Dim i As Integer
      
    ' === Step 1: Open Stapler with the job file ===
    ShellExecute 0, "open", jobFile, vbNullString, vbNullString, 1
    
    ' === Step 2: User manually staples ===
    MsgBox "Stapler for " & pileType & " will open soon. Please staple manually, review in Bluebeam, then close Bluebeam when done." & vbCrLf & vbCrLf & _
           "Click OK in Excel once finished.", vbInformation, "Manual Staple Step"
    
'    ' === Step 5 & 6: Notify and wait for user ===
'    If fso.FileExists(outputPDF) Then
'        MsgBox "Loop report generated successfully: " & vbCrLf & outputPDF & vbCrLf & vbCrLf & _
'               "Press OK to continue with the next report.", vbInformation, "Report Complete"
'    Else
'        MsgBox "No merged PDF found for this loop. Please check Stapler/Revu.", vbExclamation, "Warning"
'    End If
    
End Sub
' === Helper wait function ===
Private Sub WaitSeconds(ByVal seconds As Double)
    Dim t As Single: t = Timer + seconds
    Do While Timer < t
        DoEvents
    Loop
End Sub
