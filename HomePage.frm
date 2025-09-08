VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HomePage 
   Caption         =   "ANSgpt Home"
   ClientHeight    =   9084.001
   ClientLeft      =   60
   ClientTop       =   180
   ClientWidth     =   4644
   OleObjectBlob   =   "HomePage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub analysisMode_Click()

    If analysisMode.Value = True Then
        analysisMode.Caption = "Batch Run"
        analysisMode.BackColor = RGB(205, 91, 43)
        createLPILE.BackColor = RGB(205, 91, 43)
        importLPILE.BackColor = RGB(205, 91, 43)
        BatchAnalysis.Show vbModeless
    Else
        analysisMode.Caption = "Single Run"
        analysisMode.BackColor = RGB(170, 179, 164)
        createLPILE.BackColor = RGB(170, 179, 164)
        importLPILE.BackColor = RGB(170, 179, 164)
        Unload BatchAnalysis
    End If

End Sub

Private Sub createFixity_Click()

    Dashboard.Activate
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    'Create fixity files
    Call CreateFixityLPILE

    'Open lpile
    fileName = Settings.Range("LPILE.Folder") & "\" & Dashboard.Range("Project.Name") & "\Fixity\" & Dashboard.Range("Lpile.Name") & ".lp12d"
    openLpile = ShellExecute(0, "Open", fileName, "", "", 10)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Sub createLPILE_Click()

    Dashboard.Activate
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    isBatch = analysisMode.Value
    
    If isBatch = False Then
        With Dashboard
            .Range("Lpile.Name") = .Range("Pile.Shape") & "-Embed " & .Range("Pile.Embed") & " ft-Reveal " & .Range("Pile.Reveal") & " ft-" & _
                                    .Range("Pile.Galv") & "mil-" & .Range("Soil.Zone") & "-" & .Range("Scour.Zone")
        End With
        loading.Show vbModeless
        'KEY ANSgptCreator(BATCH, OVER, OPEN, FIXITY, FOLDER, S/W 0/1)
        Call ANSgptCreator(False, True, True, False, , 0) 'strong
        Call ANSgptCreator(False, True, True, False, , 1) 'weak

        Unload loading
    Else
        If Settings.Range("Settings.BatchReady") <> True Then
            MsgBox "Missing at one or more multiselect value from Batch Analysis Options. Please select needed options and try again.", vbCritical, "Missing Input"
            Exit Sub
        End If
        BatchResults.Range("Batch.ImportedTF") = False
        generateBatchList
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub importFixity_Click()

    Dashboard.Activate
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    'Import fixity Lpile output files
    ImportFixityLPILE
       
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Sub importLPILE_Click()

    Dashboard.Activate
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    isBatch = analysisMode.Value
        
    If isBatch = False Then
        loading.Show vbModeless
        Application.Wait Now + #12:00:01 AM# 'gives userform 1 second to load because reasons
        If InStr(1, Dashboard.Range("Lpile.Name"), "Fixity Check") > 0 Then
            fileName = Settings.Range("LPILE.Folder") & "\" & Dashboard.Range("Project.Name") & "\Fixity\" & Dashboard.Range("Lpile.Name") & ".lp12o"
        Else
            fileNameSt = Settings.Range("LPILE.Folder") & "\" & Dashboard.Range("Project.Name") & "\Single Run\" & Dashboard.Range("Lpile.Name") & "(ST).lp12o" 'strong
            fileNameWk = Settings.Range("LPILE.Folder") & "\" & Dashboard.Range("Project.Name") & "\Single Run\" & Dashboard.Range("Lpile.Name") & "(WK).lp12o" 'weak
        End If
        ANSgptOutput (fileNameSt)
        ANSgptOutputWeak (fileNameWk)
        Unload loading
    Else
        If Settings.Range("Settings.BatchReady") <> True Then
            MsgBox "Missing at one or more multiselect value from Batch Analysis Options. Please select needed options and try again.", vbCritical, "Missing Input"
            Exit Sub
        End If
        'If batchSummary.Visible = True Then Unload batchSummary
        importBatchList
        BatchResults.Range("Batch.ImportedTF") = True
        
        batchSummary.Show vbModeless
    
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Sub importTOPLs_Click()

    Dashboard.Activate
    HomePage.Enabled = False
    BatchAnalysis.Enabled = False
    DropTOPLs.Show vbModeless
    
End Sub


Private Sub projectLpileFolder_Click()

    Shell "explorer.exe" & " " & Settings.Range("LPILE.Folder") & "\" & Dashboard.Range("Project.Name"), vbNormalFocus

End Sub

Private Sub resetTool_Click()

    answer = MsgBox("Are you sure you want to clear all inputs and results? This action cannot be undone.", vbYesNo, "Reset Tool")
    
    If answer = vbYes Then
        loading.Show vbModeless
        Application.Wait Now + #12:00:01 AM# 'gives userform 1 second to load because reasons
        Application.EnableEvents = False
        Dim cel As Range
        On Error Resume Next
        'Clear inputs on Dashboard
        For Each cel In Dashboard.Range("A1:K52")
            If cel.Interior.Color = RGB(255, 230, 153) Then cel.ClearContents
            Dashboard.Range("Project.Name") = ""
            Dashboard.Range("Pile.Type") = ""
        Next cel
        'Clear inputs on Soil Zones
        For Each cel In SoilZones.Range("A1:L1000")
            If cel.Interior.Color = RGB(255, 230, 153) Then cel.ClearContents
        Next cel
        TOPLs.Range("TOPL.data").ClearContents
        TOPLs.Range("TOPL.import.TF") = False
        
        FixityResults.Range("Fixity.Results").ClearContents
        
        BatchResults.Range("Batch.Data").ClearContents
        BatchResults.Range("Batch.ImportedTF") = False
        
        PileMenu.Range("Menu.Full").ClearContents
        
        Settings.Range("Settings.BatchOptions").ClearContents
        Settings.Range("Settings.BatchReady") = False
        
        Application.EnableEvents = True
        Unload loading
    End If
    
End Sub

Private Sub saveQuit_Click()

    Unload Me
    
    ThisWorkbook.Close savechanges:=True

End Sub

Private Sub showResults_Click()

    batchSummary.Show vbModeless

End Sub

Private Sub showSettings_Click()

    HomePage.Enabled = False
    BatchAnalysis.Enabled = False
    settingsUI.Show vbModeless

End Sub

Private Sub strongButton_Click()

End Sub

Private Sub UserForm_Initialize()
   
    Call GetUser
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    With Me
    
        .StartUpPosition = 0
        .Top = Application.Top + Application.Height / 2 - .Height / 2
        .Left = Application.Left + Application.Width / 2 - Width / 2
    
    End With
 
End Sub

Private Sub UserForm_Layout()
    
    If BatchAnalysis.Visible = True Then
        With BatchAnalysis
            .StartUpPosition = 0
            .Top = HomePage.Top
            .Left = HomePage.Left + HomePage.Width
        End With
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If BatchAnalysis.Visible = True Then
        Unload BatchAnalysis
    End If

End Sub
