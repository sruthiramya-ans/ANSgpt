VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DropTOPLs 
   Caption         =   "TOPL Importer"
   ClientHeight    =   3288
   ClientLeft      =   72
   ClientTop       =   288
   ClientWidth     =   4224
   OleObjectBlob   =   "DropTOPLs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DropTOPLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub doImport_Click()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    If Me.revealHeightBox = "" Or Me.sheetsListBox.Value = "" Then
        MsgBox "Missing sheet or reveal height. Please select/input and import again.", vbCritical, "Missing Input"
        Exit Sub
    End If
    
    strPath = filePath.Caption
    TOPLs.Range("TOPL.filepath") = strPath
    
    If overwriteRadio.Value = True Then TOPLs.Range("TOPL.data").ClearContents
    
    importTOPLs
    
    TOPLs.Range("TOPL.import.TF") = "TRUE"
    
    Application.EnableEvents = True
    
End Sub

Private Sub doImportClose_Click()

    'copy TOPLs to sheet
    Call doImport_Click
    
    Unload Me
    
End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    filePath.Caption = Data.Files(1)
    doImport.Enabled = True
    doImportClose.Enabled = True
    
    Dim analysisBook As Workbook, TOPLbook As Workbook, TOPLsheet As Worksheet
    
    Set analysisBook = ThisWorkbook
    TOPLpath = filePath.Caption
    Set TOPLbook = Workbooks.Open(TOPLpath)
    
    analysisBook.Activate
    sheetsListBox.Clear
    For n = 1 To TOPLbook.Sheets.Count
        sheetsListBox.AddItem TOPLbook.Sheets(n).Name
    Next n
    TOPLbook.Close savechanges:=False
    
    Application.EnableEvents = True
    
End Sub

Private Sub UserForm_Initialize()

    'Enable drag and drop
    TreeView1.OLEDropMode = ccOLEDropManual
    doImport.Enabled = False
    doImportClose.Enabled = False
    'filePath.Caption = TOPLs.Range("TOPL.filepath")
    TOPLs.Activate

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Unload DropTOPLs
    HomePage.Enabled = True
    BatchAnalysis.Enabled = True
    Dashboard.Activate
    
    Application.EnableEvents = True
    
End Sub
