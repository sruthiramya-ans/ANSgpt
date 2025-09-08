VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} settingsUI 
   Caption         =   "Settings"
   ClientHeight    =   1248
   ClientLeft      =   96
   ClientTop       =   372
   ClientWidth     =   2880
   OleObjectBlob   =   "settingsUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "settingsUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fixityDepth_Change()

    Settings.Range("Settings.FixityDepth") = fixityDepth.Value

End Sub

Private Sub gradeDeflect_Change()

    Settings.Range("Settings.GradeDefl") = gradeDeflect.Value


End Sub

Private Sub headDeflect_Change()

    Settings.Range("Settings.HeadDefl") = headDeflect.Value

End Sub

Private Sub UserForm_Initialize()

    With Me
        If HomePage.Visible = True Then
            .StartUpPosition = 0
            .Top = HomePage.Top + HomePage.Height / 1.5
            .Left = HomePage.Left + HomePage.Width / 2 - .Width / 2
        End If
    End With

    fixityDepth = Settings.Range("Settings.FixityDepth")
    gradeDeflect = Settings.Range("Settings.GradeDefl")
    headDeflect = Settings.Range("Settings.HeadDefl")
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Unload settingsUI
    HomePage.Enabled = True
    BatchAnalysis.Enabled = True
    Dashboard.Activate
    
End Sub
