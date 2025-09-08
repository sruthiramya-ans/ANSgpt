VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loading 
   Caption         =   "Loading"
   ClientHeight    =   624
   ClientLeft      =   96
   ClientTop       =   372
   ClientWidth     =   2100
   OleObjectBlob   =   "loading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    With Me
        If HomePage.Visible = True Then
            .StartUpPosition = 0
            .Top = HomePage.Top + HomePage.Height / 3
            .Left = HomePage.Left + HomePage.Width / 2 - .Width / 2
        End If
    End With
    
    Dim i As Integer
   
    i = [RandBetween(6,16)]
    
    waitPhrase.Caption = Settings.Cells(i, Settings.Range("Settings.Waitlist").Column)
    
End Sub
