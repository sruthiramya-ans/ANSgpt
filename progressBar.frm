VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} progressBar 
   Caption         =   "Loading..."
   ClientHeight    =   288
   ClientLeft      =   36
   ClientTop       =   288
   ClientWidth     =   7488
   OleObjectBlob   =   "progressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "progressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    With Me
    
        .StartUpPosition = 0
        .Left = Application.Left + Application.Width / 2 - Me.Width / 2
        .Top = Application.Top + Application.Height / 2 - Me.Height / 2
    
    End With
    
End Sub
