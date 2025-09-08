VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BatchAnalysis 
   Caption         =   "Batch Analysis Options"
   ClientHeight    =   7476
   ClientLeft      =   -108
   ClientTop       =   -336
   ClientWidth     =   4800
   OleObjectBlob   =   "BatchAnalysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BatchAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ProcessBasedOnBendingAxis()
    Dim axis As String

    ' Read from the userform
    If strongButton.Value = True Then
        axis = "Strong"
    Else
        axis = "Weak"
    End If

    Settings.Range("Settings.axis") = axis
End Sub

'--- call this from each ListBox_Click
Private Sub ToggleSelectAll(lb As MSForms.ListBox)
    Const IDX_SELECT_ALL As Long = 0
    Dim i As Long
    With lb
        If .ListIndex = IDX_SELECT_ALL Then
            ' User clicked “Select All” ? set all others to match it
            For i = 1 To .ListCount - 1
                .Selected(i) = .Selected(IDX_SELECT_ALL)
            Next i
        Else
            ' User clicked a normal item ? refresh the “Select All” box
            Dim allOn As Boolean: allOn = True
            For i = 1 To .ListCount - 1
                If Not .Selected(i) Then
                    allOn = False
                    Exit For
                End If
            Next i
            .Selected(IDX_SELECT_ALL) = allOn
        End If
    End With
End Sub

Private Sub Frame7_Click()

End Sub

Private Sub Galv_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ToggleSelectAll Me.Galv
    
    'Clear list
    Settings.Range("Settings.GalvList").ClearContents
    
    Dim setCel As Range
    Set setCel = Settings.Cells(Settings.Range("Settings.GalvList").row, Settings.Range("Settings.GalvList").Column)
    
    'Store settings
    With Me
       For n = 1 To .Galv.ListCount - 1
            If .Galv.Selected(n) Then
                setCel = .Galv.List(n)
                Set setCel = setCel.Offset(1, 0)
            End If
        Next n
    End With
    
End Sub

Private Sub GEO_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ToggleSelectAll Me.GEO
    
    'Clear list
    Settings.Range("Settings.GeoList").ClearContents
    
    Dim setCel As Range
    Set setCel = Settings.Cells(Settings.Range("Settings.GeoList").row, Settings.Range("Settings.GeoList").Column)
    
    'Store settings
    With Me
       For n = 1 To .GEO.ListCount - 1
            If .GEO.Selected(n) Then
                setCel = .GEO.List(n)
                Set setCel = setCel.Offset(1, 0)
            End If
        Next n
    End With
    
End Sub

Private Sub intEmbed_Change()

    Settings.Range("Settings.intEmbed") = intEmbed.Value

End Sub

Private Sub Label2_Click()

End Sub

Private Sub maxEmbed_Change()

    Settings.Range("Settings.maxEmbed") = maxEmbed.Value

End Sub

Private Sub minEmbed_Change()

    Settings.Range("Settings.minEmbed") = minEmbed.Value

End Sub

Private Sub SCOUR_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ToggleSelectAll Me.SCOUR
    
    'Clear list
    Settings.Range("Settings.ScourList").ClearContents
    
    Dim setCel As Range
    Set setCel = Settings.Cells(Settings.Range("Settings.ScourList").row, Settings.Range("Settings.ScourList").Column)
    
    'Store settings
    With Me
       For n = 1 To .SCOUR.ListCount - 1
            If .SCOUR.Selected(n) Then
                setCel = .SCOUR.List(n)
                Set setCel = setCel.Offset(1, 0)
            End If
        Next n
    End With
    
End Sub

Private Sub SHAPES_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ToggleSelectAll Me.SHAPES
    
    'Clear list
    Settings.Range("Settings.ShapesList").ClearContents
    
    Dim setCel As Range
    Set setCel = Settings.Cells(Settings.Range("Settings.ShapesList").row, Settings.Range("Settings.ShapesList").Column)
    
    'Store settings
    With Me
       For n = 1 To .SHAPES.ListCount - 1
            If .SHAPES.Selected(n) Then
                setCel = .SHAPES.List(n)
                Set setCel = setCel.Offset(1, 0)
            End If
        Next n
    End With
    
End Sub

Private Sub STEEL_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ToggleSelectAll Me.STEEL
    
    'Clear list
    Settings.Range("Settings.SteelList").ClearContents
    
    Dim setCel As Range
    Set setCel = Settings.Cells(Settings.Range("Settings.SteelList").row, Settings.Range("Settings.SteelList").Column)
    
    'Store settings
    With Me
       For n = 1 To .STEEL.ListCount - 1
            If .STEEL.Selected(n) Then
                setCel = .STEEL.List(n)
                Set setCel = setCel.Offset(1, 0)
            End If
        Next n
    End With
    
End Sub

Private Sub TYPES_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ToggleSelectAll Me.TYPES
    
    'Clear list
    Settings.Range("Settings.TypesList").ClearContents
    
    Dim setCel As Range
    Set setCel = Settings.Cells(Settings.Range("Settings.TypesList").row, Settings.Range("Settings.TypesList").Column)
    
    'Store settings
    With Me
       For n = 1 To .TYPES.ListCount - 1
            If .TYPES.Selected(n) Then
                setCel = .TYPES.List(n)
                Set setCel = setCel.Offset(1, 0)
            End If
        Next n
    End With
    
End Sub

Private Sub UserForm_Initialize()

    If HomePage.Visible = True Then
        With Me
            .StartUpPosition = 0
            .Top = HomePage.Top
            .Left = HomePage.Left + HomePage.Width
        End With
    Else
        With Me
            .StartUpPosition = 0
            .Top = Application.Top + Application.Height / 2 - .Height / 2
            .Left = Application.Left + Application.Width / 2 - Width / 2
        End With
    End If
    
    Dim rng As Range

    ' Clear any existing items & Repopulate
    ' Will select items if previously stored in Settings
    With Me.Galv
        .Clear
        .AddItem "Select All"
        For Each rng In Settings.Range("Settings.Galv")
            If rng = "" Then Exit For
            .AddItem rng.Value
        Next rng
        
        For n = 0 To .ListCount - 1
             Set rng = Settings.Range("Settings.GalvList").Find(.List(n), LookIn:=xlValues, LookAt:=xlWhole)
             If Not rng Is Nothing Then
                .Selected(n) = True
             End If
         Next n
         
    End With
    
    With Me.STEEL
        .Clear
        .AddItem "Select All"
        For Each rng In Settings.Range("Settings.Steel")
            If rng = "" Then Exit For
            .AddItem rng.Value
        Next rng
        
        For n = 0 To .ListCount - 1
             Set rng = Settings.Range("Settings.SteelList").Find(.List(n), LookIn:=xlValues, LookAt:=xlWhole)
             If Not rng Is Nothing Then
                .Selected(n) = True
             End If
         Next n
         
    End With
    
    With Me.SCOUR
        .Clear
        .AddItem "Select All"
        For i = 1 To SoilZones.Range("scourZonesCt")
            .AddItem "S" & i
        Next i
        
        For n = 0 To .ListCount - 1
             Set rng = Settings.Range("Settings.ScourList").Find(.List(n), LookIn:=xlValues, LookAt:=xlWhole)
             If Not rng Is Nothing Then
                .Selected(n) = True
             End If
         Next n
         
    End With
    
    With Me.GEO
        .Clear
        .AddItem "Select All"
        For i = 1 To SoilZones.Range("soilZonesCt")
            .AddItem "G" & i
        Next i
        
        For n = 0 To .ListCount - 1
             Set rng = Settings.Range("Settings.GeoList").Find(.List(n), LookIn:=xlValues, LookAt:=xlWhole)
             If Not rng Is Nothing Then
                .Selected(n) = True
             End If
         Next n
         
    End With
    
    With Me.SHAPES
        .Clear
        .AddItem "Select All"
        For Each rng In Settings.Range("Settings.Shapes")
            If rng = "" Then Exit For
            .AddItem rng.Value
        Next rng
        
        For n = 0 To .ListCount - 1
             Set rng = Settings.Range("Settings.ShapesList").Find(.List(n), LookIn:=xlValues, LookAt:=xlWhole)
             If Not rng Is Nothing Then
                .Selected(n) = True
             End If
         Next n
         
    End With
    
    With Me.TYPES
        .Clear
        .AddItem "Select All"
        For Each rng In TOPLs.Range("TOPL.data").Columns(1).Cells
            If rng = "" Then Exit For
            .AddItem rng.Value
        Next rng
        
        For n = 0 To .ListCount - 1
             Set rng = Settings.Range("Settings.TypesList").Find(.List(n), LookIn:=xlValues, LookAt:=xlWhole)
             If Not rng Is Nothing Then
                .Selected(n) = True
             End If
         Next n
         
    End With
    
    minEmbed.Value = Settings.Range("Settings.minEmbed")
    maxEmbed.Value = Settings.Range("Settings.maxEmbed")
    intEmbed.Value = Settings.Range("Settings.intEmbed")
    
End Sub

Private Sub UserForm_Layout()
    
    If HomePage.Visible = True Then
        With HomePage
            .StartUpPosition = 0
            .Top = BatchAnalysis.Top
            .Left = BatchAnalysis.Left - HomePage.Width
        End With
    End If
    
'    If batchSummary.Visible = True Then
'        With batchSummary
'            .StartUpPosition = 0
'            .Top = BatchAnalysis.Top
'            .Left = BatchAnalysis.Left + BatchAnalysis.Width
'        End With
'    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        If HomePage.Visible = True Then
            HomePage.analysisMode = False
        End If
    End If
    
End Sub

Private Sub weakButton_Click()
    Settings.Range("Settings.axis") = "Weak"
End Sub

Private Sub strongButton_Click()
    Settings.Range("Settings.axis") = "Strong"
End Sub
