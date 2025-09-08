VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} batchSummary 
   Caption         =   "Results Summary for Lightest Passing Pile"
   ClientHeight    =   5436
   ClientLeft      =   24
   ClientTop       =   120
   ClientWidth     =   12408
   OleObjectBlob   =   "batchSummary.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "batchSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Dim TOPLarray As Variant, geoArray As Variant, scourArray As Variant
    Dim item As Variant, gZone As Variant, sZone As Variant
    Dim cel As Range, lbl As Object, shtLog As Range

    'Get all solved TOPL list
    TOPLarray = WorksheetFunction.Unique(BatchResults.Range("Batch.Data").Columns(1).Cells)
    geoArray = WorksheetFunction.Unique(BatchResults.Range("Batch.Data").Columns(8).Cells)
    scourArray = WorksheetFunction.Unique(BatchResults.Range("Batch.Data").Columns(9).Cells)
    topstart = 48
    n = 1
    
    PileMenu.Range("Menu.Full").ClearContents
    Set shtLog = PileMenu.Range("Menu.Full").Cells(1, 1)
        
    For Each gZone In geoArray
                Debug.Print gZone
        If gZone = "" Then Exit For
    
        For Each sZone In scourArray
                Debug.Print sZone
            If sZone = "" Then Exit For
            
            For Each item In TOPLarray
                Debug.Print item
                If item = "" Then Exit For
                pileWt = 999

                'Find lightest section/embed combo for TOPL
                For Each cel In BatchResults.Range("Batch.Data").Columns(1).Cells
                    If cel = "" Then Exit For
                    If cel.Offset(0, 1) <> "Pass" Then GoTo skipit
                    If cel = item And gZone = BatchResults.Cells(cel.row, 8) And sZone = BatchResults.Cells(cel.row, 9) And _
                            BatchResults.Cells(cel.row, 17) < pileWt Then
                        pileWt = BatchResults.Cells(cel.row, 17)
                        pileRow = cel.row
                    End If
skipit:
                Next cel
                                
                'Make Labels
                '***********
                On Error GoTo dataError
                leftstart = 12
                
                'Type
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = BatchResults.Cells(pileRow, 1)
                    .Left = leftstart
                    .Width = 140
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog = .Caption
                End With
                
                'Shape
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = BatchResults.Cells(pileRow, 4)
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 1) = .Caption
                End With
                
                'Galv
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = BatchResults.Cells(pileRow, 5)
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 2) = .Caption
                End With
                
                'Reveal
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = BatchResults.Cells(pileRow, 6)
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 3) = .Caption
                End With
                
                'Embed
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = BatchResults.Cells(pileRow, 7)
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 4) = .Caption
                End With
                
                'Head defl
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = Format(WorksheetFunction.Max(BatchResults.Cells(pileRow, 10), BatchResults.Cells(pileRow, 12)), "0.0000")
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 5) = .Caption
                End With
                
                'Grade defl
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = Format(WorksheetFunction.Max(BatchResults.Cells(pileRow, 11), BatchResults.Cells(pileRow, 13)), "0.0000")
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 6) = .Caption
                End With
                
                'Soil usage
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = Format(BatchResults.Cells(pileRow, 14), "0.00%")
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 7) = .Caption
                End With
                
                'Steel usage
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = Format(WorksheetFunction.Max(BatchResults.Cells(pileRow, 15), BatchResults.Cells(pileRow, 16)), "0.00%")
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 8) = .Caption
                End With
                
                'Geo zone
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = BatchResults.Cells(pileRow, 8)
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 9) = .Caption
                End With
                
                'Scour zone
                Set lbl = Me.Controls.Add("Forms.Label.1")
                With lbl
                    .Caption = BatchResults.Cells(pileRow, 9)
                    .Left = leftstart
                    .Width = 40
                    .Top = topstart
                    .BackColor = RGB(244, 244, 242)
                    leftstart = leftstart + .Width + 6
                    shtLog.Offset(0, 10) = .Caption
                End With
                
                topstart = topstart + 18
                
                'Strong axis - At grade moment
                shtLog.Offset(0, 11) = BatchResults.Cells(pileRow, 26)

                
                'Strong axis - At grade shear
                shtLog.Offset(0, 12) = BatchResults.Cells(pileRow, 27)
                
                'Strong axis - At max moment
                shtLog.Offset(0, 13) = BatchResults.Cells(pileRow, 28)
                
                'Weak axis - At grade moment
                shtLog.Offset(0, 14) = BatchResults.Cells(pileRow, 29)

                
                'Weak axis - At grade shear
                shtLog.Offset(0, 15) = BatchResults.Cells(pileRow, 30)

                'Weak axis - At max moment
                shtLog.Offset(0, 16) = BatchResults.Cells(pileRow, 31)

                Set shtLog = shtLog.Offset(1, 0)
                
            Next item
            
            'Scour break
            Set lbl = Me.Controls.Add("Forms.Label.1")
            With lbl
                .Left = 10
                .Width = 602
                .Height = 1
                .Top = topstart - 3
                .BorderStyle = 1
            End With
                    
        Next sZone
            
        'Geo break
        Set lbl = Me.Controls.Add("Forms.Label.1")
        With lbl
            .Left = 10
            .Width = 602
            .Height = 1
            .Top = topstart - 3
            .BorderStyle = 1
        End With
                    
    Next gZone
    
     Me.Height = 42 + 18 * (UBound(TOPLarray)) * (UBound(geoArray) - 1) * (UBound(scourArray) - 1) + 20
     
     Exit Sub
     
dataError:
     
    MsgBox "It's likely that a solution could not be found for one or more Pile Types. Consider expanding embedment range and/or pile shapes bfore running a new batch.", vbCritical, "Error"
    
'    If BatchAnalysis.Visible = True Then
'        With batchSummary
'            .StartUpPosition = 0
'            .Top = BatchAnalysis.Top
'            .Left = BatchAnalysis.Left + BatchAnalysis.Width
'        End With
'    End If
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Unload Me
    
End Sub
