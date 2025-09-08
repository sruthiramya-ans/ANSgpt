VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BatchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Dashboard.Range("Pile.Type") = ActiveSheet.Cells(Target.row, 1)
    
    Dashboard.Range("Pile.Shape") = ActiveSheet.Cells(Target.row, 4)
    Dashboard.Range("Pile.Galv") = ActiveSheet.Cells(Target.row, 5)
    Dashboard.Range("Pile.Reveal") = ActiveSheet.Cells(Target.row, 6)
    Dashboard.Range("Pile.Embed") = ActiveSheet.Cells(Target.row, 7)
    Dashboard.Range("Soil.Zone") = ActiveSheet.Cells(Target.row, 8)
    Dashboard.Range("Scour.Zone") = ActiveSheet.Cells(Target.row, 9)
    
    Dashboard.Range("Lpile.Name") = ActiveSheet.Cells(Target.row, 24)

    Dashboard.Activate
    Dashboard.Range("Lpile.Name").Activate
    
End Sub

