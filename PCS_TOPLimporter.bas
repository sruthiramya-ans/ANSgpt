Attribute VB_Name = "PCS_TOPLimporter"
Public Sub importTOPLs()

    Dim analysisBook As Workbook, TOPLbook As Workbook, TOPLsheet As Worksheet, TOPLrange As Range, TOPLdat As Range
    
    Application.ScreenUpdating = False
    
    Set analysisBook = ThisWorkbook
    TOPLpath = TOPLs.Range("TOPL.filepath")
    Set TOPLbook = Workbooks.Open(TOPLpath)
    Set TOPLsheet = TOPLbook.Sheets(DropTOPLs.sheetsListBox.Value)
    analysisBook.Activate
    
    'Find summary table in sheet
    With TOPLsheet
    
        For i = 1 To 100
    
            For j = 1 To 100
    
                If .Cells(j, i) = "Pile Type" Then GoTo tableFound
                
            Next j
    
        Next i
        
        MsgBox "No Pile Type header found in selected sheet.", vbExclamation, "No data found"
        TOPLbook.Close savechanges:=False
        Exit Sub
        
tableFound:
    
        'find first blank row in TOPLs sheet
        Set TOPLrange = TOPLs.Cells(TOPLs.Range("TOPL.data").row, TOPLs.Range("TOPL.data").Column)
        Do Until TOPLrange = ""
        
            Set TOPLrange = TOPLrange.Offset(1, 0)
        
        Loop
    
        'find first Pile Type
        j = j + 1
        Do Until .Cells(j, i) <> ""

            j = j + 1
        
        Loop
        
        'Loop thru Pile Types
        Do Until .Cells(j, i) = ""
        
            TOPLrange = .Cells(j, i) & " (" & DropTOPLs.revealHeightBox.Value & "ft)"
            TOPLrange.Offset(0, 1) = DropTOPLs.revealHeightBox.Value
            TOPLrange.Offset(0, 2) = .Cells(j, i + 1)
            TOPLrange.Offset(0, 3) = .Cells(j, i + 2)
            TOPLrange.Offset(0, 4) = .Cells(j, i + 3)
            TOPLrange.Offset(0, 5) = .Cells(j, i + 4)
            TOPLrange.Offset(0, 6) = .Cells(j, i + 5)
            TOPLrange.Offset(0, 7) = .Cells(j, i + 6)
            TOPLrange.Offset(0, 8) = .Cells(j, i + 7)
        
            Set TOPLrange = TOPLrange.Offset(1, 0)
            j = j + 1
        
        Loop
    
    End With

    TOPLbook.Close savechanges:=False
    
End Sub
