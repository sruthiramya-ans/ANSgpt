Attribute VB_Name = "graphsFromLpile"
Sub PlotDeflectionFromLP12P()
    Dim filePath As String
    Dim ws As Worksheet
    Dim lineText As String
    Dim arr() As String
    Dim lastRow As Long
    Dim fnum As Integer
    
    ' --- Set the file path ---
    filePath = "C:\Users\SruthiRamya\LPile\Paloma\Edge Interior Colum-5.125 ft\Edge Interior Colum-5.125 ft-W6X20-Embed 12ft-0 mil-Soil G1-Scour S1Strong.lp12p"
    
    ' --- Set worksheet ---
    Set ws = ThisWorkbook.Sheets("Sheet9")
    ws.Cells.Clear
    
    ' --- Open the file ---
    fnum = FreeFile
    Open filePath For Input As #fnum
    
    lastRow = 1
    Do While Not EOF(fnum)
        Line Input #fnum, lineText
        ' Skip empty lines
        If Trim(lineText) <> "" Then
            ' Split by comma (change if tab/space)
            arr = Split(lineText, vbTab)
            ' Write Depth in column A, Deflection in column B
            ws.Cells(lastRow, 1).Value = arr(0)
            ws.Cells(lastRow, 2).Value = arr(1)
            lastRow = lastRow + 1
        End If
    Loop
    
    Close #fnum
    
    ' --- Create chart ---
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=300, Width:=500, Top:=50, Height:=300)
    With chartObj.Chart
        .ChartType = xlXYScatterLines
        .SetSourceData ws.Range("A1:B" & lastRow - 1)
        .HasTitle = True
        .ChartTitle.Text = "Pile Deflection vs Depth"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Depth (ft/m)"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Deflection (in/mm)"
        .Axes(xlCategory).ReversePlotOrder = True
    End With
    
    MsgBox "Chart created successfully!"
End Sub

