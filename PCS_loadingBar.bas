Attribute VB_Name = "PCS_loadingBar"
Sub UpdateProgressBar(n As Long, m As Long, uCaption As String, r As Integer, g As Integer, b As Integer, Optional DisplayText As String)

    On Error GoTo ERR_HANDLE
    
    progressBar.Caption = uCaption
    progressBar.BoxProgress.BackColor = RGB(r, g, b)
    
    If n >= m Then
    
        'progressBar.Hide
        
        Unload progressBar
        
    Else
    
        If progressBar.Visible = False Then progressBar.Show
        progressBar![BoxProgress].Caption = IIf(DisplayText = vbNullString, Round(((n / m) * 10000) / 100) & "%", DisplayText)
        progressBar![BoxProgress].Width = (n / m) * 468
        DoEvents
        
    End If
    
    Exit Sub
    
ERR_HANDLE:
        Err.Clear
        'progressBar.Hide
        
        Unload progressBar
        
End Sub


