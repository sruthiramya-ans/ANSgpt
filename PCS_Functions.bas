Attribute VB_Name = "PCS_Functions"
Function GetUser() As String
    
    GetUser = Environ("Userprofile")
    
End Function

' Purpose:  Coverts blanks/erros/number string to either 0 or double

Function Nz(v As Variant) As Double
    
    If IsError(v) Then
        Nz = 0
    ElseIf IsNull(v) Or Len(Trim(v & "")) = 0 Then
        Nz = 0
    Else
        Nz = CDbl(v)
    End If
    
End Function

' Purpose:  Given a folder-path (Ex. C:\Users\'User'\Projects)
'           checks if it exists; if not, creates every missing
'           subfolder along the way. Returns the original path.

Public Function EnsureFolderExists(ByVal folderPath As String) As String

    Dim fso As Object
    Dim parts As Variant
    Dim i As Long
    Dim cumulative As String
    
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(folderPath) Then
        EnsureFolderExists = folderPath
        Exit Function
    End If
    
    parts = Split(folderPath, "\")
    
    If Left(parts(0), 2) = "\\" Then
        cumulative = "\\" & parts(1) & "\" & parts(2) & "\"
        i = 3
    Else
        cumulative = parts(0) & "\"
        i = 1
    End If
    
    If Not fso.FolderExists(cumulative) Then
        On Error Resume Next
        fso.CreateFolder cumulative
        On Error GoTo 0
    End If
    
    For i = i To UBound(parts) - 1
        If Len(parts(i)) > 0 Then
            cumulative = cumulative & parts(i) & "\"
            If Not fso.FolderExists(cumulative) Then
                On Error Resume Next
                fso.CreateFolder cumulative
                On Error GoTo 0
            End If
        End If
    Next i
    
    EnsureFolderExists = cumulative
    
End Function

Function LpileReader(filePath As String) As Variant
    
    Dim fileText As String
    Dim lines As Variant
    Dim fnum As Integer
    Dim nLines As Long
    Dim i As Long, j As Long

    ' 1) Read entire file once
    fnum = FreeFile
    
    On Error GoTo filenotfound
    Open filePath For Input As #fnum
      fileText = Input(LOF(fnum), #fnum)
    Close #fnum
    On Error GoTo 0
    
    lines = Split(fileText, vbCrLf)
    nLines = UBound(lines)
    
    ' 2) Find total load cases & Orientation
    Dim numLoads As Long
    Dim strongWeak As String
    For i = 0 To nLines
        If InStr(lines(i), "Cross-sectional Shape") > 0 Then
            strongWeak = lines(i)
        End If
        If InStr(lines(i), "Number of loads specified") > 0 Then
            numLoads = CLng(Val(Mid(lines(i), InStr(lines(i), "=") + 1)))
            Exit For
        End If
    Next i
    If numLoads <= 0 Then
        MsgBox "Failed to find 'Number of loads specified'.", vbExclamation
        Exit Function
    End If
    
    ' 3) Pre-allocate results: (1…LC)×(1…7) 1=LC, 2=PileName, 3=GradeDefl, 4=HeadDefl, 5=MuMax, 6=VuMax, 7=Strong/Weak
    Dim results() As Variant
    ReDim results(1 To numLoads, 1 To 9)
    
    ' 4) Initialize parsing variables
    Dim pileName As String: pileName = ""
    Dim revealVal As Double: revealVal = Range("Pile.Reveal")
    Dim prevDepth As Double, prevDefl As Double, prevMoment As Double, prevShear As Double
    Dim depth As Double, df As Double
    Dim curGradeDefl As Double, curGradeMoment As Double, curGradeShear As Double
    Dim tok As Variant
    Dim loadIndex As Long
    Dim defTableline As Long

    ' 5) Single in-memory pass
    For i = 0 To nLines
        Dim txt As String: txt = Trim(lines(i))
        
        ' a) Grab pile name once
        If pileName = "" And (InStr(txt, "Name of input data file") > 0 Or InStr(txt, "Name of output report file") > 0) Then
            If i + 1 <= nLines Then
                Dim fn As String: fn = Trim(lines(i + 1))
                If InStrRev(fn, ".") > 0 Then fn = Left$(fn, InStrRev(fn, ".") - 1)
                pileName = fn
            End If
            i = i + 1
            GoTo NextIteration
        End If
        
        ' b) Detect P–y table start and interpolate Deflection, Moment, and Shear at Grade
        If InStr(txt, "Pile-head conditions are Shear and Moment") > 0 Then
            defTableline = i + 10
            
            prevDepth = -1: prevDefl = 0: prevMoment = 0: prevShear = 0
            For j = i + 10 To nLines
                Dim rowText As String: rowText = Trim(lines(j))
                If rowText = "" Then
                    ' table ended early ? take last deflection
                    curGradeDefl = prevDefl
                    Exit For
                End If
                tok = Split(Application.Trim(rowText), " ")
                depth = Val(tok(0)): df = Val(tok(1)): moment = Val(tok(2)): shear = Val(tok(3))
                If prevDepth < 0 Then
                    prevDepth = depth: prevDefl = df: prevMoment = moment: prevShear = shear
                ElseIf depth >= revealVal Then
                    ' linear interpolate
                    curGradeDefl = ((prevDefl * (depth - revealVal)) + (df * (revealVal - prevDepth))) / (depth - prevDepth)
                    curGradeMoment = ((prevMoment * (depth - revealVal)) + (moment * (revealVal - prevDepth))) / (depth - prevDepth)
                    curGradeShear = ((prevShear * (depth - revealVal)) + (shear * (revealVal - prevDepth))) / (depth - prevDepth)
                    Exit For
                Else
                    prevDepth = depth: prevDefl = df: prevMoment = moment: prevShear = shear
                End If
            Next j
            
            i = j
            GoTo NextIteration
        End If

        ' c) On each Load Case summary, write one row
        If InStr(txt, "Output Summary for Load Case No.") > 0 Then
            loadIndex = loadIndex + 1
            
            ' --- Check for a ground-level deflection in summary. ---
            Dim groundVal As Double
            Dim k As Long
            For k = i To i + 11
                If InStr(lines(k), "Pile deflection at ground") > 0 Then
                    groundVal = Val( _
                      Mid(lines(k), InStr(lines(k), "=") + 1))
                End If
            Next k
            
            ' 1 = Load Case
            results(loadIndex, 1) = loadIndex
            
            ' 2 = pileName
            results(loadIndex, 2) = pileName
            
            ' 3 = either groundVal or interpolated
            If groundVal > 0 Then
                results(loadIndex, 3) = groundVal
            Else
                results(loadIndex, 3) = curGradeDefl
            End If
            
            ' 4 =  head deflection = 2 lines down
            If i + 2 <= nLines Then
                results(loadIndex, 4) = Val(Mid(lines(i + 2), InStr(lines(i + 2), "=") + 1))
            End If
            
            ' 5 = max bending moment = 4 lines down
            If i + 4 <= nLines Then
                results(loadIndex, 5) = Val(Replace(Mid(lines(i + 4), InStr(lines(i + 4), "=") + 1), "inch-lbs", ""))
            End If
            
            ' 6 = max shear force = 5 lines down
            If i + 5 <= nLines Then
                results(loadIndex, 6) = Abs(Val(Replace(Mid(lines(i + 5), InStr(lines(i + 5), "=") + 1), "lbs", "")))
            End If
            
            ' 7 = strong vs. weak axis - 0=Strong, 1=Weak
            If InStr(strongWeak, "Strong") > 0 Then
                results(loadIndex, 7) = 0
            Else
                results(loadIndex, 7) = 1
            End If
            
            ' 8 = at Grade Moment
            results(loadIndex, 8) = curGradeMoment
            
            ' 9 = at Grade Shear
            results(loadIndex, 9) = curGradeShear
            
            
            i = i + 10
        End If
        
NextIteration:
    Next i
    
    LpileReader = results
    
    Exit Function
    
filenotfound:
    

End Function

Sub CreateDir(strPath As String)

    'Make folder if non-existing

    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
    
End Sub
