Attribute VB_Name = "Functions"
#If VBA7 Then
    Public Declare PtrSafe Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If
Option Base 1
Option Explicit


Public Function IsExeRunning(sExeName As String, Optional sComputer As String = ".") As Boolean
On Error GoTo Error_Handler
Dim objProcesses As Object

Set objProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2").ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & sExeName & "'")
If objProcesses.Count <> 0 Then IsExeRunning = True

Error_Handler_Exit:
On Error Resume Next
Set objProcesses = Nothing
Exit Function

Error_Handler:
MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
        "Error Number: IsExeRunning" & vbCrLf & _
        "Error Description: " & Err.Description, _
        vbCritical, "An Error has Occured!"
Resume Error_Handler_Exit
End Function


Public Function OpenFile()

    Dim b As Integer
    Dim fileSpec As String

    fileSpec = ThisWorkbook.Path & "\LPile\" & [Project.Name].Value2 & " - ANSgpt.lp11d"
    fileName = Dir(fileSpec)
    
    If fileName <> "" Then
        b = ShellExecute(0, "Open", fileSpec, "", "", 3)
    Else
        MsgBox "No files were found that match: " & vbCrLf & vbCrLf & fileSpec & vbCrLf & vbCrLf _
                & vbCrLf & "Please create new LPile input file to continue.", , "LPile"
        Exit Function
    End If

End Function

'Removes empty elements from an array
Function RemoveEmpties(arr() As Variant) As Variant

    Dim aTEMP() As Variant, aFinal() As Variant
    Dim b As Integer, i  As Integer
    
    aTEMP() = arr()
    Erase aFinal()
  
    i = 1
    For b = LBound(aTEMP) To UBound(aTEMP)
'        If aTEMP(b) <> vbNullString Then
        If IsEmpty(aTEMP(b)) = False Then
            ReDim Preserve aFinal(i)
            aFinal(i) = aTEMP(b)
            i = i + 1
        End If
    Next b
    
    RemoveEmpties = aFinal()
  
End Function

'Remove duplicate values from an array
Function RemoveDuplicates(arr() As Variant) As Variant

    Dim aTEMP() As Variant, aFinal() As Variant
    Dim b As Integer, i  As Integer

    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    
    aTEMP() = arr()
    Erase aFinal()
    
    With d
        For i = LBound(aTEMP) To UBound(aTEMP)
            If IsMissing(aTEMP(i)) = False Then
                .item(aTEMP(i)) = 1
            End If
        Next
        aTEMP() = .Keys
    End With
    
    ReDim aFinal(1 To d.Count)
    For b = 1 To d.Count
        aFinal(b) = aTEMP(b - 1)
    Next b
    
    RemoveDuplicates = aFinal()
    
End Function

