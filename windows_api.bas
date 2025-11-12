#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
        ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
#End If

Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10

Sub WaitMilliseconds(ms As Long)
    DoEvents
    Sleep ms
End Sub

Sub MakeFormTopMost(formCaption As String)
    Dim hWnd As LongPtr
    hWnd = FindWindowA("ThunderDFrame", formCaption)
    If hWnd <> 0 Then
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    Else
        MsgBox "Form window not found. Make sure the caption is correct and the form is visible.", vbExclamation
    End If
End Sub

