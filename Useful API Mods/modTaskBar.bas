Attribute VB_Name = "modTaskBar"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd _
As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal _
wFlags As Long) As Long

Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, ByVal _
lpWindowName As String) As Long


Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40



Sub modHideTaskBar()
    Dim Thwnd As Long
    Thwnd = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Sub modShowTaskBar()
    Dim Thwnd As Long
    Thwnd = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub
