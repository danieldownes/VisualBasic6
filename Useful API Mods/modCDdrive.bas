Attribute VB_Name = "modCDdrive"
Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, ByVal _
lpstrReturnString As String, ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long


Sub modCDdriveOpen()
    retvalue = mciSendString("set CDAudio door open", _
    returnstring, 127, 0)
End Sub

Sub modCDdriveClose()
    retvalue = mciSendString("set CDAudio door closed", _
    returnstring, 127, 0)
End Sub

