Attribute VB_Name = "modSessionTime"
Declare Function GetTickCount& Lib "kernel32" ()


Function modOSSessionTime() As Long
    modOSSessionTime = GetTickCount& / 1000
End Function
