Attribute VB_Name = "modShutdown"
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved _
As Long) As Long

Sub modShutDownComp()
    ExitWindowsEx EWX_FORCE Or EWX_SHUTDOWN, 0
End Sub

Sub modRestartComp()
    ExitWindowsEx EWX_FORCE Or EWX_REBOOT, 0
End Sub

