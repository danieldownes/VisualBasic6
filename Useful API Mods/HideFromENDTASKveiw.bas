Attribute VB_Name = "modHideFromENDTASKveiw"
'To remove your program from the Ctrl+Alt+Delete list, call the MakeMeService procedure
'To restore your application to the Ctrl+Alt+Delete list, call the UnMakeMeService procedure

Public Declare Function GetCurrentProcessId _
Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess _
Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess _
Lib "kernel32" (ByVal dwProcessID As Long, _
ByVal dwType As Long) As Long

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0




Public Sub modMakeMeService()
    Dim pid As Long
    Dim reserv As Long
    
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub


Public Sub modUnMakeMeService()
    Dim pid As Long
    Dim reserv As Long
    
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, _
    RSP_UNREGISTER_SERVICE)
End Sub


