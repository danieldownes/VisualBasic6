Attribute VB_Name = "InterNetConnectionCheck"
Option Explicit
Public Declare Function InternetGetConnectedState _
Lib "wininet.dll" (ByRef lpSFlags As Long, _
ByVal dwReserved As Long) As Long

Public Const INTERNET_CONNECTION_LAN As Long = &H2
Public Const INTERNET_CONNECTION_MODEM As Long = &H1



Public Function Online() As Boolean
    'If you are online it will return True, otherwise False
    Online = InternetGetConnectedState(0&, 0&)
End Function

Public Function ViaLAN() As Boolean

    Dim SFlags As Long
    'return the flags associated with the connection
    Call InternetGetConnectedState(SFlags, 0&)
    
    'True if the Sflags has a LAN connection
    ViaLAN = SFlags And INTERNET_CONNECTION_LAN

End Function
Public Function ViaModem() As Boolean

    Dim SFlags As Long
    'return the flags associated with the connection
    Call InternetGetConnectedState(SFlags, 0&)
    
    'True if the Sflags has a modem connection
    ViaModem = SFlags And INTERNET_CONNECTION_MODEM

End Function

