Attribute VB_Name = "Module1"
Option Explicit
   
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
        ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
        (ByVal hwnd As Long, ByVal lpString As String, _
        ByVal cch As Long) As Long
Declare Function GetTopWindow Lib "user32" _
        (ByVal hwnd As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hwnd As Long, ByVal lpClassName As String, _
        ByVal nMaxCount As Long) As Long



Public Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type
Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

