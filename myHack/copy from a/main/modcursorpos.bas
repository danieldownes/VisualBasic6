Attribute VB_Name = "modcursorpos"
Option Explicit

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
