Attribute VB_Name = "modScreenSize"
    Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
   End Type

   Public Const SPI_GETWORKAREA = 48

   Declare Function SystemParametersInfo Lib "user32" _
     Alias "SystemParametersInfoA" (ByVal uAction As Long, _
     ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) _
     As Long




Function moddScreenSize() As POINTAPI
  Dim lRet As Long
  Dim apiRECT As RECT

  lRet = SystemParametersInfo(SPI_GETWORKAREA, vbNull, apiRECT, 0)

  If lRet Then

    moddScreenSize.x = apiRECT.Right - apiRECT.Left
    moddScreenSize.y = apiRECT.Bottom - apiRECT.Top
  End If
End Function
