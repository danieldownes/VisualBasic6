Attribute VB_Name = "blockKeys"

   Public Declare Function SystemParametersInfo Lib "user32" _
   Alias "SystemParametersInfoA" (ByVal uAction As Long, _
   ByVal uParam As Long, lpvParam As Any, _
   ByVal fuWinIni As Long) As Long

   Public Const SPI_SCREENSAVERRUNNING = 97
