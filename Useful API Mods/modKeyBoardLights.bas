Attribute VB_Name = "modKeyBoardLights"
'Const VK_CAPITAL = &H14
'
'Type KeyboardBytes
'kbByte(0 To 255) As Byte
'End Type
'
'kbArray As KeyboardBytes
'
'Declare Function GetKeyState Lib "user32" _
'(ByVal nVirtKey As Long) As Long
'
'Declare Function GetKeyboardState Lib "user32" _
'(kbArray As KeyboardBytes) As Long
'
'Declare Function SetKeyboardState Lib "user32" _
'(kbArray As KeyboardBytes) As Long
'
'
'
'
'
'
'
'
'
'
'
'
'
'Sub modTurnAllOff()
'
'  GetKeyboardState kbArray
'  kbArray.kbByte(VK_CAPITAL) = _
'  IIf(kbArray.kbByte(VK_CAPITAL) = 1, 0, 1)
'  SetKeyboardState kbArray
'
'End Sub
