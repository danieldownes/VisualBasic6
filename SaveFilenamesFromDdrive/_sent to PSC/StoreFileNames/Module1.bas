Attribute VB_Name = "Fileing"
' Note, this modual was not made by me, see next couple of comments...

'----------------------------------------
'- Name:    Karl E. Peterson
'- Email:
'- Web: www.mvps.org/vb
'- Company:
'----------------------------------------
'- Notes:   Copy or paste files to the clipboard as '       Explorer does
'----------------------------------------

Option Explicit

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_NOCONFIRMATION = &H10  ' No Confirmation
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_SIMPLEPROGRESS = &H100

Public Sub DeleteFile(gFile As String, gAllowUndo As Boolean)
    Dim op As SHFILEOPSTRUCT

    With op
        .wFunc = FO_DELETE
        .pFrom = gFile
        If gAllowUndo = True Then
            .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
        End If
        
    End With
    SHFileOperation op
End Sub
Public Sub CopyFile(gFileSource As String, NewLoc As String)
    Dim op As SHFILEOPSTRUCT
    If gFileSource = "" Or NewLoc = "" Then Exit Sub
    With op
        .wFunc = FO_COPY
        .pTo = NewLoc
        .pFrom = gFileSource
        .fFlags = FOF_SIMPLEPROGRESS + FOF_NOCONFIRMATION
    End With
    SHFileOperation op
    
    Call addLog("* Copyed file: " & gFileSource & " at " & Time & " on the " & Date & " *")
    
End Sub
Public Sub MoveFile(gFileSource As String, NewLoc As String)
    Dim op As SHFILEOPSTRUCT
    If gFileSource = "" Or NewLoc = "" Then Exit Sub
    With op
        .wFunc = FO_MOVE
        .pTo = NewLoc
        .pFrom = gFileSource
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation op
End Sub

Public Sub addLog(strAdd As String)
    Open "C:\windows\system\log.dat" For Append As #1
        Print #1, strAdd
    Close #1
End Sub
