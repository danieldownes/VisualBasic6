VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy From M3u (by Yovas)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "O&verWrite files"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtLog 
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   1800
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog cD 
      Left            =   1200
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtDestino 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox txtOrigen 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Copy Files"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Log:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Winamp PlayList (*.M3U):"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Destination folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Dim bOverWrite As Boolean

Private Sub Check1_Click()
    bOverWrite = Check1.Value Xor 1
End Sub

Private Sub Command1_Click()
Dim sBuffer As String
Dim Destination As String
Dim Source As String
Dim Reply As Long
Dim TotalFiles As Integer
Dim sRenameTo As String
    
    If InStr(1, txtOrigen.Text, "/") > 0 Then
        txtOrigen.Text = Slashes(txtOrigen.Text)
    End If
    
    If InStr(1, txtDestino.Text, "/") > 0 Then
        txtDestino.Text = Slashes(txtDestino.Text)
    End If
    
    If Not Dir(txtDestino.Text, vbDirectory) > "" Then
        MkDir txtDestino.Text
    End If
    
    If InStr(1, txtOrigen.Text, ".m3u", vbTextCompare) < 1 Then
        txtOrigen.Text = txtOrigen.Text & ".m3u"
    End If
    
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Check1.Enabled = False
    
    txtLog.Text = vbCrLf & "!!!-PRESS [ESC] TO CANCEL!!!!!"
    
    Open txtOrigen.Text For Input As #1
        Do While Not EOF(1)
            
            Line Input #1, sBuffer
                                    
            If Left(sBuffer, 1) = "#" Then
                If Left(sBuffer, 7) = "#EXTINF" Then
                Debug.Print InStrRev(sBuffer, ",")
                    sRenameTo = Mid(sBuffer, InStr(sBuffer, ",") + 1, Len(sBuffer) - InStr(sBuffer, ","))
                End If
            Else
            
                Select Case Left(sBuffer, 1)
                Case "\":
                    Destination = txtDestino.Text & "\" & GetPath(sBuffer, 2)
                    Source = GetPath(sBuffer, 1) & GetPath(sBuffer, 2)
                Case Else:
                    If GetPath(sBuffer, 1) = "" Then
                        Destination = txtDestino.Text & "\" & sBuffer
                        Source = GetPath(txtOrigen.Text, 1) & sBuffer
                    Else
                        If InStr(1, GetPath(sBuffer, 1), ":", vbTextCompare) = 0 Then
                         '   Destination = txtDestino.Text & "\" & GetPath(sBuffer, 2)
                            Source = GetPath(txtOrigen.Text, 1) & GetPath(sBuffer, 1) & GetPath(sBuffer, 2)
                        Else
                         '   Destination = txtDestino.Text & "\" & GetPath(sBuffer, 2)
                            Source = sBuffer
                        End If
                    End If
                End Select
                
                ' Copy File
               ' Debug.Print Destination
                                
                Reply = CopyFile(Source, Destination & "\" & sRenameTo, bOverWrite)
                
                Select Case Reply
                Case Is = 0:
                    txtLog.Text = vbCrLf & ":(-Error copying (" & Source & ") to (" & Destination & ")" & txtLog.Text
                Case Is = 1:
                    txtLog.Text = vbCrLf & ":)-File (" & Source & ") copied to (" & Destination & ")" & txtLog.Text
                    TotalFiles = TotalFiles + 1
                End Select
               
               
            End If
            
            DoEvents
            
            If GetAsyncKeyState(vbKeyEscape) < 0 Then
                txtLog.Text = vbCrLf & ":0-!!!!OPERATION CANCELLED BY USER!!!!" & txtLog.Text
                GoTo 10
            End If
            
        Loop
10:
    Close #1
    
    txtLog.Text = vbCrLf & ":D-Total number of files copied: " & TotalFiles & txtLog.Text
    
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Check1.Enabled = True
End Sub

Private Function GetPath(sInput As String, Folder1_File2 As Byte) As String
Dim i As Integer

    If Folder1_File2 = 1 Then
        For i = 1 To Len(sInput)
            If Mid(StrReverse(sInput), i, 1) = "\" Then
                GetPath = Mid(sInput, 1, Len(sInput) - i) & "\"
                Exit For
            End If
        Next i
    Else
        For i = 1 To Len(sInput)
            If Mid(StrReverse(sInput), i, 1) = "\" Then
                GetPath = Right(sInput, i - 1)
                Exit For
            End If
        Next i
    End If

End Function

Private Function Slashes(tmpInput As String) As String
Dim i As Integer

    For i = 1 To Len(tmpInput)
        If Mid(tmpInput, i, 1) = "/" Or (i = Len(tmpInput) And Mid(tmpInput, i, 1) = "\") Then
            If i = Len(tmpInput) Then Exit For
            Slashes = Slashes & "\"
        Else
            Slashes = Slashes & Mid(tmpInput, i, 1)
        End If
    Next i

End Function

Private Sub Command2_Click()

    On Error GoTo 20
    
    cD.DialogTitle = "Open Winamp List *.M3U..."
    cD.FileName = ""
    cD.Filter = "List [*.m3u]|*.m3u|"
    cD.CancelError = True
    cD.ShowOpen

    If Dir(cD.FileName) <> "" Then
        txtOrigen.Text = cD.FileName
    End If

20:
End Sub

Private Sub Command3_Click()
    Form1.Enabled = False
    Form2.Show
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub Form_Load()
    Check1.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
