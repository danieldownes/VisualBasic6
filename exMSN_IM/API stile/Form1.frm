VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   13185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   13185
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   12855
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   12840
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AddChildWindows(ByVal hwndParent As Long, ByVal Level As Long)
      Dim WT As String, CN As String, Length As Long, hwnd As Long
      
      Dim ttemp As TEXTMETRIC
      Dim tt As Long
      
        If Level = 0 Then
          hwnd = hwndParent
        Else
          hwnd = GetWindow(hwndParent, GW_CHILD)
        End If
        Do While hwnd <> 0
          WT = Space(256)
          Length = GetWindowText(hwnd, WT, 255)
          WT = Left$(WT, Length)
          CN = Space(256)
          Length = GetClassName(hwnd, CN, 255)
          CN = Left$(CN, Length)
          Me!Text1 = Me!Text1 & vbCrLf & String(2 * Level, ".") _
                   & WT & " (" & CN & ")"
          AddChildWindows hwnd, Level + 1
          
          If CN = "RichEdit20A" Then

            Length = GetTextMetrics(hwnd, ttemp)
          Me!Text1 = Me!Text1 & vbCrLf & "*" & ttemp.tmFirstChar
           ' MsgBox (temp)
          End If
          
          hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        Loop
      End Sub
   
      Sub Command1_Click()
      Text1.Text = vbNullString
      
      Dim hwnd As Long
        hwnd = GetTopWindow(0)
        If hwnd <> 0 Then
          AddChildWindows hwnd, 0
        End If
      End Sub

