VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Caption = Str(NewValue(Val(Text1.Text), Val(Text2.Text), False))
End Sub


Function NewValue(angle As Double, speed As Single, bX As Boolean) As Double
    Select Case bX
        Case True
            NewValue = Sin(angle * speed)
            NewValue = NewValue * speed
        Case False
            NewValue = Cos(angle) * speed
    End Select
    
    'If angle > 90 Then NewValue = -NewValue
End Function
