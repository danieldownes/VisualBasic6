VERSION 4.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIRTUAL NORTS AND CROSSES V3"
   ClientHeight    =   4350
   ClientLeft      =   3960
   ClientTop       =   1845
   ClientWidth     =   4605
   Height          =   4755
   Left            =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4605
   Top             =   1500
   Width           =   4725
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "Options"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   4080
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "X'es Score"
      Height          =   735
      Left            =   3600
      TabIndex        =   17
      Top             =   120
      Width           =   975
      Begin VB.TextBox lblXScore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "O's Score"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   975
      Begin VB.TextBox lblOScore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton CC 
      Height          =   735
      Left            =   3120
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CA 
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton AC 
      Height          =   735
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton AA 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton BA 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton BC 
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CB 
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton BB 
      Height          =   735
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton AB 
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "It is..."
      Height          =   735
      Left            =   1320
      TabIndex        =   20
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Height          =   2895
      Left            =   120
      Top             =   960
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1560
      X2              =   1560
      Y1              =   960
      Y2              =   3840
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   4440
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Index           =   0
      X1              =   120
      X2              =   4440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3000
      X2              =   3000
      Y1              =   960
      Y2              =   3840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim sngOscore As Single
Dim sngXscore As Single

Function check()
'Across
If AA.Caption = "0" And AB.Caption = "0" And AC.Caption = "0" Then
MsgBox "O's got a three in a row!!!"
sngOscore = sngOscore + 1
Call clear
End If

If BA.Caption = "0" And BB.Caption = "0" And BC.Caption = "0" Then
MsgBox "O's got a three in a row!!!"
sngOscore = sngOscore + 1
Call clear
End If

If CA.Caption = "0" And CB.Caption = "0" And CC.Caption = "0" Then
Print "Well done you got a line"
sngOscore = sngOscore + 1
Call clear
End If

'Down
If AA.Caption = "0" And BA.Caption = "0" And CA.Caption = "0" Then
MsgBox "O's got a three in a row!!!"
sngOscore = sngOscore + 1
Call clear
End If

If AB.Caption = "0" And BB.Caption = "0" And CB.Caption = "0" Then
MsgBox "O's got a three in a row!!!"
sngOscore = sngOscore + 1
Call clear
End If

If AC.Caption = "0" And BC.Caption = "0" And CC.Caption = "0" Then
MsgBox "O's got a three in a row!!!"
sngOscore = sngOscore + 1
Call clear
End If

'Diagnal
If AA.Caption = "0" And BB.Caption = "0" And CC.Caption = "0" Then
MsgBox "O's got a three in a row!!!"
sngOscore = sngOscore + 1
Call clear
End If

If AC.Caption = "0" And BB.Caption = "0" And CA.Caption = "0" Then
MsgBox "O's got a three in a row!!!"
sngOscore = sngOscore + 1
Call clear
End If

'Across
If AA.Caption = "X" And AB.Caption = "X" And AC.Caption = "X" Then
Print "Well done you got a line"
sngOscore = sngOscore + 1
Call clear
End If

If BA.Caption = "X" And BB.Caption = "X" And BC.Caption = "X" Then
Print "Well done you got a line"
sngXscore = sngXscore + 1
Call clear
End If

If CA.Caption = "X" And CB.Caption = "X" And CC.Caption = "X" Then
Print "Well done you got a line"
sngXscore = sngXscore + 1
Call clear
End If

'Down
If AA.Caption = "X" And BA.Caption = "X" And CA.Caption = "X" Then
Print "Well done you got a line"
sngXscore = sngXscore + 1
Call clear
End If

If AB.Caption = "X" And BB.Caption = "X" And CB.Caption = "X" Then
Print "Well done you got a line"
sngXscore = sngXscore + 1
Call clear
End If

If AC.Caption = "X" And BC.Caption = "X" And CC.Caption = "X" Then
Print "Well done you got a line"
sngXscore = sngXscore + 1
Call clear
End If

'Diagnal
If AA.Caption = "X" And BB.Caption = "X" And CC.Caption = "X" Then
Print "Well done you got a line"
sngXscore = sngXscore + 1
Call clear
End If

If AC.Caption = "X" And BB.Caption = "X" And CA.Caption = "X" Then
Print "Well done you got a line"
sngXscore = sngXscore + 1
Call clear
End If

lblOScore.Text = Str$(sngOscore)
lblXScore.Text = Str$(sngXscore)


End Function
Function clear()
AA.Caption = "": BA.Caption = "": CA.Caption = ""
AB.Caption = "": BB.Caption = "": CB.Caption = ""
AC.Caption = "": BC.Caption = "": CC.Caption = ""

End Function


Private Sub Command1_Click()

End Sub

Private Sub Command15_Click()
c
End Sub


Private Sub AA_Click()

If AA.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    AA.Caption = "0"
  Else
    AA.Caption = "X"
  End If
Else
Beep
End If

Call check

End Sub


Private Sub AB_Click()
If AB.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    AB.Caption = "0"
  Else
    AB.Caption = "X"
  End If
Else
Beep
End If
Call check
End Sub

Private Sub AC_Click()
If AC.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    AC.Caption = "0"
  Else
    AC.Caption = "X"
  End If
Else
Beep
End If
Call check
End Sub


Private Sub BA_Click()
If BA.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    BA.Caption = "0"
  Else
    BA.Caption = "X"
  End If
Else
Beep
End If
Call check
End Sub


Private Sub BB_Click()
If BB.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    BB.Caption = "0"
  Else
    BB.Caption = "X"
  End If
Else
Beep
End If
Call check
End Sub


Private Sub BC_Click()
If BC.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    BC.Caption = "0"
  Else
    BC.Caption = "X"
  End If
Else
Beep
End If
Call check
End Sub


Private Sub CA_Click()
If CA.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    CA.Caption = "0"
  Else
    CA.Caption = "X"
  End If
Else
Beep
End If
Call check
End Sub


Private Sub CB_Click()
If CB.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    CB.Caption = "0"
  Else
    CB.Caption = "X"
  End If
Else
Beep
End If
Call check
End Sub


Private Sub CC_Click()
If CC.Caption = "" Then
  If Text1.Text = "Norts turn" Then Text1.Text = "Crosses turn" Else Text1.Text = "Norts turn"
  
  If Text1.Text <> "Norts turn" Then
    CC.Caption = "0"
  Else
    CC.Caption = "X"
  End If
Else
Beep
End If
Call check
End Sub


Private Sub Command2_Click()
  MsgBox ("By Daniel Vincent and Gary Stryg")
End Sub

Private Sub cmdEnd_Click()
  End
End Sub

Private Sub cmdHelp_Click()
frmHelp.Show 1
End Sub



Private Sub cmdLoad_Click()
  frmLoad.Show 1
End Sub

Private Sub cmdNew_Click()
Call clear

End Sub

Private Sub cmdSave_Click()
  frmSave.Show 1
End Sub

Private Sub Form_Load()
Dim turns As Single
End Sub



Private Sub Picture1_Click()
Print "0"
End Sub


Private Sub Text3_Change()

End Sub


Private Sub Text2_Change(Index As Integer)

End Sub


