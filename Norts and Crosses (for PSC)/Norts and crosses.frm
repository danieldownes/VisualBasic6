VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIRTUAL NORTS AND CROSSES V3"
   ClientHeight    =   4350
   ClientLeft      =   4740
   ClientTop       =   2010
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4350
   ScaleWidth      =   4605
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   255
      Left            =   3000
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
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CA 
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton AC 
      Height          =   615
      Left            =   3120
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton AA 
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton BA 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton BC 
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CB 
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton BB 
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton AB 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Height          =   3015
      Left            =   120
      Top             =   960
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1680
      X2              =   1680
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3000
      X2              =   3000
      Y1              =   960
      Y2              =   3960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Thank you for trying this code. If you have any problems or queries please
'  contact me:
'
'     exd_founder@hotmail.com      (you may also add me to your MSN Messenger contacts)
'
'  Jump to my site, to find other software created by myself;
'
'       http://www.Ex-D.net
'
'   Daniel Downes(UK)  -  Ex-D Software Development(TM)
'
' This is one of my first Visual Basic programs, not too sure why I am posting it to PSC, mybe someone
'  can finish it off.
'
' NOTE: Do not use this code for anything without my permission.
'        I'll probably let you use it, but you must let me know how it is being used.

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

Private Sub cmdInfo_Click()
  MsgBox ("Lead Programmer: Gary Strgy.      Other Programming: Daniel Vincent")
  
End Sub

Private Sub cmdLoad_Click()
  frmLoad.Show 1
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


