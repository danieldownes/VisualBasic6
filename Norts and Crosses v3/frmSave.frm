VERSION 4.00
Begin VB.Form frmSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save - Norts and Crosses"
   ClientHeight    =   960
   ClientLeft      =   3525
   ClientTop       =   3255
   ClientWidth     =   3615
   Height          =   1365
   Left            =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3615
   Top             =   2910
   Width           =   3735
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Type a name for this session here."
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub


Private Sub cmdCancel_Click()
  frmSave.Hide
End Sub


Private Sub cmdHelp_Click()
  frmHelp.Show 1
End Sub


Private Sub cmdOk_Click()
  MsgBox ("Sorry, still in construction!!!")
End Sub


