VERSION 4.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load - Norts and Crosses"
   ClientHeight    =   3960
   ClientLeft      =   3525
   ClientTop       =   4575
   ClientWidth     =   3615
   Height          =   4365
   Left            =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3615
   Top             =   4230
   Width           =   3735
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   3720
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   3375
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   3720
      Width           =   735
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  frmLoad.Hide
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show 1
End Sub


Private Sub cmdOk_Click()
  MsgBox ("Sorry, still in construction!!!")
End Sub


