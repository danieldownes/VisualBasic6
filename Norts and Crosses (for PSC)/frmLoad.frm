VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load - Norts and Crosses"
   ClientHeight    =   3960
   ClientLeft      =   3525
   ClientTop       =   4575
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3960
   ScaleWidth      =   3615
   Begin VB.FileListBox File1 
      Height          =   3210
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   840
      TabIndex        =   1
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

Private Sub cmdCancel_Click()
  frmLoad.Hide
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show 1
End Sub


Private Sub cmdOk_Click()
  MsgBox ("Sorry, still in construction!!!")
End Sub


