VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save - Norts and Crosses"
   ClientHeight    =   960
   ClientLeft      =   3525
   ClientTop       =   3255
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   960
   ScaleWidth      =   3615
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
  frmSave.Hide
End Sub


Private Sub cmdHelp_Click()
  frmHelp.Show 1
End Sub


Private Sub cmdOk_Click()
  MsgBox ("Sorry, still in construction!!!")
End Sub


