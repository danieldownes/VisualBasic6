VERSION 4.00
Begin VB.Form frmOpts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2730
   ClientLeft      =   3345
   ClientTop       =   1920
   ClientWidth     =   5325
   Height          =   3135
   Left            =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5325
   Top             =   1575
   Width           =   5445
   Begin VB.Frame Frame1 
      Caption         =   "X'es Colour"
      Height          =   1215
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Frame fraColour 
      Caption         =   "O's Colour"
      Height          =   1215
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Frame fraFirst 
      Caption         =   "Who goes first?"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optX 
         Caption         =   "X goes first"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optO 
         Caption         =   "O goes first"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmOpts"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
