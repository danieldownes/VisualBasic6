VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timT 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5280
      Top             =   2280
   End
   Begin VB.CommandButton cmdHide 
      Appearance      =   0  'Flat
      Caption         =   "?"
      Height          =   1215
      Index           =   5
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdHide 
      Appearance      =   0  'Flat
      Caption         =   "?"
      Height          =   1215
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdHide 
      Appearance      =   0  'Flat
      Caption         =   "?"
      Height          =   1215
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdHide 
      Appearance      =   0  'Flat
      Caption         =   "?"
      Height          =   1215
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdHide 
      Appearance      =   0  'Flat
      Caption         =   "?"
      Height          =   1215
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdHide 
      Appearance      =   0  'Flat
      Caption         =   "?"
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Caption         =   "Correct:"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblMiss 
      Alignment       =   2  'Center
      Caption         =   "Incorrect:"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   1920
      Top             =   1920
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   3480
      Top             =   600
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   1
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   360
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   3600
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblMiss 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   0
      Left            =   480
      Top             =   360
      Width           =   735
   End
   Begin VB.Image imgPicture 
      Height          =   1215
      Index           =   5
      Left            =   3240
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image imgPicture 
      Height          =   1215
      Index           =   4
      Left            =   1680
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image imgPicture 
      Height          =   1215
      Index           =   3
      Left            =   120
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image imgPicture 
      Height          =   1215
      Index           =   2
      Left            =   3240
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image imgPicture 
      Height          =   1215
      Index           =   1
      Left            =   1680
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image imgPicture 
      Height          =   1215
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intPair As Integer
Dim intData(6) As Integer
Dim intScore As Integer
Dim intMiss As Integer
Dim n As Integer

Private Sub cmdHide_Click(Index As Integer)
    
    cmdHide(Index).Visible = False
    If intPair = -1 Then
        intPair = Index + 1
    Else
        timT.Enabled = True
        If intData(intPair) = intData(Index + 1) Then
            intScore = intScore + 1
            lblScore(0).Caption = Str(intScore)
            
            'Move the two hiden buttons out of sight
            For n = 0 To 5
                If cmdHide(n).Visible = False Then
                    cmdHide(n).Left = 100000
                End If
            Next n
        Else
            intMiss = intMiss + 1
            lblMiss(0).Caption = Str(intMiss)
            
        End If
        
        intPair = -1
        
        Call ableButtons(False)
    End If
End Sub

Private Sub Form_Load()
    intData(1) = 1
    intData(2) = 2
    intData(3) = 3
    intData(4) = 2
    intData(5) = 3
    intData(6) = 1
    intPair = -1
    intScore = 0
End Sub

Private Sub timT_Timer()
    For n = 0 To 5
        cmdHide(n).Visible = True
    Next n
    Call ableButtons(True)
    timT.Enabled = False
End Sub

Sub ableButtons(enable As Boolean)
    For n = 0 To 5
        cmdHide(n).Enabled = enable
    Next n
End Sub
