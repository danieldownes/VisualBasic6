VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStrenth 
      Height          =   285
      Left            =   7200
      TabIndex        =   6
      Text            =   "0.5"
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdStrenth 
      Caption         =   "Strenth"
      Height          =   255
      Left            =   7200
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Text            =   "0"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Text            =   "0"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Text            =   "0"
      Top             =   5160
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8640
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Shape shpVec 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   960
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intMax As Integer


Private Sub cmdAdd_Click()
    Dim vecIn As D3DVECTOR
    
    vecIn.x = Val(Text1(0).Text)
    vecIn.y = Val(Text1(1).Text)
    vecIn.z = Val(Text1(2).Text)
    
    Call AddPathNode(vecIn)
End Sub

Private Sub cmdDraw_Click()
    Dim vecOut As D3DVECTOR
    Dim n As Single
    
    For n = 0 To GetNodeCount - 1 Step 0.1
        vecOut = GetSplinePoint(n, Val(txtStrenth))
        If n <> 0 And n * 10 >= intMax Then Load shpVec(intMax)
        
        With shpVec(n * 10)
            .Visible = True
            .Left = vecOut.x
            .Top = vecOut.y
        End With
        
        If n * 10 > intMax Then intMax = intMax + 1
        
    Next n
End Sub

Private Sub cmdNodes_Click()
    lblNodes.Caption = Str(GetNodeCount)
End Sub

Private Sub cmdPoint_Click()
    Dim vecOut As D3DVECTOR
    vecOut = GetSplinePoint(Val(txtPoint.Text))
    MsgBox (Str(vecOut.x) + " " + Str(vecOut.y))
End Sub

Private Sub Form_Load()
    intMax = 1
End Sub
