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
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tv As New TrueVision8
Dim Scene As New Scene8
Dim InputEng As New InputEngine8
Dim mesh As New Mesh8


Private Sub Command1_Click()
    Me.Refresh
    tv.Init3DWindowedMode Me.hWnd
    tv.SetSearchDirectory App.Path  'to set the texture/object/... directory
    Set mesh = Scene.CreateMeshBuilder("TestMesh")
    
    
    Do
        tv.Clear
        DoEvents
        tv.RenderToScreen
    Loop Until InputEng.IsKeyPressed(TV_KEY_ESCAPE) = True
    Set tv = Nothing
    End
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_Terminate()

End Sub
