VERSION 5.00
Begin VB.Form frmMSN_IM 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2400
      Width           =   3735
   End
End
Attribute VB_Name = "frmMSN_IM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mesMessengerApp As MessengerApp
Dim mesMsgrObject As MsgrObject
Dim mesIMessengerApp As IMessengerApp

Dim mesIMsgrService As IMsgrService
Dim mesIMsgrServices As IMsgrServices

Dim mesIMessengerIMWindow As IMessengerIMWindow
Dim mesIMessengerIMWindows As IMessengerIMWindows

Dim mesIMsgrIMSession As IMsgrIMSession
Dim mesIMsgrIMSessions As IMsgrIMSessions

Dim mesIMsgrUser As IMsgrUser


Private Sub Form_Load()
    'mesIMessengerApp.LaunchLogonUI
    MsgBox mesIMsgrService.ServiceName
    'mesIMsgrService.LogonName = "Ex-D Founder"
    
   ' mesMsgrObject.Logon "daniel_vincent99@hotmail.com", "darkspace", mesIMsgrService
    
End Sub
