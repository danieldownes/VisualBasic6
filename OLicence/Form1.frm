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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'function OLPromptLicense(hParent: THandle; lpID: PChar):
'  Integer;
'function OLStoreLicense(hParent: THandle; lpID, lpLicense: PChar):
'  Integer;
'function OLRetrieveKey(hParent: THandle; lpID, lpPassword, lpKey: PChar; nKeySize: Integer):
'  Integer;

Private Sub Form_Load()
   ' MsgBox (OLPromptLicense(Me.hWnd, "Form1"))
    MsgBox (OLStoreLicense(Me.hWnd, "Form1", "Dan"))
    MsgBox (OLRetrieveKey(Me.hWnd, "Form1", "Dan", "Dan", 3))
End Sub
