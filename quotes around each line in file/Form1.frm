VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim sTemp As String
    
    Open App.Path & "\file.txt" For Input As 1
    Open App.Path & "\filenew.txt" For Append As 2
    
    Do
        Input #1, sTemp
        
        'sTemp = Chr(34) & sTemp & Chr(34) & ","
        
        Print #2, sTemp
        
    Loop Until EOF(1)
    
    Close
End Sub
