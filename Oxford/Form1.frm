VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "This show the operation and words that are being extracted."
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer
Dim temp As String
Dim thedata As String
Dim totalnum As Integer
Dim placestart As Integer
Dim tagg As Integer
Dim placeend As Integer

Private Sub Command1_Click()
    
    'Open input file
    Text1.Text = "Opening file..."
    Open "c:\oxford.dat" For Input As #1
    Text1.Text = "File opened."
    
    'Create output file
    Open "c:\words.dat" For Output As #2

    
    Do
        Input #1, temp
        tagg = 0
        For n = 1 To Len(temp)
            If Mid$(temp, n, 4) = "<hw>" Then
                placestart = n
                n = Len(temp)
                tagg = 1
            End If
        Next n
        
        
        For n = 1 To Len(temp)
            If Mid$(temp, n, 5) = "</hw>" And tagg = 1 Then
                placeend = n
                n = Len(temp)
                tagg = 2
            End If
        Next n
        
        If tagg = 2 Then
        
            thedata = Mid$(temp, placestart + 4, placeend - 5)
            Print #2, thedata

            Text1.Text = thedata
        
            Text1.Refresh
            
        End If
    Loop Until EOF(1)
    
    Text1.Text = "Done."
End Sub
