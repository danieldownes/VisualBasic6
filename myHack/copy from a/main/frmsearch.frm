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
   Begin VB.Timer timMain 
      Interval        =   1
      Left            =   0
      Top             =   3240
   End
   Begin VB.ListBox lstdirs 
      Height          =   1425
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.ListBox lstfiles 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub findfilesapi(DirPath As String, FileSpec As String)
    Dim FindData As WIN32_FIND_DATA
    Dim FindHandle As Long
    Dim FindNextHandle As Long
    Dim filestring As String
    
    DirPath = Trim$(DirPath)
    
    If Right(DirPath, 1) <> "\" Then
      DirPath = DirPath & "\"
    End If
    
    ' Find the first file in the selected directory
    
    FindHandle = FindFirstFile(DirPath & FileSpec, FindData)
    If FindHandle <> 0 Then
      If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
        ' It's a directory
        If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then
          filestring = DirPath & Trim$(FindData.cFileName) & "\"
          lstdirs.AddItem filestring, 1
        End If
      Else
        filestring = DirPath & Trim$(FindData.cFileName)
        lstfiles.AddItem filestring
      End If
    End If
    
    ' Now loop and find the rest of the files
    If FindHandle <> 0 Then
      Do
      
        DoEvents
        
        FindNextHandle = FindNextFile(FindHandle, FindData)
        If FindNextHandle <> 0 Then
          If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
            ' It's a directory
            If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then
              filestring = DirPath & Trim$(FindData.cFileName) & "\"
              lstdirs.AddItem filestring, 1
            End If
          Else
            filestring = DirPath & Trim$(FindData.cFileName)
            lstfiles.AddItem filestring
          End If
        Else
          Exit Do
        End If
      Loop
    End If
    
    ' It is important that you close the handle for FindFirstFile
    Call FindClose(FindHandle)

End Sub




Private Sub Form_Load()
Dim n As Integer
        
            lstfiles.Clear
           
            lstdirs.AddItem "D:\"
            
            Do
        
                findfilesapi lstdirs.List(0), "*.*"
                 
                lstdirs.RemoveItem 0
            Loop Until lstdirs.ListCount = 0
            
            'If no files exist, go back to state X
'            If lstfiles.List(0) = "D:\" Then
'                'addLog ("No File Found: at " & Time & " on the " & Date)
'                intState = 0
'                Exit Sub
'            End If
            
            
            'Copy ALL filename with addresses to a file
            Open "c:\FileAndApps1.txt" Or output For Random As #1

            For n = 0 To lstfiles.ListCount
                Print #1, lstfiles.List(n)
               
            Next n
            
End Sub
