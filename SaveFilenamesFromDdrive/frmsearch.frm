VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   0  'None
   Caption         =   "a"
   ClientHeight    =   1410
   ClientLeft      =   5685
   ClientTop       =   4020
   ClientWidth     =   3900
   Icon            =   "frmsearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox lstdirs 
      Height          =   1425
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox lstfiles 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please wait..."
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Daniel Downes : Ex-D Software Development(TM)
' Sent to PSC 08-06-2002
' Email and/or MSN Messenger: exd_founder@hotmail.com

'
'               http://www.Ex-D.net
'

' Description:
'  This little program will search and record or file with
'   any or a selected filetype. It uses the Windows API to
'   make the scan very fast, and then save every filename
'   it finds to a text file.
'  I made this as I wanted to catalogue all the files on my
'   CDs; I simply imported the filelist into Excel where it
'   could be sort, etc.
'
'  Hopefully someone out there might find it as useful as I
'   have. If you do, please contact me (above), thanks.



Option Explicit

Dim strPlaceToSearch As String
Dim strTypeOfFiles As String


Private Sub Form_Load()
    Dim n As Integer

    lstfiles.Clear
    
    strPlaceToSearch = InputBox("Folder to start logging from (include sub-folders)? - Eg; 'C:\MyFolder\")
    strTypeOfFiles = InputBox("Type of files to record (extention)? Eg; '*.*' for all, '*.bmp' for only 'bitmaps'")
    
    
    ''''''''''''''''''''''''  Place to seach  ''''''''''''''''''''''''''
    lstdirs.AddItem strPlaceToSearch
    
    frmStatus.Visible = True
    
    Do
        
        '''''''''''''''''''''''''''  File types to be recorded  '''''''''''''''''''''
        
        findfilesapi lstdirs.List(0), strTypeOfFiles

         
        lstdirs.RemoveItem 0
    Loop Until lstdirs.ListCount = 0
    
    
    'Copy ALL filename with addresses to a file
    
    ''''''''''''''''''''''''''''''  Text File to Save name to  ''''''''''''''''''''''''''''''
    Open App.Path + "\Output.txt" For Output As #1

    For n = 0 To lstfiles.ListCount
        Print #1, lstfiles.List(n)
    Next n
    
    frmStatus.Visible = False
    
    MsgBox "Finished, see: " + App.Path + "\Output.txt"
        
    End
    
End Sub
            
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

