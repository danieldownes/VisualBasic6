VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PCVB 4 - Daniel Downes"
   ClientHeight    =   3225
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   6000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstOutput 
      BackColor       =   &H00FEEED8&
      Height          =   2400
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblFieldName 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   615
   End
   Begin VB.Image imgBottom 
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":000C
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   6015
   End
   Begin VB.Image imgTop 
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":008E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileDel 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBackup 
         Caption         =   "&Backup"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Records"
      Begin VB.Menu mnuRecAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuRecDup 
         Caption         =   "D&uplicate"
      End
      Begin VB.Menu mnuRecModifiy 
         Caption         =   "&Modify"
      End
      Begin VB.Menu mnuRecDel 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecSort 
         Caption         =   "&Sort"
      End
      Begin VB.Menu mnuRecSearch 
         Caption         =   "Sea&rch"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewSelect 
         Caption         =   "&Select Fields"
      End
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Thank you for trying this code. If you have any problems or queries please
'  contact me:
'
'     exd_founder@hotmail.com      (you may also add me to your MSN Messenger contacts)
'
'  Jump to my site, to find other software created by myself;
'
'       http://www.Ex-D.net
'
'   Daniel Downes(UK)  -  Ex-D Software Development(TM)
'
' NOTE: Do not use this code for anything without my permission.
'        I'll probably let you use it, but you must let me know how it is being used.

Option Explicit

' Used in sorting process
Private Type SortData_T
    strData As String
    intOrder As Integer
End Type



' Program control varables


Public intNfieldsLoaded As Integer     ' Stores the total number of
                                       '  listboxes loaded
Dim blnLoaded As Boolean
                           
Dim intMwidth As Double

Private Sub Form_Load()
        
    blnLoaded = False
    
    intCurRec = 1
    
    intNfieldsLoaded = 0
    
    InitFieldNames
    
    SetVisibleFeilds
    
    If MsgBox("Open existing file? ('No' = New)", vbYesNo, "PCVB 4") = vbYes Then
        mnuFileOpen_Click
    Else
        mnuFileNew_Click
    End If
    
    blnLoaded = True
    
End Sub


' Menu functions...

Private Sub mnuFileNew_Click()
    
    ' Make a backup?
    If blnLoaded = True Then
        If MsgBox("Do you wish to make a backup copy of the current file before creating new?", vbYesNo, "VBPC 4") = vbYes Then
            mnuFileBackup_Click
        End If
    End If
    
    CommonDialog.DialogTitle = "Enter New Filename"
    CommonDialog.ShowOpen
    CommonDialog.Filter = "*.*"
    strCurFile = CommonDialog.FileName
    
    
    ' Check if location is possable to open
    
    On Error GoTo MakeF
    
    Open strCurFile For Output As #1
    Close #1
    
  Exit Sub
    
    
MakeF:
    
    If MsgBox("Location does not exist (read only, etc), would you like to try another.", vbYesNo, "Location not accesable!") = vbYes Then
        
        mnuFileNew_Click
    
    Else
    
        MsgBox "Qutting: Unable to initallise program without a file reference", , "Qutting PCVB 4"
        
        End
        
    End If
    
  Exit Sub
 
End Sub

Private Sub mnuFileOpen_Click()
    
    ' Make a backup?
    If blnLoaded = True Then
        If MsgBox("Do you wish to make a backup copy of the current file before opening another?", vbYesNo, "VBPC 4") = vbYes Then
            mnuFileBackup_Click
        End If
    End If
    
    CommonDialog.DialogTitle = "Select file to open"
    CommonDialog.ShowOpen
    CommonDialog.Filter = "*.*"
    strCurFile = CommonDialog.FileName
    

    On Error GoTo MakeF
    
    ' If file does Not exist, there will be an Error
    ' Attempt to read in data

    Open strCurFile For Input As #1
    Close #1
    
       
    UpDateList

  Exit Sub
  
    
MakeF:
    
    If MsgBox("The specified file does not exist, do you wish to create?", vbYesNo, "VBPC 4") <> vbYes Then
        mnuFileNew_Click
    Else
        mnuFileOpen_Click
    End If
 
    
End Sub

Private Sub mnuFileSaveAs_Click()
    
    Dim strNewFile
    
    CommonDialog.DialogTitle = "Enter new file's name"
    CommonDialog.ShowOpen
    CommonDialog.Filter = "*.*"
    strNewFile = CommonDialog.FileName
    
    
    ' Attempt to copy current file
    On Error GoTo CopyError
    FileCopy strCurFile, strNewFile
    
    strCurFile = strNewFile
    
  Exit Sub

CopyError:
    MsgBox "Error saving file", , "VBPC 4"
End Sub

Private Sub mnuFileClose_Click()
    
    ' Make a backup?
    If MsgBox("Do you wish to make a backup copy of the current file before closeing it?", vbYesNo, "VBPC 4") = vbYes Then
        mnuFileBackup_Click
    End If
    
    Close
    
    MsgBox "'" + strCurFile + "' has been closed", , "VBPC 4"
    
    strCurFile = ""
    
    Form_Load
    
End Sub

Private Sub mnuFileDel_Click()

    If MsgBox("Are you sure you wish to delete this file and all data within it?", vbYesNo, "PCVB 4") = vbYes Then
    
        ' Make a backup?
        If MsgBox("Do you wish to make a backup copy of the current file before deleting the main copy?", vbYesNo, "VBPC 4") = vbYes Then
            mnuFileBackup_Click
        End If
    
        On Error GoTo FileDelErr
    
        Kill strCurFile
        
        strCurFile = ""
        
        MsgBox "Current File has been deleted", , "PCVB 4"
        
        Form_Load
        
    End If
    
  Exit Sub
  
FileDelErr:
    
    MsgBox "The current file could not be deleted (check write protect, etc)", , "PCVB 4"

End Sub

Private Sub mnuFileBackup_Click()
    
    Dim strBakUpFile As String
    Dim intDotPos As Integer
    
    intDotPos = InStr(strCurFile, ".")
    
    If intDotPos <> 0 Then
        ' Remake the file extention
        strBakUpFile = Left(strCurFile, intDotPos) + "bak"
    Else
        ' There was no file extention on the filename
        strBakUpFile = strCurFile + ".bak"
    End If
    
    CommonDialog.DialogTitle = "Enter backup filename"
    CommonDialog.ShowOpen
    CommonDialog.FileName = strBakUpFile
    strBakUpFile = CommonDialog.FileName
    
    
    ' Attempt to copy current file
    On Error GoTo CopyError
    FileCopy strCurFile, strBakUpFile
    
  Exit Sub

CopyError:
    MsgBox "Error backing up File", , "VBPC 4"

End Sub

Private Sub mnuFileRestore_Click()
    ' Make a backup?
    If blnLoaded = True Then
        If MsgBox("Do you wish to make a backup copy of the current file before opening another?", vbYesNo, "VBPC 4") = vbYes Then
            mnuFileBackup_Click
        End If
    End If
    
    CommonDialog.DialogTitle = "Select backup file to restore"
    CommonDialog.ShowOpen
    CommonDialog.Filter = "*.bak"
    strCurFile = CommonDialog.FileName
    

    On Error GoTo MakeF
    
    ' If file does Not exist, there will be an Error
    ' Attempt to read in data

    Open strCurFile For Input As #1
    Close #1
    
       
    UpDateList

  Exit Sub
  
    
MakeF:
    
    If MsgBox("The specified file does not exist, do you wish to create?", vbYesNo, "VBPC 4") <> vbYes Then
        mnuFileNew_Click
    Else
        mnuFileOpen_Click
    End If
End Sub

Private Sub mnuFileExit_Click()
    
    If MsgBox("Are you sure you wish to exit?", vbYesNo, "VBPC 4") = vbYes Then
        
        ' Make a backup?
        If MsgBox("Do you wish to make a backup copy of the current file before exiting?", vbYesNo, "VBPC 4") = vbYes Then
            mnuFileBackup_Click
        End If
        
        End
        
    End If
    
End Sub

Private Sub mnuFilePrint_Click()
    
    On Error GoTo PrintErr
    
    Me.PrintForm
    
    Resume
    
    MsgBox "Printed Succesfully", , "VBPC 4"
    
  Exit Sub
 
PrintErr:
    MsgBox "Could not Print (Check printer setup)", vbCritical, "VBPC 4"
    
'    Resume

End Sub



Private Sub mnuRecAdd_Click()

    Load frmAdd
    frmAdd.Caption = "Add New Record"
    
End Sub

Private Sub mnuRecDup_Click()

    Dim Person As Person_T
    Dim intFreeFile As Integer
   

    If MsgBox("Are you sure you wish to duplicate the currently selected record?", vbYesNo, "VBPC 4") = vbYes Then
    
        ' Copy each record to the new file...
    
        ' Read in currrent record
        intFreeFile = FreeFile
        
        Open strCurFile For Random As intFreeFile Len = Len(Person)
        
            Seek intFreeFile, intCurRec
            Get intFreeFile, , Person
            
        Close intFreeFile
        
        
        ' Save as new record
        intFreeFile = FreeFile
        Open strCurFile For Random As intFreeFile Len = Len(Person)
            Put #intFreeFile, intTotRecs + 1, Person
        Close #intFreeFile
        
        
        ' Update the list
        frmList.UpDateList
        
    End If
    
End Sub

Private Sub mnuRecModifiy_Click()
    
    Dim intTemp As Integer
    
    intTemp = intCurRec
        
    ' Make Input form act as a modification form
    Load frmAdd
    frmAdd.Caption = "Modifiy Record " + Trim(Str(intTemp))
    
    ' Reset 'intCurRec' (as it gets modified in 'frmAdd')
    intCurRec = intTemp
    
    frmAdd.ModifiyRecord
    
End Sub

Private Sub mnuRecDel_Click()
    
    Dim Person As Person_T
    
    Dim n As Integer
    Dim nn As Integer
    Dim intFreeFile As Integer
    Dim int2edFree As Integer
    
    
    
    If MsgBox("Are you sure you wish to delete the currently selected record?", vbYesNo, "VBPC 4") = vbYes Then
    
        
        ' Move all record beond the currenty selected down one to overwrite the currently selected
        
        intFreeFile = FreeFile
        
        Open strCurFile For Random As intFreeFile Len = Len(Person)
        
            int2edFree = FreeFile
            
            Open strCurFile + ".temp" For Random As int2edFree Len = Len(Person)
        
            nn = 1
            
            For n = 1 To intTotRecs

                If n <> intCurRec Then
                    Get intFreeFile, n, Person
                    
                    Put int2edFree, nn, Person
                    
                    nn = nn + 1
                End If
                                
            Next n
           
        Close                                           ' Close all
        
        Kill strCurFile                                 ' Delete old file
        
        Name strCurFile + ".temp" As strCurFile         ' Rename new file so to replace old
        
        intTotRecs = intTotRecs - 1
        
        
        ' Update the list
        frmList.UpDateList
        
    End If
    
End Sub

Private Sub mnuRecSort_Click()
' Sorts the first column into accending order
    
    Dim SortData(200) As SortData_T
    Dim TempSwapCell As SortData_T
    Dim blnSwaped As Boolean
    
    Dim Person As Person_T
    
    Dim n As Integer
    Dim nr As Integer
    
    Dim intFreeFile As Integer
    Dim int2edFree As Integer


    If MsgBox("Are you should you wish to reorder the records in assending order of the first column", vbYesNo, "PCVB 4") = vbYes Then
        
        
        ' First get all data of the first column
        For n = 0 To intTotRecs - 1
            SortData(n).strData = lstOutput(0).List(n)
            SortData(n).intOrder = n
        Next n
        
        
        ' Now do the sort process (Bubble Sort)
        Do
        
            blnSwaped = False           ' Set flag; list is sorted when this stays unchanged
            
            For n = 0 To intTotRecs - 2
            
                If SortData(n).strData > SortData(n + 1).strData Then
                    
                    ' Swap
                    TempSwapCell = SortData(n)
                    SortData(n) = SortData(n + 1)
                    SortData(n + 1) = TempSwapCell
                    
                    blnSwaped = True
                    
                End If
            
            Next n
            
        Loop Until blnSwaped = False
        
        
        '  Now rewrite the current file in the order as based in the sort process
        intFreeFile = FreeFile
        
        Open strCurFile For Random As intFreeFile Len = Len(Person)
        
            int2edFree = FreeFile
            
            Open strCurFile + ".temp" For Random As int2edFree Len = Len(Person)
        
                For nr = 1 To intTotRecs
                
                    Get intFreeFile, SortData(nr - 1).intOrder + 1, Person
                    
                    Put int2edFree, nr, Person
                
                Next nr
                
                
        Close                                           ' Close all files
        
        Kill strCurFile                                 ' Delete old file
        
        Name strCurFile + ".temp" As strCurFile         ' Rename new file so to replace old
        
    End If
    
    ' Show changes
    UpDateList

End Sub

Private Sub mnuRecSearch_Click()
    
    Load frmSearch
    
    frmSearch.Visible = True
    
End Sub

Private Sub mnuViewSelect_Click()

    Load frmSelectFeilds
    
    frmSelectFeilds.Visible = True
    
End Sub


Private Sub lstOutput_Click(Index As Integer)
    
    Dim n As Integer

    ' Find current record
    For n = 0 To lstOutput(Index).ListCount
        If lstOutput(Index).Selected(n) = True Then
            intCurRec = n + 1
            
            Exit For
            
        End If
    Next n
    
    ' Select all other resective fields
    For n = 0 To intNfieldsLoaded - 1
        If intIndexOrder(n) <> -1 Then
            lstOutput(n).Selected(intCurRec - 1) = True
        End If
    Next n
    
End Sub



' Record List functions...

Sub UpDateList()

    Dim Person As Person_T
    Dim intFreeFile As Integer
    
    Dim nf As Integer
    Dim n As Integer
    
    ' Read from file
    intFreeFile = FreeFile
    
    Open strCurFile For Random As intFreeFile Len = Len(Person)
    
        intTotRecs = LOF(intFreeFile) / Len(Person)
        
        
        For nf = 1 To intTotRecs
    
            Seek intFreeFile, nf
            Get intFreeFile, , Person
        
            For n = 0 To iFieldMAX
            
                ' Clear each list on refresh
                If nf = 1 Then
                    lstOutput(n).Clear
                End If
            
                ' Add to list
                If intIndexOrder(n) <> -1 Then
                    Select Case intIndexOrder(n)
                        Case 0
                                lstOutput(n).AddItem Str(Person.ID)
                        Case 1
                                lstOutput(n).AddItem Person.FirstName
                        Case 2
                                lstOutput(n).AddItem Person.SurName
                        Case 3
                                lstOutput(n).AddItem Person.DOB
                        Case 4
                                lstOutput(n).AddItem Person.Gender
                        Case 5
                                lstOutput(n).AddItem Person.EthOrigin
                        Case 6
                                lstOutput(n).AddItem Person.MainHobby
            
                        Case 7
                                lstOutput(n).AddItem Person.Address1
                        Case 8
                                lstOutput(n).AddItem Person.Address2
                        Case 9
                                lstOutput(n).AddItem Person.Address3
                        Case 10
                                lstOutput(n).AddItem Person.PostCode
                        Case 11
                                lstOutput(n).AddItem Person.Telephone
                        Case 12
                                lstOutput(n).AddItem Person.Email
            
                        Case 13
                                lstOutput(n).AddItem Person.Mentor
                        Case 14
                                lstOutput(n).AddItem Person.Tutor
                        Case 15
                                lstOutput(n).AddItem Person.Coruse1
                        Case 16
                                lstOutput(n).AddItem Person.Coruse2
                        Case 17
                                lstOutput(n).AddItem Person.Coruse3
                        Case 18
                                lstOutput(n).AddItem Person.DateStarted
                        Case 19
                                lstOutput(n).AddItem Person.Active
            
                        Case 20
                                lstOutput(n).AddItem Person.SchoolFrom
                        Case 21
                                lstOutput(n).AddItem Person.GCSE_English
                        Case 22
                                lstOutput(n).AddItem Person.GCSE_Maths
                        Case 23
                                lstOutput(n).AddItem Person.GCSE_Science
                        Case 24
                                lstOutput(n).AddItem Person.KeySkillsLev
                        
                    End Select
                    
                End If
                    
            Next n
            
        Next nf
    
    Close intFreeFile
    
End Sub


Sub SetVisibleFeilds()
    
    Dim intCurPos As Long
    Dim n As Integer
    
    intCurPos = 120
    
    For n = 0 To iFieldMAX
        
        If intIndexOrder(n) <> -1 Then
            
            ' Load controls
            If n <> 0 And intNfieldsLoaded < n Then
            
                Load Me.lblFieldName(n)
                Load Me.lstOutput(n)
                
                intNfieldsLoaded = intNfieldsLoaded + 1
                
            End If
            
            lblFieldName(n).Visible = True
            lstOutput(n).Visible = True
            
            lblFieldName(n).Caption = tFieldName(intIndexOrder(n)).Item
            lblFieldName(n).Width = tFieldName(intIndexOrder(n)).Width
            lblFieldName(n).Left = intCurPos
            lstOutput(n).Width = tFieldName(intIndexOrder(n)).Width
            lstOutput(n).Left = intCurPos
            
            intCurPos = intCurPos + tFieldName(intIndexOrder(n)).Width
            
            Else
            
                lblFieldName(n).Visible = False
                lstOutput(n).Visible = False
 
        End If
        
    Next n
    
    intMwidth = intCurPos + 200
    Me.Width = intCurPos + 200
    
    
End Sub




' Misc form functions
Private Sub Form_Resize()

    Dim n As Integer
    
    If Me.WindowState = vbNormal Then
    
        If Me.Height < 2000 Then Me.Height = 2000
        
        imgTop.Width = Me.Width
        imgBottom.Width = Me.Width
        imgBottom.Top = Me.Height - (imgBottom.Height * 4) '+ 50
        
        ' Resize existing fields
        For n = 0 To intNfieldsLoaded
            lstOutput(n).Height = Me.Height - 240 - lstOutput(n).Top - 700
        Next n
        
        Me.Width = intMwidth
        
    End If
    
End Sub

Private Sub Form_Terminate()
    ' Make a backup?
    If MsgBox("Do you wish to make a backup copy of the current file before exiting?", vbYesNo, "VBPC 4") = vbYes Then
        mnuFileBackup_Click
    End If
End Sub
