VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search for Record"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4290
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMatchCase 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Match Case"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdNewSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&New Search"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "< Enter part or all your word here>"
      Top             =   240
      Width           =   2535
   End
   Begin VB.OptionButton optFromCur 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search from current &record"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.OptionButton optAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search &whole file"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Find"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearch"
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

Dim intStartingField As Integer

Private Sub Form_Load()
    ' Select defulat text ready to be overwritten or deleted
    txtSearch.SelLength = Len(txtSearch.Text)
End Sub

Private Sub Form_Terminate()
    cmdCancel_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdCancel_Click
End Sub

Private Sub txtSearch_Change()
    If Len(txtSearch.Text) <> 0 Then
        cmdFind.Enabled = True
    Else
        cmdFind.Enabled = False
    End If
End Sub

Private Sub cmdFind_Click()

    Dim intRec As Integer
    Dim strData As String
    
    Dim Person As Person_T
    Dim intFreeFile As Integer
    Dim n As Integer
    Dim nn As Integer
    
    ' Read from file
    intFreeFile = FreeFile
    
    ' Set starting record to search
    If optAll.Value = True Then intRec = 1 Else intRec = intCurRec
    
    If cmdFind.Caption <> "&Find Next" Then intStartingField = 0

    txtSearch.Enabled = False
    cmdNewSearch.Enabled = True
    optAll.Enabled = False
    optFromCur.Enabled = False
    chkMatchCase.Enabled = False
    
    Open strCurFile For Random As intFreeFile Len = Len(Person)
    
        Do
            
            Seek intFreeFile, intRec
            Get intFreeFile, , Person
        
            For n = intStartingField To iFieldMAX
            
                ' Search each feild
                Select Case intIndexOrder(n)
                    Case 0
                            strData = Str(Person.ID)
                    Case 1
                            strData = Person.FirstName
                    Case 2
                            strData = Person.SurName
                    Case 3
                            strData = Person.DOB
                    Case 4
                            strData = Person.Gender
                    Case 5
                            strData = Person.EthOrigin
                    Case 6
                            strData = Person.MainHobby
        
                    Case 7
                            strData = Person.Address1
                    Case 8
                            strData = Person.Address2
                    Case 9
                            strData = Person.Address3
                    Case 10
                            strData = Person.PostCode
                    Case 11
                            strData = Person.Telephone
                    Case 12
                            strData = Person.Email
        
                    Case 13
                            strData = Person.Mentor
                    Case 14
                            strData = Person.Tutor
                    Case 15
                            strData = Person.Coruse1
                    Case 16
                            strData = Person.Coruse2
                    Case 17
                            strData = Person.Coruse3
                    Case 18
                            strData = Person.DateStarted
                    Case 19
                            strData = Person.Active
        
                    Case 20
                            strData = Person.SchoolFrom
                    Case 21
                            strData = Person.GCSE_English
                    Case 22
                            strData = Person.GCSE_Maths
                    Case 23
                            strData = Person.GCSE_Science
                    Case 24
                            strData = Person.KeySkillsLev
                                      
                End Select
                
                ' Check to 'match case'
                If chkMatchCase.Value <> 1 Then
                    strData = LCase(strData)
                    txtSearch.Text = LCase(txtSearch.Text)
                End If
                
                ' Check data with user's input
                If InStr(strData, txtSearch.Text) <> 0 Then
                    intCurRec = intRec
                      
                    ' Select all other respective fields
                    For nn = 0 To frmList.intNfieldsLoaded
                        frmList.lstOutput(n).Selected(intCurRec - 1) = True
                    Next nn
                    
                    ' Set focus of field containing found match
                    frmList.SetFocus
                    frmList.lstOutput(n).SetFocus
                    
                    For nn = 0 To iFieldMAX
                        frmList.lstOutput(nn).BackColor = &HFEEED8
                    Next nn
                    frmList.lstOutput(n).BackColor = vbBlue
                    
                    Beep
                    
                    cmdFind.Caption = "&Find Next"
                    optFromCur.Value = True
                    
                    intStartingField = n + 1
                    
                    Exit Sub
                    
                End If
                    
            Next n
            
            intStartingField = 0
            
            intRec = intRec + 1
        Loop Until intRec > intTotRecs
        
    Close
    
    MsgBox "Finished searching", , "PCVB 4"
    
    Unload Me
    
End Sub

Private Sub cmdNewSearch_Click()
    Close           ' The user can begin a new search at any time, so make sure any files are closed
    
    txtSearch.Text = "< Enter part or all your word here>"
    txtSearch.Enabled = True
    txtSearch.SelLength = Len(txtSearch.Text)
    cmdFind.Enabled = False
    cmdNewSearch.Enabled = False
    optAll.Enabled = True
    optAll.Value = True
    optFromCur.Enabled = True
    chkMatchCase.Enabled = True
    
    intStartingField = 0
    
End Sub

Private Sub cmdCancel_Click()

    Dim nn As Integer

    Close           ' The user can cancel at any time, so make sure any files are closed
    
    
    ' Null any existing highlighting that is a result of searching
    For nn = 0 To iFieldMAX
        If intIndexOrder(nn) <> -1 Then
            frmList.lstOutput(intIndexOrder(nn)).BackColor = &HFEEED8
        End If
    Next nn
    
    Unload Me
End Sub



