VERSION 5.00
Begin VB.Form frmAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Record"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Print"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblFieldName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdd"
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

Dim intMinFormWidth As Integer





Private Sub Form_Load()
    
    Dim n As Integer
    Dim intCurTopPos As Long
      
    
    intCurTopPos = 120

    
    For n = 0 To iFieldMAX
        
        ' Load controls
        If n <> 0 Then
        
            Load Me.lblFieldName(n)
            Load Me.txtInput(n)
            
            lblFieldName(n).Visible = True
            txtInput(n).Visible = True
            
        End If
        
        lblFieldName(n).Caption = tFieldName(n).Item
        lblFieldName(n).Width = tFieldName(n).Width
        lblFieldName(n).Top = intCurTopPos
        txtInput(n).Width = tFieldName(n).Width
        txtInput(n).Top = intCurTopPos
        
        intCurTopPos = intCurTopPos + 285
        
    Next n
    
    Me.Height = intCurTopPos + 285 + cmdOk.Height + 285
    
    ' Set positions of buttons
    cmdOk.Top = intCurTopPos + cmdOk.Height - 285
    cmdCancel.Top = intCurTopPos + cmdCancel.Height - 285
    cmdPrint.Top = intCurTopPos + cmdPrint.Height - 285
    
    Me.Visible = True
    
    intCurRec = intTotRecs + 1
    
    ClearFields
    
End Sub

Sub cmdOk_Click()

    Dim Person As Person_T
    Dim n As Integer
    Dim intFreeFile As Integer
    Dim strTemp As String
    
    ' Validation
    If Len(Trim(txtInput((1)).Text)) > 10 Then
        MsgBox ("'First Name' too long (must less than 11 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((2)).Text)) > 10 Then
        MsgBox ("'SurName' too long (must be less than 11 (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput(3).Text)) > 8 Or Len(Trim(txtInput(3).Text)) < 8 Then
        MsgBox ("'Date Of Birth' should be in 'dd/mm/yy')")
        Exit Sub
    End If
    If Len(Trim(txtInput((4)).Text)) <> 1 Then
        MsgBox ("'Gender' should be 'm' for male, or 'f' for feamale.")
        Exit Sub
    End If
    If Len(Trim(txtInput((5)).Text)) > 10 Then
        MsgBox ("'Origin' too long (must be less than 11 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((6)).Text)) > 15 Then
        MsgBox ("'Main Hobby' too long (must be less than 16 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((7)).Text)) > 25 Then
        MsgBox ("'Address 1' too long (must be less than 26 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((8)).Text)) > 15 Then
        MsgBox ("'Address 2' too long (must be less than 16 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((9)).Text)) > 15 Then
        MsgBox ("'Address 3' too long (must be less than 16 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((10)).Text)) > 8 Then
        MsgBox ("'Post Code' too long (must be less than 9 charictors (e.g: AB12 3CD))")
        Exit Sub
    End If
    If Len(Trim(txtInput((11)).Text)) > 25 Then
        MsgBox ("'Telephone' too long (must be less than 26 charictors (obmit area code if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((12)).Text)) > 25 Then
        MsgBox ("'Email' too long (must be less than 26 charictors (use a shorter one if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((13)).Text)) > 10 Then
        MsgBox ("'Mentor' too long (must be less than 11 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((14)).Text)) > 10 Then
        MsgBox ("'Tutor' too long (must be less than 11 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((15)).Text)) > 10 Then
        MsgBox ("'Course 1' too long (must be less than 11 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((16)).Text)) > 10 Then
        MsgBox ("'Course 2' too long (must be less than 11 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((17)).Text)) > 10 Then
        MsgBox ("'Course 3' too long (must be less than 11 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput(18).Text)) > 8 Or Len(Trim(txtInput(18).Text)) < 8 Then
        MsgBox ("'Date Started' should be in 'dd/mm/yy')")
        Exit Sub
    End If
    txtInput(19).Text = LCase(txtInput(19).Text)
    If Not (txtInput(19).Text = "yes" Or txtInput(19).Text = "true" Or txtInput(19).Text = "1" _
      Or txtInput(19).Text = "no" Or txtInput(19).Text = "false" Or txtInput(19).Text = "0") Then
      
        MsgBox ("'Active' should be 'yes', 'true', '1', 'no', 'false' or '0'.")
        Exit Sub
    End If
    If Len(Trim(txtInput((20)).Text)) > 20 Then
        MsgBox ("'School From' too long (must be less than 21 charictors (abbreviate if necessary))")
        Exit Sub
    End If
    If Len(Trim(txtInput((21)).Text)) > 1 Then
        MsgBox ("'GCSE English' too long (must a single grade figure)")
        Exit Sub
    End If
    If Len(Trim(txtInput((22)).Text)) > 1 Then
        MsgBox ("'GCSE Maths' too long (must a single grade figure)")
        Exit Sub
    End If
    If Len(Trim(txtInput((23)).Text)) > 2 Then
        MsgBox ("'GCSE Science' too long (must a double grade figure, i.e; two letters with no seperater)")
        Exit Sub
    End If
    If Len(Trim(txtInput((23)).Text)) = 1 Then
        MsgBox ("'GCSE Science' too sort (must a double grade figure, i.e; two letters with no seperater)")
        Exit Sub
    End If
    If Int(Trim(txtInput((24)).Text)) < 0 Or Int(Trim(txtInput((24)).Text)) > 5 Then
        MsgBox ("'Key Skills Level' is invalid (must a  interger number)")
        Exit Sub
    End If
    
    ' Check if none optional information was entered
    If Len(Trim(txtInput((1)).Text)) = 0 Then
        MsgBox ("'ID' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((2)).Text)) = 0 Then
        MsgBox ("'First Name' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((3)).Text)) = 0 Then
        MsgBox ("'Surname' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((4)).Text)) = 0 Then
        MsgBox ("'Date of birth' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((6)).Text)) = 0 Then
        MsgBox ("'Origin' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((7)).Text)) = 0 Then
        MsgBox ("'Address 1' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((8)).Text)) = 0 Then
        MsgBox ("'Address 2' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((7)).Text)) = 0 Then
        MsgBox ("'ID' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((10)).Text)) = 0 Then
        MsgBox ("'Post Code' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((11)).Text)) = 0 Then
        MsgBox ("'Telephone' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((13)).Text)) = 0 Then
        MsgBox ("'Mentor' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((14)).Text)) = 0 Then
        MsgBox ("'Tutor' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((15)).Text)) = 0 Then
        MsgBox ("'Coruse 1' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((21)).Text)) = 0 Then
        MsgBox ("'GCSE English' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((22)).Text)) = 0 Then
        MsgBox ("'GCSE Maths' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((23)).Text)) = 0 Then
        MsgBox ("'GCSE Science' is not optional, please provide this information.")
        Exit Sub
    End If
    If Len(Trim(txtInput((24)).Text)) = 0 Then
        MsgBox ("'Key Skills Level' is not optional, please provide this information.")
        Exit Sub
    End If
    

    ' Add to data
    For n = 0 To iFieldMAX
        
        Select Case n
            Case 0
                If n <> -1 Then
                    ' Check if record already exists
                
                    Person.ID = Int(Val(txtInput(n).Text))
                End If
                
            Case 1
                    Person.FirstName = txtInput(n).Text
            Case 2
                    Person.SurName = txtInput(n).Text
            Case 3
                    Person.DOB = txtInput(n).Text
            Case 4
                    Person.Gender = txtInput(n).Text
            Case 5
                    Person.EthOrigin = txtInput(n).Text
            Case 6
                    Person.MainHobby = txtInput(n).Text

            Case 7
                    Person.Address1 = txtInput(n).Text
            Case 8
                    Person.Address2 = txtInput(n).Text
            Case 9
                    Person.Address3 = txtInput(n).Text
            Case 10
                    Person.PostCode = txtInput(n).Text
            Case 11
                    Person.Telephone = txtInput(n).Text
            Case 12
                    Person.Email = txtInput(n).Text

            Case 13
                    Person.Mentor = txtInput(n).Text
            Case 14
                    Person.Tutor = txtInput(n).Text
            Case 15
                    Person.Coruse1 = txtInput(n).Text
            Case 16
                    Person.Coruse2 = txtInput(n).Text
            Case 17
                    Person.Coruse3 = txtInput(n).Text
            Case 18
                    Person.DateStarted = txtInput(n).Text
            Case 19
                    
                    strTemp = LCase(txtInput(n).Text)
                    
                    Select Case strTemp
                        Case "true"
                            Person.Active = True
                        Case "yes"
                            Person.Active = True
                        Case "1"
                            Person.Active = True
                            
                        Case "false"
                            Person.Active = False
                        Case "no"
                            Person.Active = False
                        Case "0"
                            Person.Active = False
                    End Select
                    
            Case 20
                    Person.SchoolFrom = txtInput(n).Text
            Case 21
                    Person.GCSE_English = txtInput(n).Text
            Case 22
                    Person.GCSE_Maths = txtInput(n).Text
            Case 23
                    Person.GCSE_Science = txtInput(n).Text
            Case 24
                    Person.KeySkillsLev = Val(txtInput(n).Text)

            
        End Select
            
    Next n

    
    ' Write to file
    intFreeFile = FreeFile
    Open strCurFile For Random As intFreeFile Len = Len(Person)
        Put #intFreeFile, intCurRec, Person
    Close #intFreeFile
    

    ' Update the list
    frmList.UpDateList
    
    
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    On Error GoTo PrintErr
    
    ' Print the form
    Me.PrintForm
    
    Resume
    
    MsgBox "Printed Succesfully", , "VBPC 4"
    
  Exit Sub
 
PrintErr:
    MsgBox "Could not Print (Check printer setup)", vbCritical, "VBPC 4"
    
'    Resume
End Sub


Sub ClearFields()
    
    Dim n As Integer
    
    For n = 0 To iFieldMAX
        txtInput(n).Text = ""
    Next n
    
End Sub


Sub ModifiyRecord()
   
    Dim Person As Person_T
    Dim intFreeFile As Integer
    
    Dim nf As Integer
    Dim n As Integer
    
    ' Read from file
    intFreeFile = FreeFile
    
    Open strCurFile For Random As intFreeFile Len = Len(Person)
    
        Seek intFreeFile, intCurRec
        Get intFreeFile, , Person
    
        For n = 0 To iFieldMAX
        
            ' Add to list
            Select Case n
                Case 0
                        txtInput(n).Text = Trim(Str(Person.ID))
                Case 1
                        txtInput(n).Text = Trim(Person.FirstName)
                Case 2
                        txtInput(n).Text = Trim(Person.SurName)
                Case 3
                        txtInput(n).Text = Trim(Person.DOB)
                Case 4
                        txtInput(n).Text = Trim(Person.Gender)
                Case 5
                        txtInput(n).Text = Trim(Person.EthOrigin)
                Case 6
                        txtInput(n).Text = Trim(Person.MainHobby)
    
                Case 7
                        txtInput(n).Text = Trim(Person.Address1)
                Case 8
                        txtInput(n).Text = Trim(Person.Address2)
                Case 9
                        txtInput(n).Text = Trim(Person.Address3)
                Case 10
                        txtInput(n).Text = Trim(Person.PostCode)
                Case 11
                        txtInput(n).Text = Trim(Person.Telephone)
                Case 12
                        txtInput(n).Text = Trim(Person.Email)
    
                Case 13
                        txtInput(n).Text = Trim(Person.Mentor)
                Case 14
                        txtInput(n).Text = Trim(Person.Tutor)
                Case 15
                        txtInput(n).Text = Trim(Person.Coruse1)
                Case 16
                        txtInput(n).Text = Trim(Person.Coruse2)
                Case 17
                        txtInput(n).Text = Trim(Person.Coruse3)
                Case 18
                        txtInput(n).Text = Trim(Person.DateStarted)
                Case 19
                        txtInput(n).Text = Trim(Person.Active)
    
                Case 20
                        txtInput(n).Text = Trim(Person.SchoolFrom)
                Case 21
                        txtInput(n).Text = Trim(Person.GCSE_English)
                Case 22
                        txtInput(n).Text = Trim(Person.GCSE_Maths)
                Case 23
                        txtInput(n).Text = Trim(Person.GCSE_Science)
                Case 24
                        txtInput(n).Text = Trim(Str(Person.KeySkillsLev))
                
                
            End Select
                
        Next n
    
    Close intFreeFile
    
End Sub

