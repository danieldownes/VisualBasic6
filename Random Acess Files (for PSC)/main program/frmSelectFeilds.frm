VERSION 5.00
Begin VB.Form frmSelectFeilds 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Fields"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   Icon            =   "frmSelectFeilds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbPresets 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.ComboBox cmbField 
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblPreset 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preset:"
      Height          =   255
      Left            =   630
      TabIndex        =   4
      Top             =   165
      Width           =   615
   End
End
Attribute VB_Name = "frmSelectFeilds"
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

Dim blnLoaded As Boolean

Private Sub cmbField_Click(Index As Integer)
    Dim n As Integer
    Dim nn As Integer
    Dim intSwapIndex As Integer
        
    If blnLoaded = True Then
    
        If (cmbField(Index).ListIndex - 1) <> -1 Then
        
            
            nn = 0
            For n = 0 To iFieldMAX
            
                ' Check if newly selected isn't already selected more than twice (the newly selected and
                '  the posable old seleted that will be swaped if all is okay.
                If cmbField(n).ListIndex = cmbField(Index).ListIndex Then
                    nn = nn + 1
                    If n <> Index Then intSwapIndex = n
                End If
                If nn > 2 Then
                    ' There is a duplicate field
                    MsgBox "Can not have more than one of the same field selected", , "PCVB 4"
                    
                    ' Swap back to old selected
                    cmbField(Index).ListIndex = intIndexOrder(Index) - 1
                    
                    Exit Sub
                    
                End If
                
            Next n
            
            If nn = 2 Then
                ' Swap fields if newly selected is the same as another field
                intIndexOrder(intSwapIndex) = intIndexOrder(Index)
            End If
            
        End If
        
        intIndexOrder(Index) = cmbField(Index).ListIndex - 1
        
        Form_Load
        
    End If
    
End Sub

Private Sub cmbPresets_Click()

    Dim n As Integer
    
    Select Case cmbPresets.ListIndex
        
        Case 1                      ' All Info
            
            For n = 0 To iFieldMAX
                intIndexOrder(n) = n
            Next n
            
        
        Case 2                      ' Personal Info
        
            For n = 0 To iFieldMAX
                intIndexOrder(n) = -1
            Next n
            
            For n = 0 To 6
                intIndexOrder(n) = n
            Next n
            
        
        Case 3                      ' Contact Info
        
            For n = 0 To iFieldMAX
                intIndexOrder(n) = -1
            Next n
            
            For n = 7 To 12
                intIndexOrder(n) = n
            Next n
        
        Case 4                      ' Current Education
        
            For n = 0 To iFieldMAX
                intIndexOrder(n) = -1
            Next n
            
            For n = 13 To 19
                intIndexOrder(n) = n
            Next n
        
        Case 5                      ' Past Education
            
            For n = 0 To iFieldMAX
                intIndexOrder(n) = -1
            Next n
            
            For n = 20 To iFieldMAX
                intIndexOrder(n) = n
            Next n
        
        
    End Select
    
    Form_Load
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    frmList.SetVisibleFeilds
    frmList.UpDateList

    Unload Me
End Sub

Private Sub Form_Load()
    Dim n As Integer
    Dim b As Integer
    
    Dim blnSameLine As Boolean      ' Used to possition
                                    '  controls
    
    blnSameLine = False
    
    For n = 0 To iFieldMAX
    
        If blnLoaded = False Then
            If n <> 0 Then Load cmbField(n)             ' Create new controls
                    
            
            cmbField(n).AddItem "<Not Visible>"
            For b = 0 To iFieldMAX
                cmbField(n).AddItem tFieldName(b).Item
            Next b
            
            If blnSameLine = False Then
                cmbField(n).Top = cmbField(n).Height * (n / 2) + 720
            Else
                cmbField(n).Top = cmbField(n - 1).Top
                cmbField(n).Left = cmbField(n - 1).Left + cmbField(n).Width
            End If
            
            cmbField(n).Visible = True
            
            blnSameLine = Not (blnSameLine)
            
        End If
        
        ' Set currently selected
        cmbField(n).ListIndex = intIndexOrder(n) + 1
        
    Next n
    
    blnLoaded = True
    
    cmdOk.Top = cmbField(iFieldMAX).Height * (n / 2) + 720 + 220
    cmdCancel.Top = cmbField(iFieldMAX).Height * (n / 2) + 720 + 220
    
    Me.Height = cmbField(iFieldMAX).Height * (n / 2) + 720 + 640 + cmdCancel.Height
    
    ' Set available presets
    cmbPresets.AddItem "<Set Preset>"
    cmbPresets.AddItem "All Info"
    cmbPresets.AddItem "Personal Info"
    cmbPresets.AddItem "Contact Info"
    cmbPresets.AddItem "Current Education"
    cmbPresets.AddItem "Past Education"
    
    cmbPresets.ListIndex = 0
    
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    blnLoaded = False
End Sub
