VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D without any API calls  :  Daniel Downes(UK)  -  Ex-D Software Development(TM)"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8340
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4500
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   4470
      ScaleWidth      =   4425
      TabIndex        =   6
      Top             =   165
      Width           =   4455
      Begin VB.Line lin2D 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Index           =   0
         X1              =   240
         X2              =   1080
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Frame frmScale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scale"
      Height          =   1455
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   3240
      Width           =   3495
      Begin VB.CheckBox chkSizeLock 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lock"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1125
         Width           =   735
      End
      Begin VB.HScrollBar sclSize 
         Height          =   255
         Index           =   0
         LargeChange     =   50
         Left            =   120
         Max             =   200
         TabIndex        =   5
         Top             =   360
         Value           =   28
         Width           =   3255
      End
      Begin VB.HScrollBar sclSize 
         Height          =   255
         Index           =   1
         LargeChange     =   50
         Left            =   120
         Max             =   200
         TabIndex        =   4
         Top             =   600
         Value           =   20
         Width           =   3255
      End
      Begin VB.HScrollBar sclSize 
         Height          =   255
         Index           =   2
         LargeChange     =   50
         Left            =   120
         Max             =   200
         TabIndex        =   3
         Top             =   840
         Value           =   20
         Width           =   3255
      End
   End
   Begin VB.Frame frmTranslation 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Translation / Position"
      Height          =   1455
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdFlyIn 
         Caption         =   "Fly In"
         Height          =   255
         Left            =   2700
         TabIndex        =   16
         Top             =   1140
         Width           =   615
      End
      Begin VB.Timer timFlyIn 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2760
         Top             =   0
      End
      Begin VB.CommandButton cmdResetTranslation 
         Caption         =   "Reset"
         Height          =   255
         Left            =   375
         TabIndex        =   15
         Top             =   1140
         Width           =   720
      End
      Begin VB.HScrollBar sclTranslate 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   120
         Max             =   600
         Min             =   -600
         SmallChange     =   10
         TabIndex        =   14
         Top             =   840
         Width           =   3255
      End
      Begin VB.HScrollBar sclTranslate 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   120
         Max             =   300
         Min             =   -300
         SmallChange     =   10
         TabIndex        =   13
         Top             =   600
         Width           =   3255
      End
      Begin VB.HScrollBar sclTranslate 
         Height          =   255
         Index           =   0
         LargeChange     =   10
         Left            =   120
         Max             =   300
         Min             =   -300
         SmallChange     =   10
         TabIndex        =   12
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame frmCamera 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rotaion"
      Height          =   1455
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
      Begin VB.CommandButton cmdStopRotation 
         Caption         =   "Stop"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.HScrollBar sclRotate 
         Height          =   255
         Index           =   1
         LargeChange     =   5
         Left            =   120
         Max             =   30
         Min             =   -30
         TabIndex        =   9
         Top             =   720
         Value           =   -1
         Width           =   3255
      End
      Begin VB.HScrollBar sclRotate 
         Height          =   255
         Index           =   0
         LargeChange     =   5
         Left            =   120
         Max             =   30
         Min             =   -30
         TabIndex        =   8
         Top             =   360
         Value           =   1
         Width           =   3255
      End
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4500
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":EA0C
      ScaleHeight     =   4470
      ScaleWidth      =   4425
      TabIndex        =   7
      Top             =   165
      Width           =   4455
      Begin VB.Line lin2Da 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   1200
         Y1              =   240
         Y2              =   240
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Thank you for trying this code. If you have any problems or queries please
'  contact me:
'
'     exd_founder@hotmail.com      (you mat also add me to your MSN Messenger contacts)
'
'  If your browser supports CommonName (IE6 does), use:
'
'     Ex-D
'
'  OR find the HTTP link via:
'
'     http://www.v99.org.uk        (Once there, go to 'Topics', you should be able to find
'                                    the 'Ex-D' logo (and site))
'
'  ...to jump to my site, to find other software.
'
'
'   Daniel Downes(UK)  -  Ex-D Software Development(TM)



Option Explicit


Dim Cam As vec3D            ' Camera position

Dim Oblong(8) As vec3D      ' Oblong, each array element defines the position
                            '  of eight vectors/points that make up the oblong
Dim TempOb(8) As vec3D      ' Stores the default oblong; no rotation
Dim points(8) As vec2D      ' These 2D points are calculated from the 3D points
                            '  thus transformed from 3D to 2D

Dim blnBuffer As Boolean    ' Selects which set of VB lines should be shown

Dim OblongPos As vec3D      ' The position of the oblong



Private Sub Form_Load()
    
    Dim n As Integer
    
    Cam.X = 0
    Cam.Y = 0
    Cam.Z = 100
    
    ' Setup the form
    Load frmMain
    frmMain.Visible = True
    cmdResetTranslation.SetFocus
    
    ' Dynamically create the lines
    For n = 1 To 11
        Load lin2D(n)
        lin2D(n).Visible = True
        
        Load lin2Da(n)
        lin2Da(n).Visible = True
    Next n
    
    
    
    ' The trig. table
    PreCalculateTrig
    
    
    blnBuffer = False
    
    
    ' Set the size of the oblong
    sclSize_Change 0

    cmdFlyIn_Click
        
    Update
    
    
    On Error Resume Next
    
    ' Each loop represents one frame
    Do
        
        If sclRotate(0).Value <> 0 Or sclRotate(1).Value <> 0 Then
            Update
        End If
        
        DoEvents

    Loop
    
End Sub
Private Sub Update()
    
    Dim Lens As Single
    Dim n As Integer
    Dim vecTemp As vec3D
    
    
    ' Update the camera position...
    Cam.X = Cam.X + sclRotate(0).Value
    If Cam.X >= 360 Then Cam.X = Cam.X - 360
    If Cam.X < 0 Then Cam.X = Cam.X + 360
        
    Cam.Y = Cam.Y + sclRotate(1).Value
    If Cam.Y >= 360 Then Cam.Y = Cam.Y - 360
    If Cam.Y < 0 Then Cam.Y = Cam.Y + 360

    
    ' Update form's lines...
    ' First calculate the new 2D points
    For n = 1 To 8
    
        ' Reset the oblong
        Oblong(n) = TempOb(n)

        
        ' Rotate Oblong vectors
        RotateVector Oblong(n), Cam.Y, Cam.X, Cam.Z
        
        ' Now add any translation
        TranslateVector Oblong(n), OblongPos.X, OblongPos.Y, OblongPos.Z
        
        ' Now the formual can be applied to transform
        '  the 3D data to a 2D perspective representation
        On Error Resume Next
        Lens = 5000 / (Oblong(n).Z - Cam.Z)
        points(n).X = 2000 + (Oblong(n).X * Lens)
        points(n).Y = 2000 - (Oblong(n).Y * Lens)
        
    Next n
    
    ' Select while set of lines to use
    ' Two buffers = no flicker (or in this case; less flicker)
    If blnBuffer = True Then
        ' Now place them lines in the right place
        ' Foreface & Backface
        For n = 0 To 4 Step 4
            lin2D(0 + n).X1 = points(1 + n).X
            lin2D(0 + n).Y1 = points(1 + n).Y
            lin2D(0 + n).X2 = points(2 + n).X
            lin2D(0 + n).Y2 = points(2 + n).Y
            
            lin2D(1 + n).X1 = points(2 + n).X
            lin2D(1 + n).Y1 = points(2 + n).Y
            lin2D(1 + n).X2 = points(3 + n).X
            lin2D(1 + n).Y2 = points(3 + n).Y
            
            lin2D(2 + n).X1 = points(3 + n).X
            lin2D(2 + n).Y1 = points(3 + n).Y
            lin2D(2 + n).X2 = points(4 + n).X
            lin2D(2 + n).Y2 = points(4 + n).Y
        
            lin2D(3 + n).X1 = points(4 + n).X
            lin2D(3 + n).Y1 = points(4 + n).Y
            lin2D(3 + n).X2 = points(1 + n).X
            lin2D(3 + n).Y2 = points(1 + n).Y
        Next n
        
        
        
        ' Side lines
        lin2D(8).X1 = points(2).X
        lin2D(8).Y1 = points(2).Y
        lin2D(8).X2 = points(6).X
        lin2D(8).Y2 = points(6).Y

        lin2D(9).X1 = points(3).X
        lin2D(9).Y1 = points(3).Y
        lin2D(9).X2 = points(7).X
        lin2D(9).Y2 = points(7).Y

        lin2D(10).X1 = points(4).X
        lin2D(10).Y1 = points(4).Y
        lin2D(10).X2 = points(8).X
        lin2D(10).Y2 = points(8).Y

        lin2D(11).X1 = points(1).X
        lin2D(11).Y1 = points(1).Y
        lin2D(11).X2 = points(5).X
        lin2D(11).Y2 = points(5).Y
        
        picBuffer(0).ZOrder 0
        picBuffer(1).ZOrder 1
        
        picBuffer(1).Refresh
    

    Else

        ' Now place them lines in the right place
        ' Foreface & Backface
        For n = 0 To 4 Step 4
            lin2Da(0 + n).X1 = points(1 + n).X
            lin2Da(0 + n).Y1 = points(1 + n).Y
            lin2Da(0 + n).X2 = points(2 + n).X
            lin2Da(0 + n).Y2 = points(2 + n).Y
            
            lin2Da(1 + n).X1 = points(2 + n).X
            lin2Da(1 + n).Y1 = points(2 + n).Y
            lin2Da(1 + n).X2 = points(3 + n).X
            lin2Da(1 + n).Y2 = points(3 + n).Y
            
            lin2Da(2 + n).X1 = points(3 + n).X
            lin2Da(2 + n).Y1 = points(3 + n).Y
            lin2Da(2 + n).X2 = points(4 + n).X
            lin2Da(2 + n).Y2 = points(4 + n).Y
        
            lin2Da(3 + n).X1 = points(4 + n).X
            lin2Da(3 + n).Y1 = points(4 + n).Y
            lin2Da(3 + n).X2 = points(1 + n).X
            lin2Da(3 + n).Y2 = points(1 + n).Y
        Next n
        
        
        ' Side lines
        lin2Da(8).X1 = points(2).X
        lin2Da(8).Y1 = points(2).Y
        lin2Da(8).X2 = points(6).X
        lin2Da(8).Y2 = points(6).Y

        lin2Da(9).X1 = points(3).X
        lin2Da(9).Y1 = points(3).Y
        lin2Da(9).X2 = points(7).X
        lin2Da(9).Y2 = points(7).Y

        lin2Da(10).X1 = points(4).X
        lin2Da(10).Y1 = points(4).Y
        lin2Da(10).X2 = points(8).X
        lin2Da(10).Y2 = points(8).Y

        lin2Da(11).X1 = points(1).X
        lin2Da(11).Y1 = points(1).Y
        lin2Da(11).X2 = points(5).X
        lin2Da(11).Y2 = points(5).Y
        
        picBuffer(1).ZOrder 0
        picBuffer(0).ZOrder 1
        
        picBuffer(0).Refresh
    
    End If
    
    frmMain.Refresh

    
    blnBuffer = Not (blnBuffer)

End Sub

Private Sub TranslateVector(vecToTranslate As vec3D, intX As Integer, intY As Integer, intZ)
    vecToTranslate.X = vecToTranslate.X + intX
    vecToTranslate.Y = vecToTranslate.Y + intY
    vecToTranslate.Z = vecToTranslate.Z + intZ
End Sub

Private Sub RotateVector(vecToRotate As vec3D, intAngX As Integer, intAngY As Integer, intAngZ)
    Dim vecTemp As vec3D
    
    ' Y axis rotation
'    vecTemp.X = (vecToRotate.Z * Sine(intAngY)) + (vecToRotate.X * Cosine(intAngY))
'    vecToRotate.Y = vecTemp.Y
'    vecToRotate.Z = (vecToRotate.Z * Cosine(intAngY)) - (vecToRotate.X * Sine(intAngY))
'
'    ' X axis rotation
'    vecToRotate.X = vecTemp.X
'    vecTemp.Y = (vecToRotate.Y * Cosine(intAngX)) - (vecToRotate.Z * Sine(intAngX))
'    vecToRotate.X = (vecToRotate.Y * Sine(intAngX)) + (vecToRotate.Z * Cosine(intAngX))
'
'    ' Z axis rotation
'    vecTemp.X = (vecToRotate.X * Sine(intAngZ)) + (vecToRotate.Y * Cosine(intAngZ))
'    vecToRotate.Y = (vecToRotate.X * Cosine(intAngZ)) - (vecToRotate.Y * Sine(intAngZ))
'    vecToRotate.Z = vecTemp.Z

    ' Accourding to a book, titled '3D Graphics Programming' the above should work
    '  but does not! Contact me if you know why, and/or can fix it.
    
    
    ' The rotation fomuale
    ' X rotation
    vecTemp.Y = (vecToRotate.Y * Cosine(intAngX)) - (vecToRotate.Z * Sine(intAngX))
    vecToRotate.Z = (vecToRotate.Z * Cosine(intAngX)) + (vecToRotate.Y * Sine(intAngX))
    vecToRotate.Y = vecTemp.Y

    ' Y rotation
    vecTemp.Z = (vecToRotate.Z * Cosine(intAngY)) - (vecToRotate.X * Sine(intAngY))
    vecToRotate.X = (vecToRotate.X * Cosine(intAngY)) + (vecToRotate.Z * Sine(intAngY))
    vecToRotate.Z = vecTemp.Z


    
End Sub


Private Sub SetOblongData(vecSize As vec3D)

    ' Foreface...
    Oblong(1).X = -vecSize.X      ' Top-left
    Oblong(1).Y = vecSize.Y
    Oblong(1).Z = vecSize.Z
    
    Oblong(2).X = vecSize.X     ' Top-right
    Oblong(2).Y = vecSize.Y
    Oblong(2).Z = vecSize.Z
    
    Oblong(3).X = vecSize.X     ' Bottom-right
    Oblong(3).Y = -vecSize.Y
    Oblong(3).Z = vecSize.Z
    
    Oblong(4).X = -vecSize.X      ' Bottom-left
    Oblong(4).Y = -vecSize.Y
    Oblong(4).Z = vecSize.Z
   
   
    ' Backface...
    Oblong(5).X = -vecSize.X      ' Top-left
    Oblong(5).Y = vecSize.Y
    Oblong(5).Z = -vecSize.Z
    
    Oblong(6).X = vecSize.X      ' Top-right
    Oblong(6).Y = vecSize.Y
    Oblong(6).Z = -vecSize.Z
    
    Oblong(7).X = vecSize.X      ' Bottom-right
    Oblong(7).Y = -vecSize.Y
    Oblong(7).Z = -vecSize.Z
    
    Oblong(8).X = -vecSize.X      ' Bottom-left
    Oblong(8).Y = -vecSize.Y
    Oblong(8).Z = -vecSize.Z
    
End Sub


Sub PreCalculateTrig()
    
    Dim n As Integer
    
    For n = 0 To 361
        Sine(n) = Sin(n / 180 * PI)
        Cosine(n) = Cos(n / 180 * PI)
    Next n
    
End Sub




' Form events...

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub sclTranslate_Change(Index As Integer)
    OblongPos.X = sclTranslate(0).Value
    OblongPos.Y = sclTranslate(1).Value
    OblongPos.Z = sclTranslate(2).Value
    
    Update
End Sub

Private Sub cmdResetTranslation_Click()
    sclTranslate(0) = 0
    sclTranslate(1) = 0
    sclTranslate(2) = 0
    
    Update
End Sub


Private Sub cmdStopRotation_Click()
    sclRotate(0).Value = 0
    sclRotate(1).Value = 0
End Sub


Private Sub sclSize_Change(Index As Integer)
    
    Dim vecTemp As vec3D
    Dim n As Integer
    
    If chkSizeLock.Value = 1 Then
        sclSize(0).Value = sclSize(Index).Value
        vecTemp.X = sclSize(0).Value
        vecTemp.Y = sclSize(0).Value
        vecTemp.Z = sclSize(0).Value
        
        sclSize(1).Value = sclSize(0).Value
        sclSize(2).Value = sclSize(0).Value
    Else
        vecTemp.X = sclSize(0).Value
        vecTemp.Y = sclSize(1).Value
        vecTemp.Z = sclSize(2).Value
    End If
        
    SetOblongData vecTemp
    
    For n = 1 To 8
        
        ' Set the default oblong
        TempOb(n) = Oblong(n)
        
    Next n
    
    Update
    
End Sub
Private Sub chkSizeLock_Click()
    
    Dim sngMean As Single
    
    If chkSizeLock.Value = 1 Then
        ' Find the mean of the three scrollbar
        '  and set them all to it
        
        sngMean = (sclSize(0).Value + sclSize(1).Value + sclSize(2).Value) / 3
        
        sclSize(0).Value = Int(sngMean)
        sclSize(1).Value = Int(sngMean)
        sclSize(2).Value = Int(sngMean)
        
    End If
    
    Update
End Sub


Private Sub cmdFlyIn_Click()
    sclTranslate(0).Value = sclTranslate(0).Min + 115
    sclTranslate(1).Value = sclTranslate(1).Min + 120
    
    sclTranslate(2).Value = sclTranslate(2).Min
    timFlyIn.Enabled = True
End Sub

Private Sub timFlyIn_Timer()
    sclTranslate(0).Value = sclTranslate(0).Value + 3
    sclTranslate(1).Value = sclTranslate(1).Value + 3
    sclTranslate(2).Value = sclTranslate(2).Value + 10
    
    If sclTranslate(2).Value > sclTranslate(2).Max + sclTranslate(2).Min Then
        timFlyIn.Enabled = False
    End If
End Sub
