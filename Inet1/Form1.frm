VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser Inet1 
      Height          =   2055
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      ExtentX         =   6588
      ExtentY         =   3625
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub download_url(sURL As String)
Dim fn As Integer
Dim sfile As String
Dim sSource As String
Dim c() As Byte

On Error GoTo ERROR_GetPage

' *** Stop anything going on
Inet1.cancal
' *** Set protocol to HTTP
Inet1.Protocol = icHTTP
' *** Set the URL Property
Inet1.URL = sURL
' *** Retrieve the HTML data into a byte array.
c() = Inet1.OpenURL(, icByteArray)

' it get's saved here.
' change the "sfile" to what ever.
fn = FreeFile
sfile = App.Path + "\URL.htm"
Open sfile For Binary Access Write As fn
Put #fn, , c()
Close fn
EXIT_GetPage:
   Exit Sub
ERROR_GetPage:
   Exit Sub
End Sub

Private Sub Form_Load()
    download_url ("http://www.ex-d.ic24.net/index.html")
End Sub
