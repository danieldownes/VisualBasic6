Attribute VB_Name = "modOLicenceReg"
Public Declare Function OLPromptLicense Lib "OLCHK32.DLL" (ByVal hParent As Long, lpID As String) As Boolean
Public Declare Function OLStoreLicense Lib "OLCHK32.DLL" (ByVal hParent As Long, lpID As String, lpLicense As String) As Boolean
Public Declare Function OLRetrieveKey Lib "OLCHK32.DLL" (ByVal hParent As Long, lpID As String, lpPassword As String, lpKey As String, nKeySize As Integer) As String

Function OLCheckKey(lngHandel As Long) As Boolean
    Dim strExKeyBuild As String
    Dim strExPassBuild As String
    Dim strReturnedKey As String


    OLCheckKey = False
    
    strKeyBuild = "EX27HHB102"                          ' Key
    strExPassBuild = "exhhbmpanlc"                      ' Password
    
    
    strReturnedKey = OLRetrieveKey(lngHandel, "HiddenHistory", strExPassBuild, strExKeyBuild, Len(strExKeyBuild))

    If strKeyBuild = strReturnedKey Then
        OLCheckKey = True
    End If
    
End Function



