Attribute VB_Name = "modMain"
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


' Record Data
Public Type Person_T                ' Represents a single record
    ID          As Long
    FirstName   As String * 10
    SurName     As String * 10
    DOB         As String * 8
    Gender      As String * 1
    EthOrigin   As String * 10
    MainHobby   As String * 15

    Address1    As String * 25
    Address2    As String * 15
    Address3    As String * 15
    PostCode    As String * 8
    Telephone   As String * 25
    Email       As String * 25

    Mentor      As String * 10
    Tutor       As String * 10
    Coruse1     As String * 10
    Coruse2     As String * 10
    Coruse3     As String * 10
    DateStarted As String * 8
    Active      As Boolean

    SchoolFrom   As String * 20
    GCSE_English As String * 1
    GCSE_Maths   As String * 1
    GCSE_Science As String * 2
    KeySkillsLev As Integer

End Type

Public intCurRec  As Integer
Public intTotRecs As Integer


Public Type VisFields_T
    Item        As String
    Width       As Integer
End Type

Public tFieldName(24) As VisFields_T


Public intIndexOrder(24) As Integer       ' Order of visable fields ('-1' = not visable)

Public Const iFieldMAX = 24

Public strCurFile As String



Sub InitFieldNames()
    
    Dim n As Integer
    
    ' Set field names and widths...
    
    tFieldName(0).Item = "ID"
    tFieldName(0).Width = 800
    
    tFieldName(1).Item = "First Name"
    tFieldName(1).Width = 1700
    
    tFieldName(2).Item = "Surname"
    tFieldName(2).Width = 1700
    
    tFieldName(3).Item = "Date of Birth"
    tFieldName(3).Width = 800
    
    tFieldName(4).Item = "Gender"
    tFieldName(4).Width = 800
    
    tFieldName(5).Item = "Ethinic Origin"
    tFieldName(5).Width = 1700
    
    tFieldName(6).Item = "Main Hobby"
    tFieldName(6).Width = 1700
    
    tFieldName(7).Item = "Address 1"
    tFieldName(7).Width = 1700
    
    tFieldName(8).Item = "Address 2"
    tFieldName(8).Width = 1700
    
    tFieldName(9).Item = "Address 3"
    tFieldName(9).Width = 1700
    
    tFieldName(10).Item = "Post Code"
    tFieldName(10).Width = 1700
    
    tFieldName(11).Item = "Telephone"
    tFieldName(11).Width = 1700
    
    tFieldName(12).Item = "Email"
    tFieldName(12).Width = 1700
    
    tFieldName(13).Item = "Mentor"
    tFieldName(13).Width = 1700
    
    tFieldName(14).Item = "Tutor"
    tFieldName(14).Width = 1700
    
    tFieldName(15).Item = "Coruse 1"
    tFieldName(15).Width = 1700
    
    tFieldName(16).Item = "Coruse 2"
    tFieldName(16).Width = 1700
    
    tFieldName(17).Item = "Coruse 3"
    tFieldName(17).Width = 1700
    
    tFieldName(18).Item = "Date Started"
    tFieldName(18).Width = 1700
    
    tFieldName(19).Item = "Active"
    tFieldName(19).Width = 800
    
    tFieldName(20).Item = "School From"
    tFieldName(20).Width = 1700
    
    tFieldName(21).Item = "GCSE English"
    tFieldName(21).Width = 2800
    
    tFieldName(22).Item = "GCSE Maths"
    tFieldName(22).Width = 2800
    
    tFieldName(23).Item = "GCSE Science"
    tFieldName(23).Width = 2800
    
    tFieldName(24).Item = "KeySkills Level"
    tFieldName(24).Width = 2800
    
    
    ' Default order of fields
    For n = 0 To iFieldMAX
        intIndexOrder(n) = n
    Next n
    
End Sub
