VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' API Constants
Private Const ERROR_SUCCESS     As Long = 0
Private Const WS_VERSION_REQD   As Long = &H101
Private Const MIN_SOCKETS_REQD  As Long = 1
Private Const DATA_SIZE = 32
Private Const MAX_WSAD = 256
Private Const MAX_WSAS = 128
Private Const PING_TIMEOUT = 255
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const REG_SZ = 1&

' Winsock API Type defs...
Private Type ICMP_OPTIONS
    Ttl                         As Byte
    Tos                         As Byte
    Flags                       As Byte
    OptionsSize                 As Byte
    OptionsData                 As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address                     As Long
    Status                      As Long
    RoundTripTime               As Long
    DataSize                    As Long
    DataPointer                 As Long
    Options                     As ICMP_OPTIONS
    Data                        As String * 250
End Type

Private Type HOSTENT
    hName                       As Long
    hAliases                    As Long
    hAddrType                   As Integer
    hLen                        As Integer
    hAddrList                   As Long
End Type

Private Type WSADATA
    wVersion                    As Integer
    wHighVersion                As Integer
    szDescription(MAX_WSAD)     As Byte
    szSystemStatus(MAX_WSAS)    As Byte
    wMaxSockets                 As Integer
    wMaxUDPDG                   As Integer
    dwVendorInfo                As Long
End Type

' SystemTime and TimeZone API Type defs...
Private Type SYSTEMTIME
    wYear                       As Integer
    wMonth                      As Integer
    wDayOfWeek                  As Integer
    wDay                        As Integer
    wHour                       As Integer
    wMinute                     As Integer
    wSecond                     As Integer
    wMilliseconds               As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias                        As Long
    StandardName(32)            As Integer
    StandardDate                As SYSTEMTIME
    StandardBias                As Long
    DaylightName(32)            As Integer
    DaylightDate                As SYSTEMTIME
    DaylightBias                As Long
End Type

' Class Enum for host name string validation
Public Enum VALIDATE_HOST_METHOD
    VALIDATE_HOST_NONE = 0
    VALIDATE_HOST_SYNTAX = 1
    VALIDATE_HOST_PING = 2
    VALIDATE_HOST_DNS = 3
End Enum

' Class Enum for email address string validation
Public Enum VALIDATE_METHOD
    VALIDATE_NONE = 0
    VALIDATE_SYNTAX = 1
End Enum

' Class Enum for email encoding method
Public Enum ENCODE_METHOD
    MIME_ENCODE = 0
    UU_ENCODE = 1
End Enum

' Class Enum for mail priority
Public Enum MAIL_PRIORITY
    HIGH_PRIORITY = 1
    NORMAL_PRIORITY = 3
    LOW_PRIORITY = 5
End Enum

' Structure to hold mail elements
Private Type MAIL_DATA
    sToAddr()                   As String           ' To: email address
    sToDisplayName()            As String           ' To: display name
    sCcAddr()                   As String           ' Cc: email address
    sCcDisplayName()            As String           ' Cc: display name
    sBccAddr()                  As String           ' Bcc: email address
    sFromAddr                   As String           ' From: email address
    sFromDisplayName            As String           ' From: display name
    sReplyToAddr                As String           ' ReplyTo: email address
    sSubject                    As String           ' Subject
    sMailMessage                As String           ' email message body
    sAttachment()               As String           ' attachment path\filename
    sAttachNameOnly()           As String           ' attachment name only
    lAttachNameSize             As Long             ' sum of the lenght of all attachment names
    lAttachFileSize             As Long             ' sum of all file lenghts
    lAttachCount                As Long             ' number of attachments
End Type

' Class Property var's
Private utMail                  As MAIL_DATA        ' see above type def
Private etPriority              As MAIL_PRIORITY    ' mail priority, Normal - High - Low
Private psDelimiter             As String           ' string to delimit multiple entries
Private psSMTPHost              As String           ' remote host name or IP number
Private plSMTPPort              As Long             ' remote host port number
Private pbUseAuthentication     As Boolean          ' flag, use login authentication with host
Private pbHtmlText              As Boolean          ' flag, send plain text / html text
Private psContentBase           As String           ' Content base for HTML text
Private plConnectTimeout        As Long             ' timeout value for connection attempts
Private plConnectRetry          As Long             ' number of times to attempt a connection
Private plMessageTimeOut        As Long             ' timeout value for sending a message
Private plPersistentSettings    As Long             ' flag, persistent/non-persistent settings
Private etEncodeType            As ENCODE_METHOD    ' MIME / UUEncode flag
Private etEmailValidation       As VALIDATE_METHOD  ' type of email address validation to use
Private etSMTPHostValidation    As VALIDATE_METHOD  ' type of Host validation to use
Private pbReceipt               As Boolean          ' flag, request a return receipt

' Class local var's
Private psTimeZoneBias          As String           ' time zone offset bias
Private pColErrors              As Collection       ' errors collection
Private pbBase64Byt(0 To 63)    As Byte             ' base 64 encoder byte array
Private psUUEncodeChr(0 To 63)  As String           ' UU encoder string array
Private pb8BitMail              As Boolean          ' flag, 7/8 bit message body
Private pbExitImmediately       As Boolean          ' flag - unrecoverable error
Private pbConnected             As Boolean          ' flag, connection to host established
Private pbManualDisconnect      As Boolean          ' flag, stay connected until 'Disconnect' called
Private pbRequestAccepted       As Boolean          ' flag, host accepted request
Private pbDataOK                As Boolean          ' flag, received "OK" from host
Private pbAuthLoginSupported    As Boolean          ' flag, host supports auth login
Private pbAuthMailFromOK        As Boolean          ' flag, host accepts 'mail from' auth
Private pbAuthLoginSuccess      As Boolean          ' flag, Auth login accepted by remote host
Private plBytesSent             As Long             ' running total of bytes sent
Private plBytesRemaining        As Long             ' bytes remaining to be sent in sock send buffer
Private pbSending               As Boolean
Private plMailSize              As Long             ' total size of email session
Private psUserName              As String           ' Auth username - optional, not supported by all servers
Private psPassword              As String           ' Auth password - optional, not supported by all servers
Private psPriority              As String           ' string version of priority Property for MSMail
Private WithEvents sckMail      As Winsock          ' project must include the Winsock control
Attribute sckMail.VB_VarHelpID = -1
                                                    ' or a reference to the mswinsck.ocx
Private sDay()                  As String           ' day name array
Private sMonth()                As String           ' month name array

' Class Constants

' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

' error strings used with 'pColErrors' collection to report errors to the user
Private Const ERR_INVALID_HOST = "Invalid Host Name"
Private Const ERR_INVALID_PORT = "Invalid Remote Port"
Private Const ERR_INVALID_REC_EMAIL = "Missing or Invalid Recipient E-mail Address"
Private Const ERR_NO_REC_EMAIL = "No Recipient E-mail Address Specified"
Private Const ERR_INVALID_CC_EMAIL = "Invalid Cc: Recipient E-mail Address"
Private Const ERR_INVALID_BCC_EMAIL = "Invalid Bcc: Recipient E-mail Address"
Private Const ERR_INVALID_SND_EMAIL = "Missing or Invalid Sender E-mail Address"
Private Const ERR_TIMEOUT = "Timeout occurred: The SMTP Host did not respond to the request"
Private Const ERR_FILE_NOT_EXIST = "The file you tried to attach does not exist"
Private Const ERR_RECIPIENT_COUNT = "Too many recipients"
Private Const ERR_HTML_REQUIRES_MIME = "Sending HTML requires MIME encoding"

' misc startup defaults
Private Const CONNECT_TIMEOUT = 10                  ' seconds to wait before giving up
Private Const CONNECT_RETRY = 5                     ' number of times to try before giving up
Private Const MSG_TIMEOUT = 60                      ' seconds before timing out on message transmission
Private Const REG_KEY = "vbSendMail"                ' registry key
Private Const SETTINGS_KEY = "Settings"             ' registry sub key
Private Const DEFAULT_PORT As Long = 25             ' default socket port for SMTP

' maximums per RFC 821...
Private Const MAX_TEXTLINE_LEN = 1000               ' maximum total lenght of a text line
Private Const MAX_RECIPIENTS = 100                  ' maximum number of recipients that must be buffered

' list of top level Domains, obtained from www.IANA.com.
' Can and will change, used in host name syntax checking
Private Const TOP_DOMAINS = "COM ORG NET EDU GOV MIL INT AF AL DZ AS " & _
        "AD AO AI AQ AG AR AM AW AC AU AT AZ BS BH BD BB BY BZ BT BJ " & _
        "BE BM BO BA BW BV BR IO BN BG BF BI KH CM CA CV KY CF TD CL " & _
        "CN CX CC CO KM CD CG CK CR CI HR CU CY CZ DK DJ DM DO TP EC " & _
        "EG SV GQ ER EE ET FK FO FJ FI FR GF PF TF GA GM GE DE GH GI " & _
        "GR GL GD GP GU GT GG GN GW GY HT HM VA HN HK HU IS IN ID IR " & _
        "IQ IE IM IL IT JM JP JE JO KZ KE KI KP KR KW KG LA LV LB LS " & _
        "LR LY LI LT LU MO MK MG MW MY MV ML MT MH MQ MR MU YT MX FM " & _
        "MD MC MN MS MA MZ MM NA NR NP NL AN NC NZ NI NE NG NU NF MP " & _
        "NO OM PK PW PA PG PY PE PH PN PL PT PR QA RE RO RU RW KN LC " & _
        "VC WS SM ST SA SN SC SL SG SK SI SB SO ZA GS ES LK SH PM SD " & _
        "SR SJ SZ SE CH SY TW TJ TZ TH TG TK TO TT TN TR TM TC TV UG " & _
        "UA AE GB US UM UY UZ VU VE VN VG VI WF EH YE YU ZR ZM ZW UK"

' Class Events
Public Event SendSuccesful()
Public Event SendFailed(Explanation As String)
Public Event Status(Status As String)
Public Event Progress(PercentComplete As Long)

' API prototypes...
' winsock
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" _
        (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" _
        (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
        (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, _
        ByVal RequestData As String, ByVal RequestSize As Long, _
        ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, _
        ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
' registry
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Long) As Long
' misc
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
        (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Class_Initialize()

    Dim iPtr        As Integer                      ' loop counter
    Dim utTZ        As TIME_ZONE_INFORMATION        ' api time zone type

    
    ' instantiate the Error collection
    Set pColErrors = New Collection
    
    ' instantiate the Winsock Control
    Set sckMail = frmSck.WinSck
    
    ' alternate method of instantiating Winsock without a Form.
    ' use a project Reference instead of the included frmSck & Winsock control
    ' *** currently has unresolved deployment issues ***
    'Set sckMail = New Winsock
    
    ' initialize default values...
    plPersistentSettings = CLng(RegGet("PersistentSettings", "0"))
    If plPersistentSettings Then
        ' load defaults from the registry
        utMail.sFromAddr = RegGet("From", "")
        utMail.sFromDisplayName = RegGet("FromDisplayName", "")
        psSMTPHost = RegGet("RemoteHost", "")
        plSMTPPort = CLng(RegGet("RemotePort", DEFAULT_PORT))
        etSMTPHostValidation = RegGet("SMTPHostValidation", VALIDATE_HOST_DNS)
        etEmailValidation = CLng(RegGet("EmailValidation", VALIDATE_SYNTAX))
        plConnectTimeout = CLng(RegGet("ConnectTimeout", CONNECT_TIMEOUT))
        plMessageTimeOut = CLng(RegGet("MessageTimeout", MSG_TIMEOUT))
        plConnectRetry = CLng(RegGet("ConnectRetry", CONNECT_RETRY))
        etEncodeType = RegGet("EncodeType", MIME_ENCODE)
        psUserName = RegGet("Username", "")
        pbUseAuthentication = RegGet("UseAuthentication", False)
    Else
        ' load standard defaults
        plSMTPPort = DEFAULT_PORT
        etSMTPHostValidation = VALIDATE_HOST_DNS
        etEmailValidation = VALIDATE_SYNTAX
        plConnectTimeout = CONNECT_TIMEOUT
        plMessageTimeOut = MSG_TIMEOUT
        plConnectRetry = CONNECT_RETRY
        etEncodeType = MIME_ENCODE
    End If

    ' initialize the arrays for base64 & uu encoders
    For iPtr = 0 To 63
        pbBase64Byt(iPtr) = Asc(Mid$(BASE64CHR, iPtr + 1, 1))
        psUUEncodeChr(iPtr) = Chr$(iPtr + &H20)
    Next
    psUUEncodeChr(0) = Chr$(&H60)

    ' calculate the time zone offset bias
    GetTimeZoneInformation utTZ
    psTimeZoneBias = Format$(-utTZ.Bias \ 60, "00") & Format$(Abs(utTZ.Bias - (utTZ.Bias \ 60) * 60), "00")
    If InStr(psTimeZoneBias, "-") = 0 Then psTimeZoneBias = "+" & psTimeZoneBias
    
    
    ' init mail recipient arrays (sets Ubound to -1)
    utMail.sToAddr = Split("")
    utMail.sToDisplayName = utMail.sToAddr
    utMail.sCcAddr = utMail.sToAddr
    utMail.sCcDisplayName = utMail.sToAddr
    utMail.sBccAddr = utMail.sToAddr
    utMail.sAttachment = utMail.sToAddr

    ' set default delimiter
    psDelimiter = ";"

    ' set default priority
    etPriority = NORMAL_PRIORITY

    ' initialize the day/month arrays needed to support non-English systems.
    ' some email clients/servers will not accept non-English words in the
    ' date field so we need to guarantee that the day & month are English.
    ' These arrays are used in the Send Sub to format the current time/date.
    sDay() = Split(",Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")
    sMonth() = Split(",Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec", ",")


End Sub

Private Sub Class_Terminate()

    ' make sure sckMail is closed
    If sckMail.State <> sckClosed Then
        DisconnectFromHost
    End If
    
    ' release memory
    Set sckMail = Nothing
    Set pColErrors = Nothing

End Sub

' ******************************************************************************
' *      Class Properties                                                      *
' ******************************************************************************

Public Property Get AsHTML() As Boolean

    ' return the Property value
    AsHTML = pbHtmlText

End Property

Public Property Let AsHTML(ByVal NewValue As Boolean)

    ' save the new Property value
    pbHtmlText = NewValue

End Property

Public Property Get Attachment() As String

    ' return the Property value
    Attachment = Join(utMail.sAttachment, psDelimiter)

End Property

Public Property Let Attachment(ByVal NewValue As String)

    Dim sNameOnly()     As String
    Dim lPtr            As Long

    ' save the new Property value
    utMail.sAttachment = Split(NewValue, psDelimiter)
    ' set the 'name only' array to match
    utMail.sAttachNameOnly = utMail.sAttachment
    ' reset the counters
    utMail.lAttachCount = UBound(utMail.sAttachment) + 1
    utMail.lAttachFileSize = 0
    utMail.lAttachNameSize = 0

    RemoveError ERR_FILE_NOT_EXIST

    ' verify each entry...
    For lPtr = 0 To UBound(utMail.sAttachment)
        ' check that the file exists
        If Dir(utMail.sAttachment(lPtr)) = "" Then
            AddError ERR_FILE_NOT_EXIST
            Exit For
        End If
        ' extract the file name
        sNameOnly = Split(utMail.sAttachment(lPtr), "\")
        utMail.sAttachNameOnly(lPtr) = sNameOnly(UBound(sNameOnly))
        ' add up the file sizes and name lengths for later...
        utMail.lAttachFileSize = utMail.lAttachFileSize + FileLen(utMail.sAttachment(lPtr))
        utMail.lAttachNameSize = utMail.lAttachNameSize + Len(utMail.sAttachNameOnly(lPtr))
    Next

End Property

Public Property Get BccRecipient() As String

    ' return the Property value
    BccRecipient = Join(utMail.sBccAddr, psDelimiter)

End Property

Public Property Let BccRecipient(ByVal NewValue As String)

    ' save the new Property value
    utMail.sBccAddr = Split(NewValue, psDelimiter)
    ValidateAddress NewValue, ERR_INVALID_BCC_EMAIL

End Property

Public Property Get CcDisplayName() As String

    ' return the Property value
    CcDisplayName = Join(utMail.sCcDisplayName, psDelimiter)

End Property

Public Property Let CcDisplayName(ByVal NewValue As String)

    ' save the new Property value
    utMail.sCcDisplayName = Split(NewValue, psDelimiter)

End Property

Public Property Get CcRecipient() As String

    ' return the Property value
    CcRecipient = Join(utMail.sCcAddr, psDelimiter)

End Property

Public Property Let CcRecipient(ByVal NewValue As String)

    ' save the new Property value
    utMail.sCcAddr = Split(NewValue, psDelimiter)
    ValidateAddress NewValue, ERR_INVALID_CC_EMAIL

End Property

Public Property Get ConnectRetry() As Long

    ' return the Property value
    ConnectRetry = plConnectRetry

End Property

Public Property Let ConnectRetry(ByVal NewValue As Long)

    ' save the new Property value
    If NewValue > 0 And NewValue <= 20 Then plConnectRetry = NewValue
    RegSave "ConnectRetry", Str$(NewValue)

End Property

Public Property Get ConnectTimeout() As Long

    ' return the Property value
    ConnectTimeout = plConnectTimeout

End Property

Public Property Let ConnectTimeout(ByVal NewValue As Long)

    ' save the new Property value
    If NewValue > 0 And NewValue <= 120 Then plConnectTimeout = NewValue
    RegSave "ConnectTimeout", Str$(NewValue)

End Property

Public Property Get ContentBase() As String

    ' return the Property value
    ContentBase = psContentBase

End Property

Public Property Let ContentBase(ByVal NewValue As String)

    ' save the new Property value
    ' fix some common mistakes...
    If Len(NewValue) Then
        Replace$ NewValue, "\", "/"
        If InStr(1, NewValue, "http://", vbTextCompare) = 0 Then NewValue = "http://" & NewValue
        If Right$(NewValue, 1) <> "/" Then NewValue = NewValue & "/"
    End If
    psContentBase = NewValue

End Property

Public Property Get Delimiter() As String

    ' return the Property value
    Delimiter = psDelimiter

End Property

Public Property Let Delimiter(ByVal NewValue As String)

    ' save the new Property value
    psDelimiter = Left$(NewValue, 1)

End Property

Public Property Get EmailAddressValidation() As VALIDATE_METHOD

    ' return the Property value
    EmailAddressValidation = etEmailValidation

End Property

Public Property Let EmailAddressValidation(ByVal NewValue As VALIDATE_METHOD)

    ' save the new Property value
    etEmailValidation = NewValue
    RegSave "EmailValidation", Str$(NewValue)

End Property

Public Property Get EncodeType() As ENCODE_METHOD

    ' return the Property value
    EncodeType = etEncodeType

End Property

Public Property Let EncodeType(ByVal NewValue As ENCODE_METHOD)
    
    ' save the new Property value
    etEncodeType = NewValue
    RegSave "EncodeType", Str$(NewValue)

End Property

Public Property Get From() As String

    ' return the Property value
    From = utMail.sFromAddr

End Property

Public Property Let From(ByVal NewValue As String)
    
    ' save the new Property value
    utMail.sFromAddr = NewValue
    ValidateAddress NewValue, ERR_INVALID_SND_EMAIL
    RegSave "From", NewValue

End Property

Public Property Get FromDisplayName() As String

    ' return the Property value
    FromDisplayName = utMail.sFromDisplayName

End Property

Public Property Let FromDisplayName(ByVal NewValue As String)
    
    ' save the new Property value
    utMail.sFromDisplayName = NewValue
    RegSave "FromDisplayName", NewValue

End Property

Public Property Get Message() As String

    ' return the Property value
    Message = utMail.sMailMessage

End Property

Public Property Let Message(ByVal NewValue As String)

    Dim lPtr        As Long
    Dim bytTmp()    As Byte

    ' save the new Property value
    utMail.sMailMessage = FormatMail(NewValue)

    ' check for any 8 bit characters
    pb8BitMail = False
    bytTmp() = StrConv(utMail.sMailMessage, vbFromUnicode)

    For lPtr = 0 To UBound(bytTmp)
        If bytTmp(lPtr) > 127 Then
            pb8BitMail = True
            Exit For
        End If
    Next

End Property

Public Property Get MessageTimeout() As Long

    ' return the Property value
    MessageTimeout = plMessageTimeOut

End Property

Public Property Let MessageTimeout(ByVal NewValue As Long)

    ' save the new Property value
    plMessageTimeOut = Abs(NewValue)
    RegSave "MessageTimeout", Str$(NewValue)

End Property

Public Property Get Password() As String

    ' return the Property value
    Password = psPassword

End Property

Public Property Let Password(ByVal NewValue As String)

    ' save the new Property value
    psPassword = NewValue

End Property

Public Property Get PersistentSettings() As Boolean

    ' return the Property value
    PersistentSettings = plPersistentSettings

End Property

Public Property Let PersistentSettings(ByVal NewValue As Boolean)

    ' save the new Property value
    plPersistentSettings = NewValue
    RegSave "PersistentSettings", CStr(CLng(NewValue))

End Property

Public Property Get Priority() As MAIL_PRIORITY

    ' return the Property value
    Priority = etPriority

End Property

Public Property Let Priority(ByVal NewValue As MAIL_PRIORITY)

    ' save the new Property value
    etPriority = NewValue

    ' set the string version to match
    Select Case etPriority
        
        Case NORMAL_PRIORITY
            psPriority = "Normal"
        
        Case HIGH_PRIORITY
            psPriority = "High"
        
        Case LOW_PRIORITY
            psPriority = "Low"
    
    End Select

End Property

Public Property Get Receipt() As Boolean

    ' return the Property value
    Receipt = pbReceipt

End Property

Public Property Let Receipt(ByVal NewValue As Boolean)

    ' save the new Property value
    pbReceipt = NewValue

End Property

Public Property Get Recipient() As String

    ' return the Property value
    Recipient = Join(utMail.sToAddr, psDelimiter)

End Property

Public Property Let Recipient(ByVal NewValue As String)
    
    ' save the new Property value
    utMail.sToAddr = Split(NewValue, psDelimiter)
    ValidateAddress NewValue, ERR_INVALID_REC_EMAIL

End Property

Public Property Get RecipientDisplayName() As String

    ' return the Property value
    RecipientDisplayName = Join(utMail.sToDisplayName, psDelimiter)

End Property

Public Property Let RecipientDisplayName(ByVal NewValue As String)

    ' save the new Property value
    utMail.sToDisplayName = Split(NewValue, psDelimiter)

End Property

Public Property Get ReplyToAddress() As String

    ' return the Property value
    ReplyToAddress = utMail.sReplyToAddr

End Property

Public Property Let ReplyToAddress(ByVal NewValue As String)

    ' save the new Property value
    utMail.sReplyToAddr = NewValue

End Property

Public Property Get SMTPHost() As String

    ' return the Property value
    SMTPHost = psSMTPHost

End Property

Public Property Let SMTPHost(NewValue As String)

    Dim bValid      As Boolean
    
    ' validate the new host name...
    Select Case etSMTPHostValidation
        
        Case VALIDATE_HOST_SYNTAX
            bValid = IsValidIPHost(NewValue)
        
        Case VALIDATE_HOST_PING
            bValid = Ping(NewValue)
        
        Case VALIDATE_HOST_DNS
            If GetIPAddress(NewValue) <> "" Then bValid = True
        
        Case Else
            bValid = True
    
    End Select

    ' save the new Property value
    If bValid Then
        RegSave "RemoteHost", NewValue
        RemoveError ERR_INVALID_HOST
        psSMTPHost = NewValue
    Else
        AddError ERR_INVALID_HOST
    End If

End Property

Public Property Get SMTPHostValidation() As VALIDATE_HOST_METHOD

    ' return the Property value
    SMTPHostValidation = etSMTPHostValidation

End Property

Public Property Let SMTPHostValidation(ByVal NewValue As VALIDATE_HOST_METHOD)

    ' save the new Property value
    etSMTPHostValidation = NewValue
    RegSave "SMTPHostValidation", Str$(NewValue)

    ' in case this is set after the host value is set
    If psSMTPHost <> "" Then SMTPHost = psSMTPHost

End Property

Public Property Get SMTPPort() As Long

    ' return the Property value
    SMTPPort = plSMTPPort

End Property

Public Property Let SMTPPort(ByVal NewValue As Long)

    ' save the new Property value
    If NewValue < 1 Or NewValue > 65535 Then
        AddError ERR_INVALID_PORT
    Else
        plSMTPPort = NewValue
        RegSave "RemotePort", Str$(NewValue)
        RemoveError ERR_INVALID_PORT
    End If

End Property

Public Property Get Subject() As String

    ' return the Property value
    Subject = utMail.sSubject

End Property

Public Property Let Subject(ByVal NewValue As String)

    ' save the new Property value
    utMail.sSubject = NewValue

End Property

Public Property Get UseAuthentication() As Boolean

    ' return the Property value
    UseAuthentication = pbUseAuthentication

End Property

Public Property Let UseAuthentication(ByVal NewValue As Boolean)

    ' save the new Property value
    pbUseAuthentication = NewValue
    RegSave "UseAuthentication", CStr(CLng(NewValue))

End Property

Public Property Get Username() As String

    ' return the Property value
    Username = psUserName

End Property

Public Property Let Username(ByVal NewValue As String)

    ' save the new Property value
    psUserName = NewValue
    RegSave "Username", NewValue

End Property

' ******************************************************************************
' *      Class Methods                                                         *
' ******************************************************************************

Public Function Connect() As Boolean

    ' public version of ConnectToHost
    ' sets pbManualDisconnect flag so Send Sub
    ' will not disconnect when finished....
    pbManualDisconnect = True
    Connect = ConnectToHost

End Function

Public Sub Disconnect()

    ' public version of DisconnectFromHost
    ' clears pbManualDisconnect flag
    pbManualDisconnect = False
    DisconnectFromHost

End Sub

Public Function GetContentType(ByVal strFile As String, Optional strDefault As String = "application/octet-stream") As String

' ******************************************************************************
'
' Synopsis:     Get the Content Type from the Registry.
'
' Parameters:   strFile     - The filename to get the Content Type for
'               strDefault  - The default data to return if nothing is found
'
' Return:       The Content Type string
'
' Description:
' The Content Type string for registered file extensions is located in
' the system registry, in the root key HKEY_CLASSES_ROOT. Open the registry
' key for the given file extension and read the 'Content Type' value. If the
' key and/or value are not found, assign a default value of
' 'application/octet-stream'
'
' ******************************************************************************

    Dim hKey                As Long                 ' key handle
    Dim strBuff             As String               ' buffer for API to write to
    Dim lBuffLen            As Long                 ' lenght of API return string
    Dim lRet                As Long                 ' API return code
    Dim lValueType          As Long                 ' data type for retun value
    Dim iPtr                As Integer              ' scratch pointer
    Dim strValueName        As String               ' registry 'value name
    Dim strKeyName          As String               ' registry 'key name

    GetContentType = strDefault

    ' registry value name
    strValueName = "Content Type"

    ' get the passed in key name. We only want
    ' the file extension here e.g. .exe, .doc, etc.
    ' if an extension is not found, assign default
    ' value and return
    iPtr = InStrRev(strFile, ".")
    If iPtr Then
        strKeyName = Mid$(strFile, iPtr)
    Else
        Exit Function
    End If

    ' open the Registry key, if key not found, return the defaut value
    lRet = RegOpenKey(HKEY_CLASSES_ROOT, strKeyName, hKey)
    If lRet <> ERROR_SUCCESS Then Exit Function

    ' query the key value to get it's data type & length
    lRet = RegQueryValueEx(hKey, strValueName, 0&, lValueType, 0&, lBuffLen)

    ' should be type string...
    If lValueType = REG_SZ Then
        ' create a buffer & call the API again
        strBuff = String$(lBuffLen, " ")
        lRet = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal strBuff, lBuffLen)
        ' get the string value, drop the trailing '0'
        If lRet = ERROR_SUCCESS Then GetContentType = Left$(strBuff, lBuffLen - 1)
    End If

    ' close the key
    If hKey Then lRet = RegCloseKey(hKey)

End Function

Public Function GetIPAddress(sHostName As String) As String

    ' Resolves host-name to an IP address (DNS)
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim lpHost          As Long
    Dim HOST            As HOSTENT
    Dim dwIPAddr        As Long
    Dim tmpIPAddr()     As Byte
    Dim i               As Integer
    Dim sIPAddr         As String
    
    ' init winsock api
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    
    ' if no name given, use local host
    If sHostName = "" Then sHostName = GetIPHost
    sHostName = Trim$(sHostName) & Chr$(0)
    
    ' call api
    lpHost = gethostbyname(sHostName)

    If lpHost Then
        ' extract the data...
        CopyMemory HOST, ByVal lpHost, Len(HOST)
        CopyMemory dwIPAddr, ByVal HOST.hAddrList, 4
        ReDim tmpIPAddr(1 To HOST.hLen)
        CopyMemory tmpIPAddr(1), ByVal dwIPAddr, HOST.hLen

        ' convert format
        For i = 1 To HOST.hLen
            sIPAddr = sIPAddr & tmpIPAddr(i) & "."
        Next

        ' set the return value
        GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    
    Else
        WSAGetLastError
        GetIPAddress = ""
    End If
    
    SocketsCleanup

End Function

Public Function GetIPHost() As String

    ' Resolves the local host name
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim sHostName   As String
    Dim iPtr        As Integer

    ' create a buffer
    sHostName = String$(256, Chr$(0))

    ' init winsock api
    If Not SocketsInitialize() Then Exit Function

    ' get the loacal hosts name
    If gethostname(sHostName, Len(sHostName)) = ERROR_SUCCESS Then
        iPtr = InStr(sHostName, Chr$(0))
        If iPtr > 1 Then GetIPHost = Mid$(sHostName, 1, iPtr - 1)
    End If

    SocketsCleanup

End Function

Public Function IsValidEmailAddress(AddressString As String)  ' As Boolean

    Dim sTmp()      As String

    ' assume failure
    IsValidEmailAddress = False

    ' sould have one "@"
    sTmp = Split(AddressString, "@")
    If UBound(sTmp) <> 1 Then Exit Function

    IsValidEmailAddress = IsValidIPHost(sTmp(1))

End Function

Public Function Ping(Address As String, _
        Optional RoundTripTime As String = "", _
        Optional DataSize As String = "", _
        Optional DataMatch As Boolean = False) As Boolean

    ' Ping a remote host
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim ECHO            As ICMP_ECHO_REPLY
    Dim iPtr            As Integer
    Dim Dt              As String
    Dim sAddress        As String
    Dim hPort           As Long
    Dim lAddress        As Long
    Dim bytAddr(3)      As Byte
    
    If bInEXE Then On Local Error GoTo DPErr

    ' assume failure
    Ping = False

    ' if passed a name, get the IP address
    If Not IsDottedQuad(Address) Then
        sAddress = GetIPAddress(Address)
    Else
        sAddress = Address
    End If

    If sAddress = "" Then Exit Function

    If SocketsInitialize Then

        ' build string of random characters
        For iPtr = 1 To DATA_SIZE
            Dt = Dt & Chr$(Rnd() * 254 + 1)
        Next

        ' ping an ip address, passing the
        ' address and the ECHO structure
        lAddress = AddressStringToLong(sAddress)
        hPort = IcmpCreateFile()
        IcmpSendEcho hPort, lAddress, Dt, Len(Dt), 0, ECHO, Len(ECHO), PING_TIMEOUT
        IcmpCloseHandle hPort

        ' get the results from the ECHO structure
        RoundTripTime = ECHO.RoundTripTime
        CopyMemory bytAddr(0), ECHO.Address, 4
        Address = CStr(bytAddr(0)) & "." & _
                  CStr(bytAddr(1)) & "." & _
                  CStr(bytAddr(2)) & "." & _
                  CStr(bytAddr(3))

        DataSize = ECHO.DataSize & " bytes"

        iPtr = InStr(ECHO.Data, Chr$(0))
        If iPtr > 1 Then DataMatch = (Left$(ECHO.Data, iPtr - 1) = Dt)
        If ECHO.Status = 0 And ECHO.Address = lAddress Then Ping = True

        SocketsCleanup

    End If

    Exit Function

DPErr:

End Function

Public Sub Send()

    Dim sSenderName         As String
    Dim sToHeader           As String
    Dim sCcHeader           As String
    Dim iCtr                As Integer
    Dim sAuth               As String
    Dim sTxt                As String
    Dim strBoundry          As String
    Dim bMimeMultiPart      As Boolean
    Dim sQT                 As String
    Dim fStart              As Single
    Dim fTimeOut            As Single
    Dim lSendBuffSize       As Long
    
    ' Quote " character
    sQT = Chr$(34)

    ' general catch all error handler only
    ' works when running in stand alone EXE
    If bInEXE Then On Local Error GoTo Err_Send

    ' check for multipart MIME
    If etEncodeType = MIME_ENCODE And utMail.lAttachCount > 0 Then
        bMimeMultiPart = True
    Else
        bMimeMultiPart = False
    End If

    ' HTML & UU Encode are mutually exclusive
    If pbHtmlText = True And etEncodeType = UU_ENCODE Then AddError ERR_HTML_REQUIRES_MIME

    ' check recipient count
    If UBound(utMail.sToAddr) < 0 Then AddError ERR_NO_REC_EMAIL
    If UBound(utMail.sToAddr) + UBound(utMail.sCcAddr) + UBound(utMail.sBccAddr) + 3 > MAX_RECIPIENTS Then AddError ERR_RECIPIENT_COUNT

    ' we won't try to send if there's already an error
    If pColErrors.Count > 0 Then
        SendFail
        Exit Sub
    End If

    ' get the mail size
    plMailSize = EstimateMailSize

    ' this flag gets set when a socket error occurs or the host cannot process an
    ' input command, see 'SendFail', 'sckMail_DataArrival' & 'WaitUntilTrue' Subs
    pbExitImmediately = False

    With sckMail

        ' if not already conected then connect to the remote host
        If .State <> sckConnected Then
            If Not ConnectToHost Then Exit Sub
        End If

        ' reset the progress counter
        plBytesSent = 0

        ' tell the host who the mail is 'From
        RaiseEvent Status("Sending Sender Information...")
        pbRequestAccepted = False
        If pbAuthMailFromOK Then sAuth = " AUTH=" & utMail.sFromAddr Else sAuth = vbNullString
        .SendData "MAIL FROM: <" & utMail.sFromAddr & ">" & sAuth & vbCrLf
        If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub

        ' tell the host who the recipients are
        ' build the 'To:' header string 'sToHeader' too
        RaiseEvent Status("Sending Recipient Information...")
        For iCtr = 0 To UBound(utMail.sToAddr)
            ' send the recipient address & wait for a reply
            pbRequestAccepted = False
            .SendData "RCPT TO: <" & Trim$(utMail.sToAddr(iCtr)) & ">" & vbCrLf
            If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub

            ' build the 'To:' header string for later...
            If iCtr <= UBound(utMail.sToDisplayName) Then
                If Len(Trim$(utMail.sToDisplayName(iCtr))) Then sToHeader = sToHeader & sQT & Trim$(utMail.sToDisplayName(iCtr)) & sQT
                If Len(sToHeader) Then sToHeader = sToHeader & " "
            End If
            sToHeader = sToHeader & "<" & utMail.sToAddr(iCtr) & ">"
            If iCtr < UBound(utMail.sToAddr) Then sToHeader = sToHeader & ", " & vbCrLf & "    "
        Next

        ' send Cc: recipient addresses (just more 'RCPT TO' addresses)
        ' build the 'Cc:' header string too
        For iCtr = 0 To UBound(utMail.sCcAddr)
            ' send the recipient address & wait for a reply
            pbRequestAccepted = False
            .SendData "RCPT TO: <" & Trim$(utMail.sCcAddr(iCtr)) & ">" & vbCrLf
            If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub

            ' build the 'Cc:' header string for later...
            If iCtr <= UBound(utMail.sCcDisplayName) Then
                If Len(Trim$(utMail.sCcDisplayName(iCtr))) Then sCcHeader = sCcHeader & sQT & Trim(utMail.sCcDisplayName(iCtr)) & sQT
                If Len(sCcHeader) Then sCcHeader = sCcHeader & " "
            End If
            sCcHeader = sCcHeader & "<" & utMail.sCcAddr(iCtr) & ">"
            If iCtr < UBound(utMail.sCcAddr) Then sCcHeader = sCcHeader & ", " & vbCrLf & "    "
        Next

        ' send Bcc: recipient addresses (more of the same)
        ' no display headers here, these are blind
        For iCtr = 0 To UBound(utMail.sBccAddr)
            ' send the recipient address & wait for a reply
            pbRequestAccepted = False
            .SendData "RCPT TO: <" & Trim$(utMail.sBccAddr(iCtr)) & ">" & vbCrLf
            If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub
        Next

        ' tell the remote host we're ready to send data
        RaiseEvent Status("Sending Message...")
        pbDataOK = False
        .SendData "DATA" & vbCrLf
        If Not WaitUntilTrue(pbDataOK, plMessageTimeOut, True) Then Exit Sub

        ' OK, the host is ready for data, this is where the mail message starts
        ' Send the mail headers (the ones displayed on the target email client)
        pbRequestAccepted = False
        ' send English foramted date/time string
        '''.SendData "Date: " & Format$(Now, "ddd, d mmm yyyy hh:nn:ss ") & psTimeZoneBias & vbCrLf
        .SendData "Date: " & sDay(Weekday(Now)) & ", " & Day(Now) & " " & sMonth(Month(Now)) & _
                  Format$(Now, " yyyy hh:nn:ss ") & psTimeZoneBias & vbCrLf
        If Len(Trim$(utMail.sFromDisplayName)) Then sSenderName = sQT & utMail.sFromDisplayName & sQT & " "
        sSenderName = sSenderName & "<" & utMail.sFromAddr & ">"
        .SendData "From: " & sSenderName & vbCrLf
        .SendData "To: " & sToHeader & vbCrLf
        If Len(sCcHeader) Then .SendData "Cc: " & sCcHeader & vbCrLf
        If Len(utMail.sReplyToAddr) Then .SendData "Reply-to: <" & utMail.sReplyToAddr & ">" & vbCrLf
        .SendData "Subject: " & utMail.sSubject & vbCrLf

        ' MIME headers...
        If etEncodeType = MIME_ENCODE Then
            ' create a Unique-Boundary string for multi-part MIME encoding
            strBoundry = "----_=_NextPart_000_" & Format$(Now, "00000000.00000000")

            ' send the MIME start headers
            .SendData "MIME-Version: 1.0" & vbCrLf
            .SendData "X-Priority: " & Trim$(Str$(etPriority)) & vbCrLf
            .SendData "X-MSMail-Priority: " & psPriority & vbCrLf
            .SendData "X-Mailer: vbSendMail " & App.Major & "." & App.Minor & vbCrLf
            If pbReceipt Then .SendData "Disposition-Notification-To: " & sSenderName & vbCrLf
            ' if it's multi part send the boundry info
            If bMimeMultiPart Then
                .SendData "Content-Type: multipart/mixed;" & vbCrLf
                .SendData vbTab & "boundary=" & sQT & strBoundry & sQT & vbCrLf & vbCrLf
                .SendData "This is a multi-part message in MIME format." & vbCrLf & vbCrLf
                ' send the MIME boundry and content headers for the message body
                .SendData "--" & strBoundry & vbCrLf
            End If
            ' plain or html text...
            If pbHtmlText Then sTxt = "text/html;" Else sTxt = "text/plain;"
            .SendData "Content-Type: " & sTxt & vbCrLf
            .SendData vbTab & "charset=" & sQT & "iso-8859-1" & sQT & vbCrLf
            If pb8BitMail Then sTxt = "8bit" Else sTxt = "7bit"
            .SendData "Content-Transfer-Encoding: " & sTxt & vbCrLf
            ' if we're sending html & the user supplied the content base then send it too
            If pbHtmlText Then If Len(psContentBase) Then .SendData "Content-Base: " & sQT & psContentBase & sQT & vbCrLf
        End If

        .SendData vbCrLf

        ' Send the message body
        .SendData utMail.sMailMessage & vbCrLf & vbCrLf

        ' Send attachments, if any...
        For iCtr = 0 To utMail.lAttachCount - 1
            RaiseEvent Status("Sending Attachment, " & utMail.sAttachNameOnly(iCtr) & "...")
            If etEncodeType = MIME_ENCODE Then
                ' send the next MIME boundry & content headers
                .SendData "--" & strBoundry & vbCrLf
                .SendData "Content-Type: " & GetContentType(utMail.sAttachNameOnly(iCtr)) & ";" & vbCrLf
                .SendData vbTab & "name=" & sQT & utMail.sAttachNameOnly(iCtr) & sQT & vbCrLf
                .SendData "Content-Transfer-Encoding: base64" & vbCrLf
                .SendData "Content-Disposition: attachment;" & vbCrLf
                .SendData vbTab & "filename=" & sQT & utMail.sAttachNameOnly(iCtr) & sQT & vbCrLf & vbCrLf
                ' send the encoded file
                EncodeAndSendFile utMail.sAttachment(iCtr), MIME_ENCODE
                If pbExitImmediately Then Exit Sub
                .SendData vbCrLf
            Else
                ' start a UUEncode session
                .SendData "begin 600 " & utMail.sAttachNameOnly(iCtr) & vbCrLf
                ' send the encoded file
                EncodeAndSendFile utMail.sAttachment(iCtr), UU_ENCODE
                If pbExitImmediately Then Exit Sub
                ' send the ending sequence
                .SendData "end" & vbCrLf
            End If

            ' wait for the sckMail send buffer to clear
            Do While plBytesRemaining
                ' timeout code...
                fStart = Timer
                ' Deal with timer being reset at Midnight
                If fStart + plMessageTimeOut < 86400 Then
                    fTimeOut = fStart + plMessageTimeOut
                Else
                    fTimeOut = (fStart - 86400) + plMessageTimeOut
                End If
                ' wait for a change in the send buffer
                ' if it's changing, everything is OK
                lSendBuffSize = plBytesRemaining
                Do Until lSendBuffSize <> plBytesRemaining
                    If Timer >= fTimeOut Then
                        TimeOut
                        Exit Sub
                    End If
                    If plBytesRemaining = 0 Then Exit Do
                    Sleep (10)
                    DoEvents
                Loop
            Loop
        Next iCtr

        If bMimeMultiPart = True Then
            ' send the MIME closing boundry header
            .SendData "--" & strBoundry & "--" & vbCrLf
        End If

        ' Send the 'end of mail' string
        pbRequestAccepted = False
        .SendData "." & vbCrLf
        If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub
        
    End With

    ' send completion notifications...
    RaiseEvent Progress(100)
    RaiseEvent Status("Transmission Complete...")
    RaiseEvent SendSuccesful

    ' if the Public Function Connect() was called,
    ' stay connected to the host, otherwise disconnect
    If Not pbManualDisconnect Then DisconnectFromHost

    Exit Sub

Err_Send:

    ' add the error to the error collection
    AddError Err.Description
    SendFail

End Sub

Public Sub Shutdown()

' stub function, here to maintain binary
' compatibility with previous versions.

End Sub


' ******************************************************************************
' *      Private Class Functions                                               *
' ******************************************************************************

Private Sub AddError(ByVal ErrStr As String)

    ' add error string to the error collection
    
    On Local Error Resume Next
    pColErrors.Add ErrStr, ErrStr

End Sub

Private Function AddressStringToLong(ByVal tmp As String) As Long

    ' convert an ip address string to a long value
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim sParts()    As String

    sParts = Split(tmp, ".")

    If UBound(sParts) <> 3 Then
        AddressStringToLong = 0
        Exit Function
    End If

    ' build the long value out of the
    ' hex of the extracted strings
    AddressStringToLong = Val("&H" & Right$("00" & Hex$(sParts(3)), 2) & _
            Right$("00" & Hex$(sParts(2)), 2) & _
            Right$("00" & Hex$(sParts(1)), 2) & _
            Right$("00" & Hex$(sParts(0)), 2))

End Function

Private Function bInEXE() As Boolean

' ******************************************************************************
'
' Synopsis:     Check if application is running in the VB IDE or stand alone EXE.
'
' Parameters:   none
'
' Return:       True if running in EXE, False if running in IDE
'
' Description:
'
' Debug.print 1/0 will error produce a divide by zero error if running in IDE.
' If running in exe debug.print statement will be ignored
'
' ******************************************************************************

' modified version of Brian Gillham's code
' sample available at www.freevbcode.com

    On Local Error GoTo ErrorHandler

    Debug.Print 1 / 0                               ' this line will fail in the IDE
    bInEXE = True                                   ' this line will execute only in EXE or dll

    Exit Function

ErrorHandler:

    bInEXE = False

End Function

Private Function ConnectToHost() As Boolean

    Dim iCtr            As Integer
    Dim sHello          As String

    If bInEXE Then On Local Error GoTo Connect_Error

    ' reset var's
    pbRequestAccepted = False
    pbDataOK = False
    pbAuthLoginSupported = False
    pbAuthMailFromOK = False
    pbAuthLoginSuccess = False
    pbExitImmediately = False
    ConnectToHost = False

    ' open an SMTP session...
    With sckMail

        ' setup the port
        If .State = sckClosed Then
            .RemoteHost = psSMTPHost
            .RemotePort = plSMTPPort
            pbConnected = False
        ElseIf .State = sckConnected Then
            ConnectToHost = True
            Exit Function
        End If

        ' open a connection with the remote host
        ' try 'plConnectRetry' times before giving up
        RaiseEvent Status("Connecting to Server...")
        For iCtr = 1 To plConnectRetry
            If .State <> sckConnected Then
                If .State = sckClosed Then .Connect
                If WaitUntilTrue(pbConnected, plConnectTimeout, False) Then Exit For
                If pbExitImmediately Then Exit Function
                If .State = sckError Then .Close
            Else
                pbConnected = True
                Exit For
            End If
        Next
        
        ' if the connect attempt failed, exit
        If Not pbConnected Then
            TimeOut
            Exit Function
        End If

        ' once a connection is established, say 'hello
        RaiseEvent Status("Initializing Communications...")
        pbRequestAccepted = False
        ' EHLO is the extended (ESMTP) hello command, HELO is the standard hello command
        If pbUseAuthentication Then sHello = "EHLO " Else sHello = "HELO "
        .SendData sHello & Mid$(utMail.sFromAddr, InStr(utMail.sFromAddr, "@") + 1) & vbCrLf
        If Not WaitUntilTrue(pbRequestAccepted, plConnectTimeout, True) Then Exit Function

        ' Login Authentication ...
        ' the 'EHLO" command will cause the host to send a list of supported extensions
        ' via a series of 250 replies, wait to see if 'Auth Logon' is listed. The Sub
        ' sckMail_DataArrival will set pbUseAuthentication = True if Auth Login is
        ' supported by the remote host. If it is supported, Sub sckMail_DataArrival will
        ' respond to the host's Username & Password requests (psUserName, psPassword).
        If pbUseAuthentication = True Then
            If WaitUntilTrue(pbAuthLoginSupported, 5, False) Then
                RaiseEvent Status("Sending Login Authentication...")
                .SendData "AUTH Login" & vbCrLf
                If WaitUntilTrue(pbAuthLoginSuccess, 5, False) Then
                    RaiseEvent Status("Host Login OK!")
                Else
                    RaiseEvent Status("Host Login Failed!")
                    Exit Function
                End If
                If pbExitImmediately Then Exit Function
            Else
                RaiseEvent Status("Login Not Supported by Host, Continuing...")
            End If
        End If

    End With

    ConnectToHost = True

Connect_Error:

End Function

Private Function DecodeBase64String(ByVal str2Decode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim Bits(1 To 4)    As Byte
    Dim strDecode       As String

    ' for each 4 character group....
    For lPtr = 1 To Len(str2Decode) Step 4
        iLen = 4
        For iCtr = 0 To 3
            ' retrive the base 64 value, 4 at a time
            iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: Bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                ' not found
                Case 0: Exit Function
            End Select
        Next

        ' convert the 4, 6 bit values into 3, 8 bit values
        Bits(1) = Bits(1) * &H4 + (Bits(2) And &H30) \ &H10
        Bits(2) = (Bits(2) And &HF) * &H10 + (Bits(3) And &H3C) \ &H4
        Bits(3) = (Bits(3) And &H3) * &H40 + Bits(4)

        ' add the three new characters to the output string
        For iCtr = 1 To iLen - 1
            strDecode = strDecode & Chr$(Bits(iCtr))
        Next

    Next

    DecodeBase64String = strDecode

End Function

Private Sub DisconnectFromHost()

    With sckMail
        ' notify the user
        If .State <> sckClosed Then RaiseEvent Status("Closing Connection...")

        ' tell the host we're closing the connection...
        If .State = sckConnected Then
            pbRequestAccepted = False
            .SendData "QUIT" & vbCrLf
            WaitUntilTrue pbRequestAccepted, 2, False
        End If

        ' close the connection
        .Close
    End With

End Sub

Private Sub EncodeAndSendFile(ByVal strFile As String, ByVal Encode As ENCODE_METHOD)

' ******************************************************************************
'
' Synopsis:     Send a file attachment via an open socket
'
' Parameters:   strFile  - The input file name
'               Encode   -  type of encoding to use; MIME or UU
'
' Return:       nothing
'
' Description:
' Open the file & read characters in. Send the characters through the
' appropriate encoder, either MIME (Base64) or UUEncode, before
' tranmission via an open socket.
'
' ******************************************************************************

    Dim hFile               As Integer              ' file handle
    Dim sValue              As String               ' temp string buffer
    Dim bInFile()           As Byte                 ' byte array file buffer
    Dim lEventCtr           As Long                 ' counter
    Dim lChunkSize          As Long                 ' number of bytes to get
    Dim lNumBytes           As Long                 ' file pointer

    ' in case there's a file io error
    If bInEXE Then On Local Error GoTo File_Error

    ' open the file
    hFile = FreeFile
    Open strFile For Binary Access Read As #hFile
    ' bytes to read
    lNumBytes = LOF(hFile)

    If Encode = MIME_ENCODE Then

        Do While lNumBytes
            ' set input buffer size, MUST be a multiple of 57
            lChunkSize = IIf(lNumBytes > 11400, 11400, lNumBytes)

            ' read & Base 64 encode a group of characters
            bInFile = InputB(lChunkSize, #hFile)
            sckMail.SendData EncodeBase64Byte(bInFile)

            ' adjust file pointer
            lNumBytes = lNumBytes - lChunkSize

            DoEvents
        Loop

    ElseIf Encode = UU_ENCODE Then

        Do While Not EOF(hFile)
            ' read & UU encode a line of characters
            sValue = Input(45, #hFile)
            sckMail.SendData UUEncodeString(sValue) & vbCrLf

            ' DoEvents (occasionally)
            lEventCtr = lEventCtr + 1
            If lEventCtr Mod 50 = 0 Then DoEvents
        Loop

    End If

File_Done:

    Close #hFile

    Exit Sub

File_Error:

    AddError Err.Description
    SendFail
    pbExitImmediately = True
    Resume File_Done

End Sub

Private Function EncodeBase64Byte(InArray() As Byte) As Byte()

'******************************************************************************
'
' Synopsis:     Base 64 encode a byte array
'
' Parameters:   InArray  - The input byte array
'
' Return:       encoded byte array
'
' Description:
'   Convert a byte array to a Base 64 encoded byte array. Coerce 3 bytes into
'   4 by converting 3, 8 bit bytes into 4, 6 bit values. Each 6 bit value
'   (0 to 63) is then used as a pointer into a base64 byte array to derive a
'   character.
'
'******************************************************************************

Dim lInPtr              As Long         ' pointer into input array
Dim lOutPtr             As Long         ' pointer into output array
Dim OutArray()          As Byte         ' output byte array buffer
Dim lLen                As Long         ' number of extra bytes past 3 byte boundry
Dim iNewLine            As Long         ' line counter

' if size of input array is not a multiple of 3,
' increase it to the next multiple of 3
lLen = (UBound(InArray) - LBound(InArray) + 1) Mod 3
If lLen Then
    lLen = 3 - lLen
    ReDim Preserve InArray(UBound(InArray) + lLen)
End If

' create an output buffer
ReDim OutArray(UBound(InArray) * 2 + 100)

' step through the input array, 3 bytes at a time
For lInPtr = 0 To UBound(InArray) Step 3
    
    ' add CrLf as required
    If iNewLine = 19 Then
        OutArray(lOutPtr) = 13
        OutArray(lOutPtr + 1) = 10
        lOutPtr = lOutPtr + 2
        iNewLine = 0
    End If
    
    ' convert 3 bytes into 4 base 64 encoded bytes
    OutArray(lOutPtr) = pbBase64Byt((InArray(lInPtr) And &HFC) \ 4)
    OutArray(lOutPtr + 1) = pbBase64Byt((InArray(lInPtr) And &H3) * &H10 + (InArray(lInPtr + 1) And &HF0) \ &H10)
    OutArray(lOutPtr + 2) = pbBase64Byt((InArray(lInPtr + 1) And &HF) * 4 + (InArray(lInPtr + 2) And &HC0) \ &H40)
    OutArray(lOutPtr + 3) = pbBase64Byt(InArray(lInPtr + 2) And &H3F)
    
    ' update pointers
    lOutPtr = lOutPtr + 4
    iNewLine = iNewLine + 1
Next

' add terminator '=' as required
Select Case lLen
    Case 1: OutArray(lOutPtr - 1) = 61
    Case 2: OutArray(lOutPtr - 1) = 61: OutArray(lOutPtr - 2) = 61
End Select

' add CrLf if not already there
If OutArray(lOutPtr - 2) <> 13 Then
    OutArray(lOutPtr) = 13
    OutArray(lOutPtr + 1) = 10
    lOutPtr = lOutPtr + 2
End If

' resize output buffer and return
ReDim Preserve OutArray(lOutPtr - 1)
EncodeBase64Byte = OutArray

End Function

Private Function EncodeBase64String(ByRef str2Encode As String) As String

' ******************************************************************************
'
' Synopsis:     Base 64 encode a string
'
' Parameters:   str2Encode  - The input string
'
' Return:       encoded string
'
' Description:
' Convert a string to a byte array and pass to EncodeBase64Byte function (above)
' for Base64 conversion. Convert byte array back to a string and return.
'
' ******************************************************************************

    Dim tmpByte()   As Byte

    If Len(str2Encode) Then
    
        ' convert string to byte array
        tmpByte = StrConv(str2Encode, vbFromUnicode)

        ' pass to the byte array encoder
        tmpByte = EncodeBase64Byte(tmpByte)

        ' convert back to string & return
        EncodeBase64String = StrConv(tmpByte, vbUnicode)

    End If
    
End Function

Private Function EstimateMailSize() As Long

' ******************************************************************************
'
' Synopsis:     Estimate the size (number of bytes) of the mail message
'
' Parameters:   none
'
' Return:       long - number of bytes
'
' Description:
' Estimate the size in bytes of the mail message being sent. Include the
' message body, headers, attachments, etc. Account for type of encoding.
' The result is used to calculate send progress.
'
' ******************************************************************************

    Dim lNumBytes       As Long
    Dim iCtr            As Integer

    lNumBytes = 93

    ' Mail From
    lNumBytes = lNumBytes + Len(utMail.sFromAddr)

    ' login authentication
    If pbUseAuthentication Then
        lNumBytes = lNumBytes + 25 + Len(utMail.sFromAddr)
        If Len(psUserName) > 0 Then lNumBytes = lNumBytes + (Len(psUserName) * 4 \ 3)
        If Len(psPassword) > 0 Then lNumBytes = lNumBytes + (Len(psPassword) * 4 \ 3)
    End If

    ' To: recipients
    For iCtr = 0 To UBound(utMail.sToAddr)
        lNumBytes = lNumBytes + 15 + Len(utMail.sToAddr(iCtr)) * 2  ' sent twice, RCPT & 'To:' header
        If iCtr > 0 Then lNumBytes = lNumBytes + 6
    Next

    ' To Display
    For iCtr = 0 To UBound(utMail.sToDisplayName)
        lNumBytes = lNumBytes + Len(utMail.sToDisplayName(iCtr)) + 11
    Next

    ' Cc: recipients
    For iCtr = 0 To UBound(utMail.sCcAddr)
        lNumBytes = lNumBytes + 15 + Len(utMail.sCcAddr(iCtr)) * 2  ' sent twice, RCPT & 'Cc:' header
        If iCtr > 0 Then lNumBytes = lNumBytes + 6                  ' header
    Next

    ' Cc Display
    For iCtr = 0 To UBound(utMail.sCcDisplayName)
        lNumBytes = lNumBytes + Len(utMail.sCcDisplayName(iCtr)) + 11
    Next

    ' Bcc: recipients
    For iCtr = 0 To UBound(utMail.sBccAddr)
        lNumBytes = lNumBytes + 15 + Len(utMail.sBccAddr(iCtr))  ' RCPT & 'Bcc:' header
        If iCtr > 0 Then lNumBytes = lNumBytes + 6               ' header
    Next

    ' From:
    If Len(utMail.sFromDisplayName) Then lNumBytes = lNumBytes + Len(utMail.sFromDisplayName) + 3
    lNumBytes = lNumBytes + Len(utMail.sFromAddr)

    ' ReplyTo
    If Len(utMail.sReplyToAddr) Then lNumBytes = lNumBytes + Len(utMail.sReplyToAddr) + 14

    ' Subject
    lNumBytes = lNumBytes + Len(utMail.sSubject)

    ' Message body
    lNumBytes = lNumBytes + Len(utMail.sMailMessage)

    ' MIME headers....
    If etEncodeType = MIME_ENCODE Then
        lNumBytes = lNumBytes + 64
        If pbHtmlText = True And Len(psContentBase) > 0 Then lNumBytes = lNumBytes + 18 + Len(psContentBase)
        If pbReceipt Then lNumBytes = lNumBytes + 36 + Len(utMail.sFromDisplayName) + Len(utMail.sFromAddr)
    End If

    ' attachments
    If utMail.lAttachCount > 0 Then
        If etEncodeType = MIME_ENCODE Then
            lNumBytes = lNumBytes + utMail.lAttachFileSize * 4 \ 3 + 42  ' length of encoded file
            lNumBytes = lNumBytes + (utMail.lAttachFileSize \ 57) * 2    ' add CrLf for each line
            lNumBytes = lNumBytes + utMail.lAttachNameSize * 2           ' add file name twice
            lNumBytes = lNumBytes + (utMail.lAttachCount * 182)          ' attachment header per file
            lNumBytes = lNumBytes + 290                                  ' additional MIME headers
        Else
            lNumBytes = lNumBytes + utMail.lAttachFileSize * 4 \ 3       ' length of encoded file
            lNumBytes = lNumBytes + (utMail.lAttachFileSize \ 45) * 3    ' add length char + CrLf for each line
            lNumBytes = lNumBytes + utMail.lAttachNameSize               ' add file name once
            lNumBytes = lNumBytes + (utMail.lAttachCount * 20)           ' attachment header per file
        End If
    End If

    EstimateMailSize = lNumBytes

End Function

Private Function FormatMail(ByVal strIn As String) As String

' ******************************************************************************
'
' Synopsis:     Re-format text lines per RFC 821
'
' Parameters:   strIn   - The input string to be formated
'
' Return:       re-formated string
'
' Description:
' RFC 821 places the following restrictions on user text:
' 1) Before sending a line of mail text begining with a '.
' the sender will add an additional '.
'
' 2) The receiver checks each line of mail text, if a line is single '.
' it is the end of the mail message. If the first character is
' a '.' and there are other characters on the line, the first '.
' is deleted.
'
' 3) The maximum line lenght will not exceed 1000 characters
'
' ******************************************************************************

    Dim sTextLine()     As String
    Dim sRemainder      As String
    Dim sNewLine        As String
    Dim sDelimiter      As String
    Dim lPtr            As Long
    Dim lSplit          As Long

    If Len(strIn) = 0 Then Exit Function

    ' Select the correct delimiter character
    If InStr(strIn, vbCrLf) Then
        sDelimiter = vbCrLf
    ElseIf InStr(strIn, vbCr) Then
        sDelimiter = vbCr
    Else
        sDelimiter = vbNullString
    End If

    ' split the text into seperate lines
    sTextLine() = Split(strIn, sDelimiter)

    ' process each line
    For lPtr = 0 To UBound(sTextLine)
        ' check for lines starting with a '.
        ' when found, add a second '.
        If Left$(sTextLine(lPtr), 1) = "." Then sTextLine(lPtr) = "." & sTextLine(lPtr)

        ' check that the line is not too long (account for 2 extra characters - vbCrLf)
        ' break into smaller elements as required
        If Len(sTextLine(lPtr)) > MAX_TEXTLINE_LEN - 2 Then
            sRemainder = sTextLine(lPtr)
            sNewLine = vbNullString
            If sDelimiter = vbNullString Then sDelimiter = vbCrLf
            Do While Len(sRemainder) > MAX_TEXTLINE_LEN - 2
                ' try to split at a space character, if not then split at MAX_TEXTLINE_LEN - 2
                lSplit = InStrRev(sRemainder, " ", MAX_TEXTLINE_LEN - 2)
                If lSplit = 0 Then lSplit = MAX_TEXTLINE_LEN - 2
                ' insert a vbCrLf at the split point
                sNewLine = sNewLine & Mid$(sRemainder, 1, lSplit) & sDelimiter
                sRemainder = Mid$(sRemainder, lSplit + 1)
            Loop
            sTextLine(lPtr) = sNewLine & sRemainder
        End If
    Next

    FormatMail = Join(sTextLine, sDelimiter)

End Function

Private Function IsDottedQuad(ByVal HostString As String) As Boolean

    ' verify that a string is 'xxx.xxx.xxx.xxx' format
    
    Dim sSplit()        As String
    Dim iCtr            As Integer

    ' split at the "."
    sSplit = Split(HostString, ".")

    ' should be 4 elements
    If UBound(sSplit) <> 3 Then Exit Function

    ' check each element
    For iCtr = 0 To 3
        ' should be numeric
        If Not IsNumeric(sSplit(iCtr)) Then Exit Function

        ' range check
        If iCtr = 0 Then
            If Val(sSplit(iCtr)) > 239 Then Exit Function
        Else
            If Val(sSplit(iCtr)) > 255 Then Exit Function
        End If
    Next
    
    IsDottedQuad = True

End Function

Private Function IsValidIPHost(ByVal HostString As String) As Boolean

    ' validate a host string

    Dim sHost               As String
    Dim sSplit()            As String

    sHost = UCase$(Trim$(HostString))

    ' if it's a dotted quad it's OK
    If IsDottedQuad(sHost) Then
        IsValidIPHost = True
        Exit Function
    End If

    sSplit = Split(sHost, ".")

    ' it's dotted quad, top level domain?
    If UBound(sSplit) > 0 And InStr(TOP_DOMAINS, sSplit(UBound(sSplit))) > 0 Then
        IsValidIPHost = True
        Exit Function
    End If

End Function

Private Function RegGet(ByVal sSettingName As String, ByVal sDefaultValue As String) As String

    ' get registry setting
    RegGet = GetSetting(REG_KEY, SETTINGS_KEY, sSettingName, sDefaultValue)

End Function

Private Sub RegSave(ByVal sSettingName As String, ByVal sNewValue As String)

    ' save registry setting
    If plPersistentSettings Then SaveSetting REG_KEY, SETTINGS_KEY, sSettingName, sNewValue

End Sub

Private Sub RemoveError(ByVal ErrStr As String)

    ' remove an error string from the error collection

    Dim i   As Long

    On Local Error Resume Next

    ' walk the collection looking for the string to remove
    For i = 1 To pColErrors.Count
        If pColErrors(i) = ErrStr Then pColErrors.Remove ErrStr
    Next

End Sub

Private Sub SendFail()

    Dim iCtr            As Integer
    Dim sErrorString    As String

    ' report all errors to the user
    For iCtr = 1 To pColErrors.Count
        sErrorString = sErrorString & pColErrors(iCtr) & vbCrLf
    Next

    RaiseEvent SendFailed(sErrorString)

    ' close the connection with the remote host
    If sckMail.State <> sckClosed Then DisconnectFromHost

    ' set flag to exit 'Send' Sub without further processing
    pbExitImmediately = True

    ' clear all errors
    Set pColErrors = New Collection

End Sub

Private Sub SocketsCleanup()

    ' Cleanup Windows sockets
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    WSACleanup

End Sub

Private Function SocketsInitialize() As Boolean

    ' Initialize Windows sockets
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim WSAD            As WSADATA

    SocketsInitialize = False

    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then Exit Function
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then Exit Function

    SocketsInitialize = True

End Function

Private Sub TimeOut()

    ' time out occured, add the 'Timeout' error
    ' to the error collection
    AddError ERR_TIMEOUT
    SendFail

End Sub

Private Function UUEncodeString(ByRef str2UUEncode As String) As String

' ******************************************************************************
'
' Synopsis:     UUEncode a string
'
' Parameters:   str2UUEncode  - The input string
'
' Return:       encoded string
'
' Description:
' UU Encode a string. Coerce 3 bytes into 4 by converting 3, 8 bit bytes into
' 4, 6 bit values. Each 6 bit value (0 to 63) is then used as a pointer into
' the UUEncode string array to derive the correct character. The string will
' be a multiple of 4 bytes in lenght after conversion, padded with '=' as
' required. The line length will be encoded as a leading character
' (same 0 to 63 encoding) in the return string.
'
' ******************************************************************************

    Dim sValue              As String
    Dim lPtr                As Long
    Dim lCtr                As Long
    Dim lLen                As Long
    Dim lLineLen            As Long
    Dim sEncoded            As String
    Dim Bits8(1 To 3)       As Byte
    Dim Bits6(1 To 4)       As Byte

    lLineLen = Len(str2UUEncode)
    ' lines are limited to 63
    If lLineLen > 63 Then Exit Function

    For lCtr = 1 To Len(str2UUEncode) Step 3

        ' Get 3 characters
        sValue = Mid$(str2UUEncode, lCtr, 3)
        lLen = Len(sValue)

        ' Move string data into a byte array, then
        ' swap bits to create 4, 6 bit values (0-63)
        If lLen < 3 Then Erase Bits8
        CopyMemory Bits8(1), ByVal sValue, lLen
        Bits6(1) = (Bits8(1) And &HFC) \ &H4
        Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
        Bits6(3) = (Bits8(2) And &HF) * &H4 + (Bits8(3) And &HC0) \ &H40
        Bits6(4) = Bits8(3) And &H3F

        ' Encode new 4 byte string by selecting a character from
        ' the array. Length is determined by 'lLen' to make sure
        ' the file attachment is the right length
        For lPtr = 1 To lLen + 1
            sEncoded = sEncoded & psUUEncodeChr(Bits6(lPtr))
        Next

    Next

    ' add the line length character
    sEncoded = psUUEncodeChr(lLineLen) & sEncoded

    ' The decoder expects the size to be a multiple of 4 bytes.
    ' Possible sizes for the last packet are: 2, 3 & 4.
    Select Case lLen + 1
        Case 2: sEncoded = sEncoded & "=="          ' send two pad characters
        Case 3: sEncoded = sEncoded & "="           ' send one pad character
    ' no pad characers needed
    End Select

    UUEncodeString = sEncoded

End Function

Private Sub ValidateAddress(ByVal sRecip As String, ByVal sError As String)
    
    ' Validate Recipient, Cc and Bcc email address
    ' Appropriate validation methods for are:
    ' VALIDATE_NONE, VALIDATE_SYNTAX

    Dim iPtr            As Integer
    Dim sRecipArray()   As String

    RemoveError sError

    ' if VALIDATE_SYNTAX...
    If etEmailValidation = VALIDATE_SYNTAX Then
        ' split components into an array
        sRecipArray = Split(sRecip, psDelimiter)
        For iPtr = 0 To UBound(sRecipArray)
            ' validate address...
            If IsValidEmailAddress(sRecipArray(iPtr)) = False Then
                AddError sError
                Exit For
            End If
        Next
    End If

End Sub

Private Function WaitUntilTrue(ByRef Flag As Boolean, ByVal SecondsToWait As Long, Optional ByVal bRaiseTimeOutError As Boolean = True) As Boolean

    Dim fStart              As Single
    Dim fTimetoQuit         As Single

    If SecondsToWait < 1 Then Exit Function

    fStart = Timer

    ' Deal with timer being reset at Midnight
    If fStart + SecondsToWait < 86400 Then
        fTimetoQuit = fStart + SecondsToWait
    Else
        fTimetoQuit = (fStart - 86400) + SecondsToWait
    End If

    Do Until Flag = True
        If Timer >= fTimetoQuit Then
            If bRaiseTimeOutError Then TimeOut
            Exit Function
        End If
        If pbExitImmediately Then Exit Function
        DoEvents
        Sleep (10)                                  ' added to reduce CPU load during wait periods
    Loop

    WaitUntilTrue = Flag

End Function


' ******************************************************************************
' *      Private Winsock OCX Events                                            *
' ******************************************************************************

Private Sub sckMail_Close()

    ' keep track of connection state
    pbConnected = False

End Sub

Private Sub sckMail_Connect()

    ' keep track of connection state
    pbConnected = True

End Sub

Private Sub sckMail_DataArrival(ByVal bytesTotal As Long)

' ********************************************************
' SMTP Reply codes outlined in RFC 821
' ********************************************************
' 211 - System status/help reply
' 214 - Help message
' 220 - <domain> Service ready
' 221 - <domain> Service closing channel
' 250 - OK: action completed
' 251 - User not local, will forward to <domain>
' 354 - OK: Start mail input, end with <CrLf>.<CrLf>
' 421 - <domain> Service not available, closing channel
' 450 - Mailbox busy, action not taken
' 451 - Requested action aborted: error in processing
' 452 - Requested action not taken: insufficient system storage
' 500 - Syntax error, command unrecognized
' 501 - Syntax error in parameters or arguments
' 502 - Command not implimented
' 503 - Bad sequence of commands
' 504 - Command parameter not implimented
' 550 - Mailbox unavailable, action not taken
' 553 - Requested action not taken: mailbox name not allowed / invalid
' 551 - User not local, please try <forward-path>
' 552 - Requested action not taken: exceeds storage allocation
' 554 - Transaction failed

' ********************************************************
' ESMTP AUTHentication extensions outlined in RFC 2554
' ********************************************************
' 235 - Authentication successful
' 334 - Server challenge / ready response
' 432 - A password transition is needed
' 454 - Temporary authentication failure
' 530 - Authentication required
' 534 - Authentication mechanism is too weak
' 535 - Server rejected authentication
' 538 - Encryption required for requested authentication mechanism

    Dim strAns          As String
    Dim sMsg            As String

    sckMail.GetData strAns, vbString

    ' clean up the message portion
    sMsg = Trim$(Mid$(strAns, 4))
    If Right$(strAns, 2) = vbCrLf Then sMsg = Mid$(sMsg, 1, Len(sMsg) - 2)

    ''Debug.Print strAns

    Select Case Val(Left$(strAns, 3))

        ' Ready
        Case 220
            pbRequestAccepted = True


        ' OK, check for authentication support
        Case 250
            pbRequestAccepted = True
            If InStr(1, strAns, "auth login", vbTextCompare) Then pbAuthLoginSupported = True
            If InStr(1, strAns, "auth=login", vbTextCompare) Then pbAuthMailFromOK = True

        ' OK
        Case 221, 251
            pbRequestAccepted = True


        ' Auth Login OK
        Case 235
            pbAuthLoginSuccess = True

        ' mail host 'AUTH' challenge
        Case 334

            ' respond to the hosts challenge
            
            ' username requested
            If InStr(1, DecodeBase64String(sMsg), "username", vbTextCompare) Then
                sckMail.SendData EncodeBase64String(psUserName) & vbCrLf
            
            ' password reequested
            ElseIf InStr(1, DecodeBase64String(sMsg), "password", vbTextCompare) Then
                sckMail.SendData EncodeBase64String(psPassword) & vbCrLf
            
            ' unexpected or unsupported challenge, cancel Auth request
            ' which will result in a 501 error reply from the host
            Else
                sckMail.SendData vbCrLf & "*" & vbCrLf
            End If

        ' OK, send data
        Case 354
            pbDataOK = True

        ' do nothing
        Case 211, 214

        ' host didn't like what we sent or couldn't process it
        Case Else
            AddError sMsg
            SendFail

    End Select

End Sub

Private Sub sckMail_Error(ByVal Number As Integer, Description As String, _
        ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, _
        ByVal HelpContext As Long, CancelDisplay As Boolean)

    ' socket error, add the error to the error collection
    AddError Description
    SendFail

End Sub

Private Sub sckMail_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

    Dim lNewValue           As Long
    Static lProgressLast    As Long

    ' add up sent bytes
    plBytesSent = plBytesSent + bytesSent

    ' calculate the percentage of the total
    If plMailSize > 0 Then lNewValue = CLng(CSng(plBytesSent / plMailSize) * 100)
    If lNewValue > 100 Then lNewValue = 100

    ' update if the value changed
    If lNewValue <> lProgressLast Then
        lProgressLast = lNewValue
        RaiseEvent Progress(lNewValue)
    End If

    ' keep track of what's left
    plBytesRemaining = bytesRemaining

End Sub
