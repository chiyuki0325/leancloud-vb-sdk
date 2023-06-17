VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLeanCloudUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Enums
Public Enum RequestAuthModes
    UseAppKey
    UseSign
End Enum

'Properties (pass from cLeanCloud object)
Private AppID As String, AppKey As String, Host As String
Private MasterKey As String
Private ApiVersion As String
Private RequestAuthMode As RequestAuthModes

'Objects
Private MD5CryptoServiceProvider As Object
Private UTF8Encoding As Object

Sub Initialize( _
    ByRef sAppID As String, _
    ByRef sAppKey As String, _
    ByRef sHost As String, _
    ByRef sApiVersion As String _
)
    'Pass properties
    AppID = sAppID
    AppKey = sAppKey
    Host = sHost
    ApiVersion = sApiVersion
    
    'Select auth mode
    'Try Sign mode first
    'If .NET Framework is unvailable, use X-LC-Key to auth
    On Error GoTo ErrThenUseAppKey
    Set MD5CryptoServiceProvider = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    Set UTF8Encoding = CreateObject("System.Text.UTF8Encoding")
    RequestAuthMode = UseSign
    Exit Sub
ErrThenUseAppKey:
    If Err.Number = 429 Then
        RequestAuthMode = UseAppKey
    End If
End Sub

Sub InitMasterKey(ByRef sMasterKey As String)
    MasterKey = sMasterKey
End Sub


'Create a new WinHttpRequest object
Function NewRequest( _
         ByRef sMethod As String, _
         ByRef sURL As String, _
         Optional UseMasterKey As Boolean = False, _
         Optional ContentType = "application/json" _
       ) As WinHttp.WinHttpRequest
    Set NewRequest = New WinHttp.WinHttpRequest
    With NewRequest
        .Open sMethod, ("https://" + Host + "/" + ApiVersion + sURL), True
        .Option(WinHttpRequestOption_SecureProtocols) = WinHttp.SecureProtocol_ALL Or _
                                                        WinHttp.SecureProtocol_TLS1 Or _
                                                        WinHttp.SecureProtocol_TLS1_1 Or _
                                                        WinHttp.SecureProtocol_TLS1_2
        .SetRequestHeader "X-LC-Id", AppID
        If RequestAuthMode = UseSign Then
            'Sign
            If UseMasterKey Then
                .SetRequestHeader "X-LC-Sign", CalculateRequestSign(MasterKey) + ",master"
            Else
                .SetRequestHeader "X-LC-Sign", CalculateRequestSign(AppKey)
            End If
        Else
            'AppKey
            If UseMasterKey Then
                .SetRequestHeader "X-LC-Key", MasterKey + ",master"
            Else
                .SetRequestHeader "X-LC-Key", AppKey
            End If
        End If
        If sMethod <> "GET" Then .SetRequestHeader "Content-Type", ContentType
    End With
End Function


'Get current Unix timestamp
Public Function GetTimeStamp() As String
    'https://stackoverflow.com/questions/54535040/unix-timestamp-with-millisecond-precision-in-vb6
    With New cLeanCloudTimeConvert
        GetTimeStamp = CStr(Round(CDec(DateDiff("s", "1/1/1970", .ConvertLocalToGMT(Date)) + Timer) * 1000))
    End With
End Function


'Calculate request sign
Private Function CalculateRequestSign(Key As String) As String
    'md5( timestamp + App Key )
    '= md5(1453014943466UtOCzqb67d3sN12Kts4URwy8)
    '= d5bcbb897e19b2f6633c716dfdfaf9be
    '  -H "X-LC-Sign: d5bcbb897e19b2f6633c716dfdfaf9be,1453014943466" \
    
    Dim TimeStamp As String, MD5Hash As String
    TimeStamp = GetTimeStamp
    
    With CreateObject("MSXML2.DOMDocument")
        .LoadXML "<root />"
        .documentelement.DataType = "bin.Hex"
        .documentelement.nodeTypedValue = MD5CryptoServiceProvider.ComputeHash_2(UTF8Encoding.GetBytes_4(TimeStamp + Key))
        MD5Hash = LCase(Replace(.documentelement.Text, vbLf, ""))
    End With
    
    CalculateRequestSign = MD5Hash + "," + TimeStamp
End Function


'Convert Unix TimeStamp to Date
Public Function UnixTimeStampToDate(ByRef UnixTimeStamp As String) As Date
    With New cLeanCloudTimeConvert
        UnixTimeStampToDate = .GetLocalTimeFromGMT(DateAdd("s", Int(CDec(UnixTimeStamp) / 1000), "1/1/1970 00:00:00"))
    End With
End Function



'Convert Date To ISO Date format
Public Function ConvertDateToISO(InputDate As Date) As String
    ConvertDateToISO = Format(InputDate, "yyyy-mm-ddTHH:MM:SS.000Z")
End Function



'Convert ISO Date format to Date
Public Function ConvertISOToDate(InputISO As String) As Date
    ConvertISOToDate = CDate(Left$(Replace(InputISO, "T", " "), Len(InputISO) - 5))
End Function


