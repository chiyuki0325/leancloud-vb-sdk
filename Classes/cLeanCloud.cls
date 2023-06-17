VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLeanCloud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Constants
Private Const API_VERSION As String = "1.1"

'Properties
Private AppID As String, AppKey As String, Host As String
Private MasterKey As String

'Sub classes
Private Utils As New cLeanCloudUtils

'LCApplication.Initialize
Sub Initialize( _
    ByRef sAppID As String, _
    ByRef sAppKey As String, _
    ByRef sHost As String _
)
    AppID = sAppID
    AppKey = sAppKey
    Host = sHost
    
    'Init sub classes
    Utils.Initialize sAppID, sAppKey, sHost, API_VERSION
End Sub

'Init master key
Sub InitMasterKey(ByRef sMasterKey As String)
    MasterKey = sMasterKey
    'Init sub classes
    Utils.InitMasterKey sMasterKey
End Sub


'LCQuery
Function Query(ByRef ClassName As String) As LCQuery
    Set Query = New LCQuery
    Query.Initialize Utils, ClassName
End Function


'LCObject
Function Object( _
    Optional ByRef sClassName As String, _
    Optional ByRef sObjectID As String _
) As LCObject
    Set Object = New LCObject
    Object.Initialize Utils, sClassName, sObjectID
End Function



