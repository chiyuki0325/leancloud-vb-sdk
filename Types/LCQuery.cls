VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LCQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Properties
Private Utils As cLeanCloudUtils
Private ClassName As String

Public Sub Initialize( _
    ByRef oUtils As cLeanCloudUtils, _
    ByRef sClassName As String _
)
    Set Utils = oUtils
    ClassName = sClassName
End Sub


Public Function GetObject(ByRef ObjectId As String) As Object
    Dim JSON As New cJSON
    With Utils.NewRequest("GET", ("/classes/" + ClassName + "/" + ObjectId))
        .Send
        .WaitForResponse
        If .Status = 404 Then
            Err.Raise 101, , "Class or object doesn't exists."
        End If
        If .ResponseText = "{}" Then
            Err.Raise 101, , "Object doesn't exists."
            Exit Function
        End If
        Set GetObject = New LCObject
        GetObject.Initialize Utils, ClassName, ObjectId
        GetObject.Deserialize JSON.Parse(.ResponseText)
    End With
End Function





