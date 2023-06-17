VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LCObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Utils As cLeanCloudUtils
Private Dict As New Dictionary

Public ClassName As String, ObjectID As String
Public CreatedAt As Date, UpdatedAt As Date

Public Sub Initialize( _
    ByRef oUtils As cLeanCloudUtils, _
    ByRef sClassName As String, _
    Optional ByRef sObjectID As String _
)
    Set Utils = oUtils
    ClassName = sClassName
    ObjectID = sObjectID
End Sub

Public Property Get Item(Key)
    If VarType(Dict.Item(Key)) = vbObject Then
        Set Item = Dict.Item(Key)
    Else
        Item = Dict.Item(Key)
    End If
End Property

Public Property Let Item(Key, Value)
    If Dict.Exists(Key) Then
        Dict.Item(Key) = Value
    Else
        Dict.Add Key, Value
    End If
End Property

Public Property Set Item(Key, Value)
    If Dict.Exists(Key) Then
        Set Dict.Item(Key) = Value
    Else
        Dict.Add Key, Value
    End If
End Property

'If need to 'Set'
Private Function IsObjectType(Var) As Boolean
    Select Case VarType(Var)
        Case vbInteger, vbBoolean, vbString, vbLong
            IsObjectType = False
        Case vbDate, vbArray + vbByte, vbArray + vbString, vbArray + vbInteger, vbArray + vbLong, vbArray + vbBoolean
            IsObjectType = True
        Case Else
            Select Case TypeName(Var)
                Case "LCObject" 'Pointer
                'TODO: LCFile
                    IsObjectType = True
                Case Else
                    Err.Raise 13
            End Select
    End Select
End Function

Private Function ToLeanObject(Var) As Object  'Collection or Dictionary
    Select Case TypeName(Var)
        Case "Date"
            Dim OutputDict As New Dictionary
            With OutputDict
                .Add "__type", "Date"
                .Add "iso", Utils.ConvertDateToISO(Var)
            End With
            Set ToLeanObject = OutputDict
        Case "Byte()"
            Dim OutputDict As New Dictionary
            Dim XMLDocument As Object
            Set XMLDocument = CreateObject("MSXML2.DOMDocument")
            With XMLDocument
                .LoadXML "<root />"
                .documentelement.DataType = "bin.base64"
                .documentelement.nodeTypedValue = Var
            End With
            With OutputDict
                .Add "__type", "Bytes"
                .Add "base64", Replace(XMLDocument.documentelement.Text, vbLf, "")
            End With
            Set XMLDocument = Nothing
            Set ToLeanObject = OutputDict
        Case "LCObject"  'Pointer
            Dim OutputDict As New Dictionary
            With OutputDict
                .Add "__type", "Pointer"
                .Add "className", Var.ClassName
                .Add "objectId", Var.ObjectID
            End With
            Set ToLeanObject = OutputDict
        'TODO: LCFile
        Case "String()", "Integer()", "Boolean()", "Long()"
            Dim i, OutputCollection As New Collection
            For Each i In Var
                OutputCollection.Add i
            Next
        Case Else
            Err.Raise 13
    End Select
End Function

Private Function ToVBObject(LeanObject As Object)
    If TypeName(LeanObject) = "Dictionary" Then
        If LeanObject.Exists("__type") Then
            Select Case LeanObject("__type")
                Case "Date"
                    ToVBObject = Utils.ConvertISOToDate(LeanObject("iso"))
                Case "Bytes"
                    With CreateObject("MSXML2.XMLDocument")
                        .LoadXML "<root />"
                        .documentelement.DataType = "bin.base64"
                        .documentelement.Text = LeanObject("base64")
                        ToVBObject = .documentelement.nodeTypedValue
                    End With
                Case "LCObject"  'Pointer
                    Set ToVBObject = New LCObject
                    ToVBObject.Initialize Utils, LeanObject("className"), LeanObject("objectId")
                'TODO: LCFile
            End Select
        Else
            Set ToVBObject = LeanObject
        End If
    Else
        Set ToVBObject = LeanObject
    End If
End Function

Public Function Serialize() As Dictionary
    Set Serialize = New Dictionary
    Dim Key
    For Each Key In Dict.Keys
        If IsObjectType(Dict(Key)) Then
            Serialize.Add Key, ToLeanObject(Dict(Key))
        Else
            Serialize.Add Key, Dict(Key)
        End If
    Next
    With Serialize
        .Add "objectId", ObjectID
        If Not CreatedAt = "0:00.00" Then .Add "createdAt", Utils.ConvertDateToISO(CreatedAt)
        If Not UpdatedAt = "0:00.00" Then .Add "updatedAt", Utils.ConvertDateToISO(UpdatedAt)
    End With
End Function

Public Sub Deserialize(InputDict As Dictionary)
    Dim Key
    If InputDict.Exists("objectId") Then
        ObjectID = InputDict("objectId")
        InputDict.Remove ("objectId")
    End If
    If InputDict.Exists("createdAt") Then
        CreatedAt = Utils.ConvertISOToDate(InputDict("createdAt"))
        InputDict.Remove ("createdAt")
    End If
    If InputDict.Exists("updatedAt") Then
        UpdatedAt = Utils.ConvertISOToDate(InputDict("updatedAt"))
        InputDict.Remove ("updatedAt")
    End If
    For Each Key In InputDict.Keys
        If VarType(InputDict(Key)) = vbObject Then
            If VarType(ToVBObject(InputDict(Key))) = vbObject Then
                Set Me(Key) = ToVBObject(InputDict(Key))
            Else
                Me(Key) = InputDict(Key)
            End If
        Else
            Me(Key) = InputDict(Key)
        End If
    Next
End Sub

Public Sub Save()
    Dim JSON As New cJSON
    If ObjectID = "" Then
        'New object
        With Utils.NewRequest("POST", "/classes/" + ClassName)
            .Send JSON.ToString(Serialize)
            .WaitForResponse
            If .Status = 201 Then
                Dim Response As Dictionary
                Set Response = JSON.Parse(.ResponseText)
                If Response.Exists("error") Then
                    Err.Raise 0, , Response("error")
                Else
                    'Success
                    ObjectID = Response("objectId")
                    CreatedAt = Utils.ConvertISOToDate(Response("createdAt"))
                End If
            Else
                Err.Raise 0, , .StatusText + vbCrLf + .ResponseText
            End If
        End With
    Else
        'TODO: update
    End If
End Sub


