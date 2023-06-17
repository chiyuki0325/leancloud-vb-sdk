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

Public Enum AtomicOperationTypes
    aoDelete = 0
    aoAdd = 1
    aoAddUnique = 2
    aoRemove = 3
    aoIncrement = 4
    aoDecrement = 5
    aoBitAnd = 6
    aoBitOr = 7
    aoBitXor = 8
End Enum

Private AtomicOperationTypeStrings()

Private Utils As cLeanCloudUtils
Private Dict As New Dictionary

Public UpdateMode As Boolean  'Record updated fields
Private UpdatedFields As New Collection

Public ClassName As String, ObjectId As String
Public CreatedAt As Date, UpdatedAt As Date

Public Sub Initialize( _
    ByRef oUtils As cLeanCloudUtils, _
    ByRef sClassName As String, _
    Optional ByRef sObjectId As String _
)
    Set Utils = oUtils
    ClassName = sClassName
    ObjectId = sObjectId
    
    AtomicOperationTypeStrings = Array("Delete", "Add", "AddUnique", "Remove", "Increment", "Decrement", "BitAnd", "BitOr", "BitXor")
End Sub

Public Property Get Item(Key$)
    If VarType(Dict.Item(Key)) = vbObject Then
        Set Item = Dict.Item(Key)
    Else
        Item = Dict.Item(Key)
    End If
End Property

Public Property Let Item(Key$, Value)
    If Dict.Exists(Key) Then
        Dict.Item(Key) = Value
    Else
        Dict.Add Key, Value
    End If
    If UpdateMode Then UpdatedFields.Add Key
End Property

Public Property Set Item(Key$, Value)
    If Dict.Exists(Key) Then
        Set Dict.Item(Key) = Value
    Else
        Dict.Add Key, Value
    End If
    If UpdateMode Then UpdatedFields.Add Key
End Property

'If need to 'Set'
Private Function IsObjectType(Var) As Boolean
    Select Case VarType(Var)
        Case vbInteger, vbBoolean, vbString, vbLong, vbSingle, vbDouble
            IsObjectType = False
        Case vbDate, vbArray + vbByte, vbArray + vbString, vbArray + vbInteger, vbArray + vbLong, vbArray + vbBoolean, vbArray + vbSingle, vbArray + vbDouble
            IsObjectType = True
        Case Else
            Select Case TypeName(Var)
                Case "LCObject" 'Pointer
                'TODO: LCFile
                    IsObjectType = True
                Case Else
                    Err.Raise 13, , TypeName(Var)
            End Select
    End Select
End Function

Private Function ToLeanObject(Var) As Object  'Collection or Dictionary
    Dim OutputDict As New Dictionary
    Select Case TypeName(Var)
        Case "Date"
            With OutputDict
                .Add "__type", "Date"
                .Add "iso", Utils.ConvertDateToISO(Var)
            End With
            Set ToLeanObject = OutputDict
        Case "Byte()"
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
            With OutputDict
                .Add "__type", "Pointer"
                .Add "className", Var.ClassName
                .Add "objectId", Var.ObjectId
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
        If Not ObjectId = "" Then .Add "objectId", ObjectId
        If Not CreatedAt = "0:00.00" Then .Add "createdAt", Utils.ConvertDateToISO(CreatedAt)
        If Not UpdatedAt = "0:00.00" Then .Add "updatedAt", Utils.ConvertDateToISO(UpdatedAt)
    End With
End Function

Public Sub Deserialize(InputDict As Dictionary)
    Dim Key
    If InputDict.Exists("objectId") Then
        ObjectId = InputDict("objectId")
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
                Set Me(CStr(Key)) = ToVBObject(InputDict(Key))
            Else
                Me(CStr(Key)) = ToVBObject(InputDict(Key))
            End If
        Else
            Me(CStr(Key)) = InputDict(Key)
        End If
    Next
    UpdateMode = True
End Sub

Public Sub Save()
    Dim JSON As New cJSON
    If ObjectId = "" Then
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
                    ObjectId = Response("objectId")
                    CreatedAt = Utils.ConvertISOToDate(Response("createdAt"))
                    UpdateMode = True
                End If
            Else
                Err.Raise 0, , .StatusText + vbCrLf + .ResponseText
            End If
        End With
    Else
        'Update
        Dim DeltaObject As New LCObject, Key
        For Each Key In UpdatedFields
            If VarType(Me(Key)) = vbObject Then
                Set DeltaObject(Key) = Me(Key)
            Else
                DeltaObject(Key) = Me(Key)
            End If
        Next
        With Utils.NewRequest("PUT", "/classes/" + ClassName + "/" + ObjectId)
            .Send JSON.ToString(DeltaObject.Serialize)
            .WaitForResponse
            If .Status = 200 Then
                UpdatedAt = Utils.ConvertISOToDate(JSON.Parse(.ResponseText)("update"))
            Else
                Err.Raise 0, , .StatusText + vbCrLf + .ResponseText
            End If
        End With
    End If
End Sub

Private Sub AtomicOperation( _
    Key$, _
    AtomicOperationType As AtomicOperationTypes, _
    DictOp As Dictionary _
)
    Dim Dict As New Dictionary, DictOp2 As New Dictionary
    Dim Key2
    With DictOp2
        .Add "__op", AtomicOperationTypeStrings(AtomicOperationType)
        For Each Key2 In DictOp
            .Add Key2, DictOp(Key2)
        Next
    End With
    Dict.Add Key, DictOp2
    With Utils.NewRequest("PUT", "/classes/" + ClassName + "/" + ObjectId)
        .Send JSON.ToString(Dict)
        .WaitForResponse
        If .Status = 200 Then
            UpdatedAt = Utils.ConvertISOToDate(JSON.Parse(.ResponseText)("update"))
        Else
            Err.Raise 0, , .StatusText + vbCrLf + .ResponseText
        End If
    End With
End Sub

Public Sub Increase(Key$, IncreaseBy)
    Select Case VarType(Me(Key))
        Case vbInteger, vbLong, vbSingle, vbDouble
            Me(Key) = Me(Key) + IncreaseBy
            Dim DictOp As New Dictionary
            With DictOp
                .Add "amount", IncreaseBy
            End With
            AtomicOperation Key, aoIncrement, DictOp
        Case Else
            Err.Raise 13
    End Select
End Sub

Public Sub AppendTo(Key$, ItemsToAdd())
    If TypeName(Me(Key)) = "Collection" Then
        Dim ColCurrent As Collection
        Set ColCurrent = Me(Key)
        Dim v
        Dim DictOp As New Dictionary, ColObjects As New Collection
        For Each v In ItemsToAdd
            ColCurrent.Add v
            ColObjects.Add v
        Next
        DictOp.Add "objects", ColObjects
        AtomicOperation Key, aoAdd, DictOp
    Else
        Err.Raise 13
    End If
End Sub

Public Sub AppendUniqueTo(Key$, ItemsToAdd())
    If TypeName(Me(Key)) = "Collection" Then
        Dim ColCurrent As Collection
        Set ColCurrent = Me(Key)
        Dim v, w
        Dim DictOp As New Dictionary, ColObjects As New Collection
        Dim Exists As Boolean
        For Each v In ItemsToAdd
            Exists = False
            For Each w In ColCurrent
                If w = v Then
                    Exists = True
                    Exit For
                End If
            Next
            If Not Exists Then ColCurrent.Add v
            
            ColObjects.Add v
        Next
        DictOp.Add "objects", ColObjects
        AtomicOperation Key, aoAddUnique, DictOp
    Else
        Err.Raise 13
    End If
End Sub


Public Sub RemoveFrom(Key$, ItemsToRemove())
    If TypeName(Me(Key)) = "Collection" Then
        Dim ColCurrent As Collection
        Set ColCurrent = Me(Key)
        Dim v
        Dim i%
        Dim DictOp As New Dictionary, ColObjects As New Collection
        For Each v In ItemsToRemove
            For i = 1 To ColCurrent.Count
                If ColCurrent(i) = v Then
                    ColCurrent.Remove i
                    Exit For
                End If
            Next
            ColObjects.Add v
        Next
        DictOp.Add "objects", ColObjects
        AtomicOperation Key, aoRemove, DictOp
    Else
        Err.Raise 13
    End If
End Sub


Public Sub Remove(Key$)
    If Dict.Exists(Key) Then
        Dict.Remove Key
        Dim DictOp As New Dictionary
        DictOp.Add "delete", True
        AtomicOperation Key, aoDelete, DictOp
    Else
        Err.Raise 9
    End If
End Sub


Public Sub Fetch()
    Dim JSON As New cJSON
    Dim Query As New LCQuery
    Query.Initialize Utils, ClassName
    Me.Deserialize Query.GetObject(ObjectId).Serialize
    Set UpdatedFields = New Collection
End Sub



