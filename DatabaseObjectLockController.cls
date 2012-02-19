VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseObjectLockController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' ___________________________________________________
'
'  © Hi-Integrity Systems 2007. All rights reserved.
'  www.hisystems.com.au - Toby Wicks
' ___________________________________________________
'


''' --------------------------------------------------------------------------------
''' <summary>
''' This is the controller class that initializes the lock table and the user
''' ID that is to be associated with all locking operations. In most situations,
''' only one instance of this class is ever created and this instance is passed into
''' the constructor for all DatabaseObjectLockable and DatabaseObjectUsingAttributesLockable
''' instances.
''' </summary>
''' --------------------------------------------------------------------------------
'''
    
Option Explicit
    
Private pstrCurrentUserID As String
Private pstrLockTableName As String
Private pobjConnection As ConnectionController

Public Sub Initialize(ByVal objDatabase As Database, ByVal strLockTableName As String, ByVal strCurrentUserID As String)

    If objDatabase Is Nothing Then
        RaiseError dboErrorObjectIsNothing, "Database"
    ElseIf strCurrentUserID = vbNullString Then
        RaiseError dboErrorInvalidArgument, "UserID is null"
    ElseIf strLockTableName = vbNullString Then
        RaiseError dboErrorInvalidArgument, "LockTableName is null"
    End If

    Set pobjConnection = objDatabase.Connection
    pstrCurrentUserID = strCurrentUserID
    pstrLockTableName = strLockTableName

    EnsureTableExists

End Sub

Private Function EnsureTableExists() As Boolean

    Dim bTableExists As Boolean
    Dim objTableExists As SQLTableExists
    
    Set objTableExists = New SQLTableExists
    objTableExists.Name = pstrLockTableName

    pobjConnection.Start

    With pobjConnection.Execute(objTableExists, adCmdText)
        bTableExists = Not .EOF
        .Close
    End With

    If Not bTableExists Then
        pobjConnection.Execute CreateTable, adCmdText
        pobjConnection.Execute CreateTableIndex, adCmdText
    End If

    pobjConnection.Finished

End Function

Public Property Get IsLocked(ByVal objCollection As IDatabaseObjects, ByVal objObject As IDatabaseObject) As Boolean
        
    IsLocked = LockRecordExists(objCollection.TableName, objObject)

End Property

Public Property Get IsLockedByCurrentUser(ByVal objCollection As IDatabaseObjects, ByVal objObject As IDatabaseObject) As Boolean

    Dim objFilter As SQLCondition
    Set objFilter = New SQLCondition
    
    objFilter.FieldName = "UserID"
    objFilter.compare = dboComparisonEqualTo
    objFilter.Value = pstrCurrentUserID

    IsLockedByCurrentUser = LockRecordExists(objCollection.TableName, objObject, objFilter)

End Property

Public Property Get LockedByUserID(ByVal objCollection As IDatabaseObjects, ByVal objObject As IDatabaseObject) As String
'
'    If Not Me.IsLocked(objCollection, objObject) Then
'        RaiseError dboErrorGeneral, "Object is not locked"
'    End If

    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect

    objSelect.Fields.Add "UserID"
    objSelect.Tables.Add pstrLockTableName
    objSelect.Where.Add "TableName", dboComparisonEqualTo, objCollection.TableName
    objSelect.Where.Add "RecordID", dboComparisonEqualTo, CStr(objObject.DistinctValue)

    pobjConnection.Start
    
    Dim objRecordset As adodb.Recordset
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    
    If Not objRecordset.EOF Then
        LockedByUserID = CStr(objRecordset(0))
    Else
        RaiseError dboErrorGeneral, "Object is not locked"
    End If

    objRecordset.Close
    pobjConnection.Finished

End Property

Private Function LockRecordExists( _
    ByVal strTableName As String, _
    ByVal objObject As IDatabaseObject, _
    Optional ByVal objAdditionalCondition As SQLCondition = Nothing) As Boolean

    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect

    objSelect.Fields.Add vbNullString, eAggregateFunction:=dboAggregateCount
    objSelect.Tables.Add pstrLockTableName
    objSelect.Where.Add "TableName", dboComparisonEqualTo, strTableName
    objSelect.Where.Add "RecordID", dboComparisonEqualTo, CStr(objObject.DistinctValue)
    If Not objAdditionalCondition Is Nothing Then
        objSelect.Where.AddCondition objAdditionalCondition
    End If

    pobjConnection.Start
    
    Dim objRecordset As adodb.Recordset
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)

    LockRecordExists = CInt(objRecordset(0)) <> 0

    objRecordset.Close
    pobjConnection.Finished

End Function

Public Function Lock_(ByVal objCollection As IDatabaseObjects, ByVal objObject As IDatabaseObject) As Boolean
 
    If Not objObject.IsSaved Then
        RaiseError dboErrorInvalidArgument, "Object is not saved and cannot be locked"
    End If
 
    Dim objInsert As SQLInsert
    Set objInsert = New SQLInsert
    objInsert.TableName = pstrLockTableName
    objInsert.Fields.Add "TableName", objCollection.TableName
    objInsert.Fields.Add "RecordID", CStr(objObject.DistinctValue)
    objInsert.Fields.Add "UserID", pstrCurrentUserID
 
    pobjConnection.Start
 
    On Error Resume Next
    
    'If another user/connection has managed to add a record to the database just before
    'this connection has a error will be raised because duplicate keys will
    'be added to the table.
    pobjConnection.Execute objInsert, adCmdText
    
    Dim bSuccess As Boolean
    bSuccess = Err.Number = 0
    On Error GoTo 0
    
    pobjConnection.Finished
 
    Lock_ = bSuccess
 
End Function
 
Public Sub UnLock_(ByVal objCollection As IDatabaseObjects, ByVal objObject As IDatabaseObject)
 
    'If the table is locked by someone else
    If Not Me.IsLockedByCurrentUser(objCollection, objObject) Then
        RaiseError dboErrorGeneral, "Object locked by another user"
    ElseIf Not objObject.IsSaved Then
        RaiseError dboErrorGeneral, "Object is not saved and cannot be unlocked"
    End If
 
    Dim objDelete As SQLDelete
    Set objDelete = New SQLDelete
    objDelete.TableName = pstrLockTableName
    objDelete.Where.Add "TableName", dboComparisonEqualTo, objCollection.TableName
    objDelete.Where.Add "RecordID", dboComparisonEqualTo, CStr(objObject.DistinctValue)
    objDelete.Where.Add "UserID", dboComparisonEqualTo, pstrCurrentUserID
 
    pobjConnection.Start
    pobjConnection.Execute objDelete, adCmdText
    pobjConnection.Finished
 
End Sub
 

''' --------------------------------------------------------------------------------
''' <summary>
''' Provides a means by which to ensure all locks have been removed for this user
''' in situations where an unexpected exception occurs and/or the user logs out of
''' system.
''' </summary>
''' --------------------------------------------------------------------------------
Public Sub UnlockAll()

    Dim objDelete As SQLDelete
    Set objDelete = New SQLDelete
    objDelete.TableName = pstrLockTableName
    objDelete.Where.Add "UserID", dboComparisonEqualTo, pstrCurrentUserID

    pobjConnection.Start
    pobjConnection.Execute objDelete, adCmdText
    pobjConnection.Finished

End Sub

Private Function CreateTable() As ISQLStatement

    Dim objTable As SQLCreateTable
    Set objTable = New SQLCreateTable

    objTable.Name = pstrLockTableName
    objTable.Fields.Add "TableName", dboDataTypeVariableCharacter, 50
    objTable.Fields.Add "RecordID", dboDataTypeVariableCharacter, 20
    objTable.Fields.Add "UserID", dboDataTypeVariableCharacter, 255  'Accounts for windows user names

    Set CreateTable = objTable

End Function

Private Function CreateTableIndex() As ISQLStatement

    Dim objIndex As SQLCreateIndex
    Set objIndex = New SQLCreateIndex

    objIndex.Name = "Primary"
    objIndex.IsUnique = True
    objIndex.TableName = pstrLockTableName
    objIndex.Fields.Add "TableName"
    objIndex.Fields.Add "RecordID"
    objIndex.Fields.Add "UserID"

    Set CreateTableIndex = objIndex

End Function

