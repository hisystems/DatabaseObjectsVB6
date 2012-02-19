VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ___________________________________________________
'
'  © Hi-Integrity Systems 2007. All rights reserved.
'  www.hisystems.com.au - Toby Wicks
' ___________________________________________________
'

Option Explicit

Public Enum ConnectionTypeEnum
    'SQLServer and MicrosoftAccess will produce identical SQL code,
    'but are seperate in case this changes in future
    dboConnectionTypeSQLServer
    dboConnectionTypeMicrosoftAccess
    dboConnectionTypeMySQL
End Enum

Public Enum ErrorEnum
    dboErrorGeneral = vbObjectError + 1
    dboErrorIndexOutOfBounds = vbObjectError + 2
    dboErrorNotIntegerOrString = vbObjectError + 3
    dboErrorObjectIsNothing = vbObjectError + 4
    dboErrorObjectAlreadyExists = vbObjectError + 5
    dboErrorObjectDoesNotExist = vbObjectError + 6
    dboErrorInvalidPropertyValue = vbObjectError + 7
    dboErrorInvalidArgument = vbObjectError + 8
    dboErrorObjectNotDeletable = vbObjectError + 9
    dboErrorObjectNotSaved = vbObjectError + 10
    dboErrorNotSupported = vbObjectError + 11
    dboErrorMethodOrPropertyLocked = vbObjectError + 12
End Enum

Private pobjByOrdinalLastCollection As IDatabaseObjects
Private pobjConnection As ConnectionController
Private pobjTransactions As DatabaseTransactions

Public Sub Connect( _
    ByVal strConnectionString As String, _
    Optional ByVal eConnectionType As ConnectionTypeEnum = dboConnectionTypeSQLServer)

    Set pobjConnection = New ConnectionController
    pobjConnection.Initialize strConnectionString, eConnectionType
    
    Set pobjTransactions = New DatabaseTransactions
    pobjTransactions.Initialize pobjConnection
    
End Sub

'Useful when the DBO - global instance is being used
Public Property Get Instance() As Database

    Set Instance = Me

End Property

Public Property Get Connection() As ConnectionController

    Set Connection = pobjConnection

End Property

Public Property Get Transactions() As DatabaseTransactions

    Set Transactions = pobjTransactions

End Property

Public Function Object( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal vDistinctValue As Variant) As IDatabaseObject

    Dim objRecordset As adodb.Recordset
    Dim objSubset As SQLConditions
    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect
    
    With objSelect
        .Tables.Add objCollection.TableName
        Set .Tables.Joins = objCollection.TableJoins(.Tables(1), .Tables)
        .Where.Add objCollection.DistinctFieldName, dboComparisonEqualTo, vDistinctValue
        Set objSubset = objCollection.Subset
        If Not objSubset Is Nothing Then
            If Not objSubset.IsEmpty Then
                .Where.AddConditions objSubset
            End If
        End If
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    Set Object = ObjectFromRecordset(objCollection, objRecordset)
    
    objRecordset.Close
    pobjConnection.Finished
    
End Function

Public Function ObjectLockRecord( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal objItem As IDatabaseObject) As SQLFieldValues

    Set ObjectLockRecord = ObjectLockRecordByDistinctValue(objCollection, objItem.DistinctValue)
    
End Function

Public Function ObjectLockRecordByDistinctValue( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal vDistinctValue As Variant) As SQLFieldValues

    Dim objRecordset As adodb.Recordset
    Dim objSubset As SQLConditions
    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect
    
    With objSelect
        .PerformLocking = True
        .Tables.Add objCollection.TableName
        .Where.Add objCollection.DistinctFieldName, dboComparisonEqualTo, vDistinctValue
        Set objSubset = objCollection.Subset
        If Not objSubset Is Nothing Then
            If Not objSubset.IsEmpty Then
                .Where.AddConditions objSubset
            End If
        End If
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    Set ObjectLockRecordByDistinctValue = FieldValuesFromRecordSet(objCollection, objRecordset)
    objRecordset.Close
    
    pobjConnection.Finished
    
End Function

Public Sub ObjectSave( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal objItem As IDatabaseObject)
    
    Dim objRecordset As adodb.Recordset
    Dim objFieldValues As SQLFieldValues
    Dim objSubset As SQLConditions

    Set objFieldValues = objItem.Save
    If objFieldValues Is Nothing Then
        RaiseError dboErrorGeneral, TypeName(objItem) & " IDatabaseObject_Save not implemented"
    End If

    'Clear the ordinal collection if an object is being saved
    Set pobjByOrdinalLastCollection = Nothing

    'Add the distinct field value if it hasn't been added via the Save sub
    If Not objFieldValues.Exists(objCollection.DistinctFieldName) Then
        If Not objCollection.DistinctFieldAutoIncrements Then
            objFieldValues.Add objCollection.DistinctFieldName, objItem.DistinctValue
        End If
    End If
    
    ItemKeyEnsureValid objCollection, objItem, objFieldValues
    
    pobjConnection.Start
    
    If objItem.IsSaved Then
        Dim objUpdate As SQLUpdate
        Set objUpdate = New SQLUpdate
        
        objUpdate.TableName = objCollection.TableName
        Set objUpdate.Fields = objFieldValues
        objUpdate.Where.Add objCollection.DistinctFieldName, dboComparisonEqualTo, objItem.DistinctValue
        Set objSubset = objCollection.Subset
        If Not objSubset Is Nothing Then
            If Not objSubset.IsEmpty Then
                objUpdate.Where.AddConditions objSubset
            End If
        End If
        
        pobjConnection.Execute objUpdate, adCmdText Or adExecuteNoRecords
    Else
        Dim objIncrementField As SQLAutoIncrementValue
        Dim objInsert As SQLInsert
        
        Set objInsert = New SQLInsert
        objInsert.TableName = objCollection.TableName
        Set objInsert.Fields = objFieldValues
        pobjConnection.Execute objInsert, adCmdText Or adExecuteNoRecords
        
        'get the identity value for the newly inserted record
        If objCollection.DistinctFieldAutoIncrements Then
            Set objIncrementField = New SQLAutoIncrementValue
            Set objRecordset = pobjConnection.Execute(objIncrementField, adCmdText)
            objItem.DistinctValue = objRecordset.Fields(0)
            objRecordset.Close
        End If
                
        objItem.IsSaved = True
    End If
    
    pobjConnection.Finished

End Sub

Private Sub ItemKeyEnsureValid( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal objItem As IDatabaseObject, _
    ByVal objFieldValues As SQLFieldValues)
    
    Dim objRecordset As adodb.Recordset
    Dim objSelect As SQLSelect
    Dim vKeyFieldValue As Variant
    Dim objSubset As SQLConditions
    
    If objCollection.KeyFieldName <> vbNullString Then
        vKeyFieldValue = ItemKeyFieldValue(objCollection, objItem, objFieldValues)
        EnsureKeyDataTypeValid vKeyFieldValue
        
        If VarType(vKeyFieldValue) = vbString Then
            If Trim$(vKeyFieldValue) = vbNullString Then
                RaiseError dboErrorGeneral, TypeName(objItem) & " " & objCollection.KeyFieldName & " field is Null"
            End If
        End If
        
        Set objSelect = New SQLSelect
        
        With objSelect
            .Tables.Add objCollection.TableName
            .Fields.Add objCollection.KeyFieldName
            .Where.Add objCollection.KeyFieldName, dboComparisonEqualTo, vKeyFieldValue
            Set objSubset = objCollection.Subset
            If Not objSubset Is Nothing Then
                If Not objSubset.IsEmpty Then
                    .Where.AddConditions objSubset
                End If
            End If
        
            If objItem.IsSaved Then
                .Where.Add objCollection.DistinctFieldName, dboComparisonNotEqualTo, objItem.DistinctValue
            End If
        End With
        
        pobjConnection.Start
        
        Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
        
        If Not objRecordset.EOF Then
            RaiseError dboErrorObjectAlreadyExists, TypeName(objItem) & " '" & vKeyFieldValue & "'"
        End If
        
        objRecordset.Close
        pobjConnection.Finished
    End If
    
End Sub

Public Function ObjectItem( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal vIndex As Variant) As IDatabaseObject

    Dim objItem As IDatabaseObject

    Select Case VarType(vIndex)
    Case vbInteger, vbLong
        Set objItem = ObjectByOrdinal(objCollection, vIndex)
    Case Else
        Set objItem = ObjectByKey(objCollection, vIndex)
    End Select
    
    Set ObjectItem = objItem

End Function

'Gets and returns the field value from the database record associated with the object and collection.
Public Function ObjectGetFieldValue( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal objItem As IDatabaseObject, _
    ByVal strFieldName As String) As Variant

    If Not objItem.IsSaved Then
        RaiseError dboErrorObjectNotSaved
    End If

    Dim objSelect As SQLSelect
    Dim objSubset As SQLConditions
    Set objSelect = New SQLSelect

    With objSelect
        .Fields.Add strFieldName
        .Tables.Add objCollection.TableName
        .Where.Add objCollection.DistinctFieldName, dboComparisonEqualTo, objItem.DistinctValue
        Set objSubset = objCollection.Subset
        If Not objSubset Is Nothing Then
            If Not objSubset.IsEmpty Then
                .Where.AddConditions objSubset
            End If
        End If
    End With

    Me.Connection.Start
    
    Dim objRecordset As adodb.Recordset
    Set objRecordset = Me.Connection.Execute(objSelect, adCmdText)

    If Not objRecordset.EOF Then
        ObjectGetFieldValue = objRecordset(0).Value
    Else
        RaiseError dboErrorObjectDoesNotExist, "DistinctValue: " & objItem.DistinctValue
    End If

    objRecordset.Close
    Me.Connection.Finished

End Function

'Sets the field value from the database record associated with the object and collection.
Public Sub ObjectSetFieldValue( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal objItem As IDatabaseObject, _
    ByVal strFieldName As String, _
    ByVal vNewValue As Variant)

    If Not objItem.IsSaved Then
        RaiseError dboErrorObjectNotSaved
    End If

    Dim objSubset As SQLConditions
    Dim objUpdate As SQLUpdate
    
    Set objUpdate = New SQLUpdate
    
    objUpdate.TableName = objCollection.TableName
    objUpdate.Fields.Add strFieldName, vNewValue
    objUpdate.Where.Add objCollection.DistinctFieldName, dboComparisonEqualTo, objItem.DistinctValue
    
    Set objSubset = objCollection.Subset
    
    If Not objSubset Is Nothing Then
        If Not objSubset.IsEmpty Then
            objUpdate.Where.AddConditions objSubset
        End If
    End If

    Me.Connection.Start
    Me.Connection.Execute objUpdate, adCmdText
    Me.Connection.Finished

End Sub
    
Public Function ObjectByKey( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal vKey As Variant) As IDatabaseObject
    
    Dim objRecordset As adodb.Recordset
    Dim objSubset As SQLConditions
    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect
    
    EnsureKeyDataTypeValid vKey

    With objSelect
        .Tables.Add objCollection.TableName
        Set .Tables.Joins = objCollection.TableJoins(.Tables(1), .Tables)
        .Where.Add objCollection.KeyFieldName, dboComparisonEqualTo, vKey
        Set objSubset = objCollection.Subset
        If Not objSubset Is Nothing Then
            If Not objSubset.IsEmpty Then
                .Where.AddConditions objSubset
            End If
        End If
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    
    If objRecordset.EOF Then
        RaiseError dboErrorObjectDoesNotExist, TypeName(objCollection) & " Key: '" & vKey & "'"
    Else
        Set ObjectByKey = ObjectFromRecordset(objCollection, objRecordset)
    End If
    
    objRecordset.Close
    pobjConnection.Finished
    
End Function

Public Function ObjectByOrdinalFirst( _
    ByVal objCollection As IDatabaseObjects) As IDatabaseObject

    Dim objRecordset As adodb.Recordset
    Dim objSelect As SQLSelect
    Dim objSubset As SQLConditions
    
    Set objSelect = New SQLSelect
    
    With objSelect
        'only select the first row of the recordset
        .Top = 1
        .Tables.Add objCollection.TableName
        Set .Tables.Joins = objCollection.TableJoins(.Tables(1), .Tables)
        Set .Where = objCollection.Subset
        Set .OrderBy = objCollection.OrderBy
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    Set ObjectByOrdinalFirst = ObjectFromRecordset(objCollection, objRecordset)
    
    objRecordset.Close
    pobjConnection.Finished
    
End Function


Public Function ObjectByOrdinal( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal lngIndex As Long) As IDatabaseObject

    Static objObjects As Collection
    
    If Not pobjByOrdinalLastCollection Is objCollection Then
        Set objObjects = ObjectsCollection(objCollection)
        Set pobjByOrdinalLastCollection = objCollection
    End If
    
    Set ObjectByOrdinal = objObjects(lngIndex)

End Function

Public Function ObjectsCollection( _
    ByVal objCollection As IDatabaseObjects) As Collection
    
    'Returns all of the objects in the collection and returns a Collection object
    'This is useful when implementing the For Each command

    Dim objRecordset As adodb.Recordset
    Dim objSelect As SQLSelect
    Dim objObjects As Collection
    
    Set objObjects = New Collection
    Set objSelect = New SQLSelect
    
    With objSelect
        .Tables.Add objCollection.TableName
        Set .Tables.Joins = objCollection.TableJoins(.Tables(1), .Tables)
        Set .Where = objCollection.Subset
        Set .OrderBy = objCollection.OrderBy
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    
    While Not objRecordset.EOF
        objObjects.Add ObjectFromRecordset(objCollection, objRecordset)
        objRecordset.MoveNext
    Wend

    Set ObjectsCollection = objObjects
    
    objRecordset.Close
    pobjConnection.Finished
    
End Function

Public Function ObjectExists( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal vKey As Variant) As Boolean

    Dim objRecordset As adodb.Recordset
    Dim objSelect As SQLSelect
    Dim objSubset As SQLConditions
    
    Set objSelect = New SQLSelect
    
    EnsureKeyDataTypeValid vKey

    With objSelect
        .Tables.Add objCollection.TableName
        '.Fields.Add objCollection.DistinctFieldName
        .Where.Add objCollection.KeyFieldName, dboComparisonEqualTo, vKey
        Set .Tables.Joins = objCollection.TableJoins(.Tables(1), .Tables)
        Set objSubset = objCollection.Subset
        If Not objSubset Is Nothing Then
            If Not objSubset.IsEmpty Then
                .Where.AddConditions objSubset
            End If
        End If
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    
    ObjectExists = Not objRecordset.EOF
    
    objRecordset.Close
    pobjConnection.Finished

End Function

Public Function ObjectExistsByDistinctValue( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal vDistinctValue As Variant) As Boolean

    Dim objRecordset As adodb.Recordset
    Dim objSelect As SQLSelect
    Dim objSubset As SQLConditions
    
    Set objSelect = New SQLSelect
    
    EnsureKeyDataTypeValid vDistinctValue

    With objSelect
        .Tables.Add objCollection.TableName
        .Where.Add objCollection.DistinctFieldName, dboComparisonEqualTo, vDistinctValue
        Set objSubset = objCollection.Subset
        If Not objSubset Is Nothing Then
            If Not objSubset.IsEmpty Then
                .Where.AddConditions objSubset
            End If
        End If
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    
    ObjectExistsByDistinctValue = Not objRecordset.EOF
    
    objRecordset.Close
    pobjConnection.Finished

End Function

Public Sub ObjectDelete( _
    ByVal objCollection As IDatabaseObjects, _
    ByRef objItem As IDatabaseObject)
    
    Dim objDelete As SQLDelete
    Dim objSubset As SQLConditions
    
    Set objDelete = New SQLDelete
    
    If objItem.IsSaved Then
        'Clear the ordinal collection if an object is being saved
        Set pobjByOrdinalLastCollection = Nothing
        
        With objDelete
            .TableName = objCollection.TableName
            .Where.Add objCollection.DistinctFieldName, dboComparisonEqualTo, objItem.DistinctValue
            Set objSubset = objCollection.Subset
            If Not objSubset Is Nothing Then
                If Not objSubset.IsEmpty Then
                    .Where.AddConditions objSubset
                End If
            End If
        End With
        
        pobjConnection.Start
        pobjConnection.Execute objDelete, adCmdText Or adExecuteNoRecords
        pobjConnection.Finished
        
        objItem.IsSaved = False
    End If
    
    'delete the instance
    'for the function that calls ObjectDelete objItem MUST be ByRef for this to have any effect
    Set objItem = Nothing

End Sub

Public Sub ObjectsDeleteAll( _
    ByVal objCollection As IDatabaseObjects)
    
    Dim objDelete As SQLDelete
    Set objDelete = New SQLDelete
    
    'Clear the ordinal collection if an object is being saved
    Set pobjByOrdinalLastCollection = Nothing

    With objDelete
        .TableName = objCollection.TableName
        Set .Where = objCollection.Subset
    End With
    
    pobjConnection.Start
    pobjConnection.Execute objDelete, adCmdText Or adExecuteNoRecords
    pobjConnection.Finished
    
End Sub

Public Function ObjectsCount( _
    ByVal objCollection As IDatabaseObjects) As Long
    
    Dim objRecordset As adodb.Recordset
    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect
        
    With objSelect
        .Tables.Add objCollection.TableName
        .Fields.Add , , dboAggregateCount
        Set .Where = objCollection.Subset
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    ObjectsCount = ConvertNullToLong(objRecordset.Fields(0))

    objRecordset.Close
    pobjConnection.Finished

End Function

Public Function ObjectsSearch( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal objSearchCriteria As SQLConditions) As Collection
    
    Dim objRecordset As adodb.Recordset
    Dim objSelect As SQLSelect
    Dim colResults As Collection
    
    If objSearchCriteria.IsEmpty Then
        RaiseError dboErrorInvalidArgument, "Search criteria is empty"
    End If
    
    Set colResults = New Collection
    Set objSelect = New SQLSelect
    
    With objSelect
        .Tables.Add objCollection.TableName
        Set .Tables.Joins = objCollection.TableJoins(.Tables(1), .Tables)
        Set .OrderBy = objCollection.OrderBy
        Set .Where = objCollection.Subset
        
        If Not objSearchCriteria Is Nothing Then
            If .Where Is Nothing Then Set .Where = New SQLConditions
            .Where.AddConditions objSearchCriteria
        End If
    End With
    
    pobjConnection.Start
    
    Set objRecordset = pobjConnection.Execute(objSelect, adCmdText)
    
    While Not objRecordset.EOF
        colResults.Add ObjectFromRecordset(objCollection, objRecordset)
        objRecordset.MoveNext
    Wend
    
    Set ObjectsSearch = colResults
    
    objRecordset.Close
    pobjConnection.Finished
    
End Function

Public Function ObjectDistinctValue( _
    ByVal objItem As IDatabaseObject) As Variant
    
    ObjectDistinctValue = objItem.DistinctValue
    
End Function

Public Function ObjectIsSaved( _
    ByVal objItem As IDatabaseObject) As Boolean
    
    ObjectIsSaved = objItem.IsSaved
    
End Function

Public Sub RaiseError( _
    ByVal eError As ErrorEnum, _
    Optional ByVal strExtra As String)
    
    modMisc.RaiseError eError, strExtra

End Sub

Public Function ObjectFromRecordset( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal rsSource As Recordset) As IDatabaseObject
    
    Set ObjectFromRecordset = ObjectFromFieldValues(objCollection, FieldValuesFromRecordSet(objCollection, rsSource))

End Function

Private Function FieldValuesFromRecordSet( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal rsSource As Recordset) As SQLFieldValues

    Dim intIndex As Integer
    Dim strTablePrefix As String
    Dim objFieldValues As SQLFieldValues
     
    'check that the distinct field name exists
    If Not FieldExists(rsSource, objCollection.DistinctFieldName) Then
        RaiseError dboErrorGeneral, "DistinctFieldName '" & objCollection.DistinctFieldName & "' is invalid."
    End If

    If rsSource.EOF Then
        RaiseError dboErrorGeneral, "Object could not be loaded from the RecordSet. Recordset is EOF."
    End If

    Set objFieldValues = New SQLFieldValues
    strTablePrefix = objCollection.TableName & "."

    'copy the recordset values into the SQLFieldValues object
    For intIndex = 0 To rsSource.Fields.Count - 1
        With rsSource.Fields(intIndex)
            'if the recordset has been loaded with a join then it may be prefixed with
            'the table name - this is the case with Microsoft Access
            'If so remove the table name if the table prefix is the same as objCollection.TableName
            'All of the other joined field with tablename prefixes on the fields will remain. This is fine considering
            'most of the time an inner join has been performed where the fields are equal in the 2 joined tables
            If InStr(1, .Name, strTablePrefix, vbTextCompare) = 1 Then
                objFieldValues.Add Mid$(.Name, Len(strTablePrefix) + 1), .Value
            Else
                objFieldValues.Add .Name, .Value
            End If
        End With
    Next

    Set FieldValuesFromRecordSet = objFieldValues

End Function

Public Function ObjectFromFieldValues( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal objFieldValues As SQLFieldValues) As IDatabaseObject
    
    Dim objItem As IDatabaseObject
    Set objItem = objCollection.ItemInstance
    
    objItem.DistinctValue = objFieldValues(objCollection.DistinctFieldName).Value
    objItem.Load objFieldValues
    objItem.IsSaved = True
    
    Set ObjectFromFieldValues = objItem
    
End Function

Private Function FieldExists( _
    ByVal rsSource As Recordset, _
    ByVal strFieldName As String) As Boolean
    
    Dim intIndex As Integer
    Dim bExists As Boolean
    
    'copy the recordset values into the
    For intIndex = 0 To rsSource.Fields.Count - 1
        If StrComp(rsSource.Fields(intIndex).Name, strFieldName, vbTextCompare) = 0 Then
            bExists = True
            Exit For
        End If
    Next
    
    FieldExists = bExists
    
End Function

Private Sub EnsureKeyDataTypeValid( _
    ByRef vKey As Variant)
    
    Select Case VarType(vKey)
    Case vbBoolean, vbByte, vbDate, vbDouble, vbInteger, vbLong, vbSingle, vbString
    Case Else
        RaiseError dboErrorGeneral, "Invalid data type."
    End Select

End Sub

Private Function ItemKeyFieldValue( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal objItem As IDatabaseObject, _
    ByVal objFieldValues As SQLFieldValues) As Variant

    'on the rare occurance that the KeyField is the same as the DistinctField
    'then the key value may not have been set in the Save and therefore be
    'available in the objFieldValues collection. In which case the
    'key has to be extracted from the objItem.DistinctField.
    Dim vKeyFieldValue As Variant
    
    If StrComp(objCollection.DistinctFieldName, objCollection.KeyFieldName, vbTextCompare) = 0 Then
        vKeyFieldValue = objItem.DistinctValue
    Else
        vKeyFieldValue = objFieldValues(objCollection.KeyFieldName)
    End If
    
    ItemKeyFieldValue = vKeyFieldValue
    
End Function

Private Function ConvertNullToLong( _
    ByVal vValue As Variant) As Long

    If IsNull(vValue) Then
        ConvertNullToLong = 0
    Else
        ConvertNullToLong = vValue
    End If

End Function
