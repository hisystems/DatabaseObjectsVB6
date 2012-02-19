Attribute VB_Name = "modMisc"
' ___________________________________________________
'
'  © Hi-Integrity Systems 2007. All rights reserved.
'  www.hisystems.com.au - Toby Wicks
' ___________________________________________________
'

Option Explicit

Public Enum SQLTableFieldsAlterModeEnum
    dboTableFieldsModeAdd
    dboTableFieldsModeAlter
    dboTableFieldsModeDrop
End Enum

'This is used as the default connection type when manually instantiating
'a SQLSelect, SQLDelete, SQLUpdate or SQLInsert command and is set by the last
'Database.Connect function call's connectiontype argument
Public ConnectionType As ConnectionTypeEnum

Public Function SQLConvertIdentifierName( _
    ByVal strIdentifierName As String, _
    ByVal eConnectionType As ConnectionTypeEnum) As String

    'This function places tags around a field name or table name to ensure it doesn't
    'conflict with a reserved word or if it contains spaces it is not misinterpreted

    Select Case eConnectionType
        Case dboConnectionTypeMicrosoftAccess, dboConnectionTypeSQLServer
            SQLConvertIdentifierName = "[" & Trim$(strIdentifierName) & "]"
        Case dboConnectionTypeMySQL
            SQLConvertIdentifierName = "`" & Trim$(strIdentifierName) & "`"
    End Select
    
End Function

Public Function SQLConvertAggregate( _
    ByVal eAggregate As SQLAggregateFunctionEnum) As String
    
    Dim strAggregate As String
    
    Select Case eAggregate
        Case dboAggregateAverage
            strAggregate = "AVG"
        Case dboAggregateCount
            strAggregate = "COUNT"
        Case dboAggregateMaximum
            strAggregate = "MAX"
        Case dboAggregateMinimum
            strAggregate = "MIN"
        Case dboAggregateStandardDeviation
            strAggregate = "STDEV"
        Case dboAggregateSum
            strAggregate = "SUM"
        Case dboAggregateVariance
            strAggregate = "VAR"
    End Select
    
    SQLConvertAggregate = strAggregate

End Function

Public Function SQLConvertValue( _
    ByVal vValue As Variant, _
    ByVal eConnectionType As ConnectionTypeEnum) As String
    
    Dim strValue As String
    
    If SQLValueIsNull(vValue) Then
        strValue = "NULL"
    Else
        Select Case VarType(vValue)
            Case vbByte, vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                strValue = vValue
            Case vbString
                Select Case eConnectionType
                    Case dboConnectionTypeMicrosoftAccess, dboConnectionTypeSQLServer
                        strValue = "'" & Replace$(vValue, "'", "''") & "'"
                    Case dboConnectionTypeMySQL
                        strValue = "'" & Replace$(Replace$(vValue, "\", "\\"), "'", "\'") & "'"
                    Case Else
                        RaiseError dboErrorNotSupported, "Connection Type: " & eConnectionType
                End Select
            Case vbDate
                strValue = Year(vValue) & "-" & Month(vValue) & "-" & Day(vValue)
                
                If Hour(vValue) <> 0 Or Minute(vValue) <> 0 Or Second(vValue) <> 0 Then
                    strValue = strValue & " " & Hour(vValue) & ":" & Minute(vValue) & ":" & Second(vValue)
                End If
                
                If eConnectionType = dboConnectionTypeMicrosoftAccess Then
                    strValue = "#" & strValue & "#"
                Else
                    strValue = "'" & strValue & "'"
                End If
            Case vbBoolean
                strValue = IIf(vValue, "1", "0")
                
            Case (vbByte Or vbArray)
                Dim bytData() As Byte
                bytData = vValue
                strValue = SQLConvertByteArray(bytData)
            Case Else
                RaiseError dboErrorGeneral, "Invalid Variant Datatype: " & vValue
        End Select
    End If
    
    SQLConvertValue = strValue
    
End Function
 
Private Function SQLConvertByteArray(ByRef bytData() As Byte) As String
 
    Dim lngIndex As Long
    Dim objHexString As StringBuilder
    
    Set objHexString = New StringBuilder
    'Make it so that only 1 chunk is allocated as we know the final size of the string
    objHexString.ChunkSize = (UBound(bytData) - LBound(bytData) + 1) * 2 + 2
    objHexString.Append "0x"
    
    For lngIndex = LBound(bytData) To UBound(bytData)
        objHexString.Append SQLConvertByteToHex(bytData(lngIndex))
    Next
 
    SQLConvertByteArray = objHexString.Value
 
End Function
 
Private Function SQLConvertByteToHex(ByVal bytData As Byte) As String
 
    Dim strValue As String
    
    strValue = Hex$(bytData)
    
    If Len(strValue) = 1 Then
        strValue = "0" & strValue
    End If
    
    SQLConvertByteToHex = strValue
 
End Function

Public Function SQLValueIsNull(ByVal vValue As Variant) As Boolean

    Select Case VarType(vValue)
        Case vbObject
            SQLValueIsNull = vValue Is Nothing
        Case vbNull
            SQLValueIsNull = True
    End Select
    
End Function

Public Function SQLConvertCompare( _
    ByVal eCompare As SQLComparisonOperatorEnum) As String
    
    Dim strCompare As String
    
    Select Case eCompare
        Case dboComparisonEqualTo
            strCompare = "="
        Case dboComparisonGreaterThan
            strCompare = ">"
        Case dboComparisonGreaterThanOrEqualTo
            strCompare = ">="
        Case dboComparisonLessThan
            strCompare = "<"
        Case dboComparisonLessThanOrEqualTo
            strCompare = "<="
        Case dboComparisonNotEqualTo
            strCompare = "<>"
        Case dboComparisonLike
            strCompare = "LIKE"
        Case dboComparisonNotLike
            strCompare = "NOT LIKE"
        Case Else
            RaiseError dboErrorGeneral, "Invalid SQLComparisonOperatorEnum value " & eCompare
    End Select
    
    SQLConvertCompare = strCompare
    
End Function

Public Function SQLConvertLogicalOperator( _
    ByVal eLogicalOperator As SQLLogicalOperatorEnum) As String
    
    Dim strLogicalOperator As String

    Select Case eLogicalOperator
        Case dboLogicalAnd
            strLogicalOperator = "AND"
        Case dboLogicalOr
            strLogicalOperator = "OR"
    End Select
    
    SQLConvertLogicalOperator = strLogicalOperator
    
End Function

Public Function SQLFieldNameAndTablePrefix( _
    ByVal objTable As SQLSelectTable, _
    ByVal strFieldName As String, _
    ByVal eConnectionType As ConnectionTypeEnum) As String
    
    Dim strTablePrefix As String
    
    If Not objTable Is Nothing Then
        'If Trim$(objTable.Alias) = vbNullString Then
            strTablePrefix = objTable.Name
        'Else
        '    strTablePrefix = objTable.Alias
        'End If
        strTablePrefix = SQLConvertIdentifierName(strTablePrefix, eConnectionType) & "."
    End If
    
    SQLFieldNameAndTablePrefix = strTablePrefix & SQLConvertIdentifierName(strFieldName, eConnectionType)
    
End Function

'Must copy the value into the variant because sometimes it will require the use of the 'Set' keyword
Public Sub SQLConditionValue(ByVal vValue As Variant, ByRef vCopyInto As Variant)
 
    Select Case VarType(vValue)
        Case vbObject
            If vValue Is Nothing Then
                Set vCopyInto = Nothing
            ElseIf TypeOf vValue Is SQLFieldValue Then
                Dim objSQLFieldValue As SQLFieldValue
                Set objSQLFieldValue = vValue
                Set vCopyInto = objSQLFieldValue.Value
            Else
                RaiseError dboErrorGeneral, "Invalid Object Type"
            End If
        Case vbArray, vbDataObject, vbEmpty, vbError, vbUserDefinedType, vbVariant
            RaiseError dboErrorGeneral, "Invalid Data-Type"
        Case Else
            vCopyInto = vValue
    End Select
 
End Sub
Public Sub CompareValuePairAssertValid( _
    ByVal eCompare As SQLComparisonOperatorEnum, _
    ByRef vValue As Variant)
    
    If VarType(vValue) <> vbString And (eCompare = dboComparisonLike Or eCompare = dboComparisonNotLike) Then
        RaiseError dboErrorGeneral, "The LIKE operator cannot be used in conjunction with a non-string data type"
    ElseIf VarType(vValue) = vbBoolean And Not (eCompare = dboComparisonEqualTo Or eCompare = dboComparisonNotEqualTo) Then
        RaiseError dboErrorGeneral, "A boolean value can only be used in conjunction with the dboComparisonEqualTo or dboComparisonNotEqualTo operators"
    End If
    
End Sub

Public Sub SQLConvertBooleanValue( _
    ByRef vValue As Variant, _
    ByRef eCompare As SQLComparisonOperatorEnum)
    
    'If a boolean variable set to true then use the opposite
    'operator and compare it to 0. ie. if the condition is 'field = true' then
    'SQL code should be 'field <> 0'
    '-1 is true in MSAccess and 1 is true in SQLServer.

    If VarType(vValue) = vbBoolean Then
        If vValue = True Then
            If eCompare = dboComparisonEqualTo Then
                eCompare = dboComparisonNotEqualTo
            Else
                eCompare = dboComparisonEqualTo
            End If
            vValue = False
        End If
    End If
    
End Sub

Public Function CollectionRemoveItem( _
    ByVal objCollection As Collection, _
    ByVal objItem As Object) As Boolean
    
    Dim intIndex As Integer
    
    For intIndex = 1 To objCollection.Count
        If objCollection(intIndex) Is objItem Then
            objCollection.Remove intIndex
            CollectionRemoveItem = True
            Exit For
        End If
    Next

End Function

Public Sub RaiseError( _
    ByVal eError As ErrorEnum, _
    Optional ByVal strExtra As String)

    Select Case eError
        Case dboErrorGeneral: Err.Raise dboErrorGeneral, , strExtra
        Case dboErrorIndexOutOfBounds: Err.Raise dboErrorIndexOutOfBounds, , "Index out of bounds " & strExtra
        Case dboErrorNotIntegerOrString: Err.Raise dboErrorNotIntegerOrString, , "Invalid data type, expected Integer or String"
        Case dboErrorObjectIsNothing: Err.Raise dboErrorObjectIsNothing, , "Object is Nothing"
        Case dboErrorObjectAlreadyExists: Err.Raise dboErrorObjectAlreadyExists, , "Object already exists " & strExtra
        Case dboErrorObjectDoesNotExist: Err.Raise dboErrorObjectDoesNotExist, , "Object does not exist " & strExtra
        Case dboErrorInvalidPropertyValue: Err.Raise dboErrorInvalidPropertyValue, , "Invalid property value " & strExtra
        Case dboErrorInvalidArgument: Err.Raise dboErrorInvalidArgument, , "Invalid argument " & strExtra
        Case dboErrorObjectNotDeletable: Err.Raise dboErrorObjectNotDeletable, , "Object is not deletable " & strExtra
        Case dboErrorObjectNotSaved: Err.Raise dboErrorObjectNotSaved, , "Objects not saved " & strExtra
        Case dboErrorNotSupported: Err.Raise dboErrorNotSupported, , "Method or Property not supported " & strExtra
        Case dboErrorMethodOrPropertyLocked: Err.Raise dboErrorMethodOrPropertyLocked, , "Method or Property locked " & strExtra
    End Select

End Sub
