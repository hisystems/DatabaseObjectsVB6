VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectOrderByFields"
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

Option Explicit

'Ascending and Descending are 0 and -1 so that
'the VB Not operator will negate the ordering from ascending to descending
'and visa versa.
'ie.  'Not dboOrderAscending' will equal dboOrderDescending
Public Enum SQLOrderByEnum
    dboOrderNone = 1
    dboOrderAscending = 0
    dboOrderDescending = -1
End Enum

Private pcolOrderByFields As Collection

Public Function Add( _
    Optional ByVal strFieldName As String, _
    Optional ByVal eOrder As SQLOrderByEnum = dboOrderAscending, _
    Optional ByVal objTable As SQLSelectTable, _
    Optional ByVal eAggregate As SQLAggregateFunctionEnum = 0) As SQLSelectOrderByField
    
    Dim objFieldOrder As SQLSelectOrderByField
    Set objFieldOrder = New SQLSelectOrderByField
    
    With objFieldOrder
        Set .Table = objTable
        .Name = strFieldName
        .Order = eOrder
        .AggregateFunction = eAggregate
    End With
    
    pcolOrderByFields.Add objFieldOrder
    
    Set Add = objFieldOrder
    
End Function

Public Property Get Item(ByVal vIndex As Variant) As SQLSelectOrderByField
Attribute Item.VB_UserMemId = 0

    Select Case VarType(vIndex)
    Case vbInteger, vbLong
        Set Item = pcolOrderByFields(vIndex)
    Case vbString
        Set Item = pcolOrderByFields(FieldNameIndex(vIndex))
    Case Else
        RaiseError dboErrorNotIntegerOrString
    End Select
    
End Property

Public Function Exists(ByVal strFieldName As String) As Boolean

    Exists = FieldNameIndex(strFieldName) <> 0

End Function

Public Sub Delete(ByRef objOrderByField As SQLSelectOrderByField)
    
    If Not CollectionRemoveItem(pcolOrderByFields, objOrderByField) Then
        RaiseError dboErrorObjectDoesNotExist
    End If
    Set objOrderByField = Nothing
    
End Sub

Public Property Get Count() As Integer

    Count = pcolOrderByFields.Count
    
End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Set Enumerator = pcolOrderByFields.[_NewEnum]

End Property

Private Sub Class_Initialize()

    Set pcolOrderByFields = New Collection

End Sub

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim intIndex As Integer
    Dim strSQL As String
    
    For intIndex = 1 To Me.Count
        strSQL = strSQL & Me.Item(intIndex).SQL(eConnectionType)
        If intIndex <> Me.Count Then
            strSQL = strSQL & ", "
        End If
    Next

    SQL = strSQL

End Property

Private Function FieldNameIndex( _
    ByVal strFieldName As String) As Integer

    Dim intIndex As Integer
    Dim objOrderByField As SQLSelectOrderByField

    strFieldName = Trim$(strFieldName)

    For intIndex = 1 To Me.Count
        Set objOrderByField = pcolOrderByFields(intIndex)
        If StrComp(objOrderByField.Name, strFieldName, vbTextCompare) = 0 Then
            FieldNameIndex = intIndex
            Exit For
        End If
    Next

End Function
