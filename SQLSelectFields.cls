VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectFields"
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

Private pcolFieldNames As Collection

Public Function Add( _
    Optional ByVal strFieldName As String, _
    Optional ByVal strAlias As String, _
    Optional ByVal eAggregateFunction As SQLAggregateFunctionEnum = 0, _
    Optional ByVal objTable As SQLSelectTable) As SQLSelectField
    
    Dim objSQLField As SQLSelectField
    
    Set objSQLField = New SQLSelectField
    
    With objSQLField
        Set .Table = objTable
        .Name = strFieldName
        .Alias = strAlias
        .AggregateFunction = eAggregateFunction
    End With

    pcolFieldNames.Add objSQLField

    Set Add = objSQLField

End Function

Public Sub AddItems( _
    ParamArray vFieldNames() As Variant)
    
    Dim intIndex As Integer
    
    For intIndex = LBound(vFieldNames) To UBound(vFieldNames)
        If VarType(vFieldNames(intIndex)) = vbString Then
            Me.Add vFieldNames(intIndex)
        Else
            RaiseError dboErrorGeneral, "Invalid data type, expected string."
        End If
    Next
    
End Sub

Public Property Get Item(ByVal vIndex As Variant) As SQLSelectField
Attribute Item.VB_UserMemId = 0

    Select Case VarType(vIndex)
    Case vbInteger, vbLong
        Set Item = pcolFieldNames(vIndex)
    Case vbString
        Set Item = pcolFieldNames(FieldNameIndex(vIndex))
    Case Else
        RaiseError dboErrorNotIntegerOrString
    End Select
    
End Property

Public Function Exists(ByVal strFieldName As String) As Boolean

    Exists = FieldNameIndex(strFieldName) <> 0

End Function

Public Sub Delete(ByRef objSelectField As SQLSelectField)

    If Not CollectionRemoveItem(pcolFieldNames, objSelectField) Then
        RaiseError dboErrorObjectDoesNotExist
    End If
    Set objSelectField = Nothing
    
End Sub

Public Property Get Count() As Integer

    Count = pcolFieldNames.Count
    
End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Set Enumerator = pcolFieldNames.[_NewEnum]

End Property

Private Sub Class_Initialize()

    Set pcolFieldNames = New Collection

End Sub

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim intIndex As Integer
    Dim strSQL As String

    If Me.Count = 0 Then
        strSQL = "*"
    Else
        For intIndex = 1 To Me.Count
            strSQL = strSQL & Me.Item(intIndex).SQL(eConnectionType)
            If intIndex <> Me.Count Then
                strSQL = strSQL & ", "
            End If
        Next
    End If

    SQL = strSQL

End Property

Private Function FieldNameIndex( _
    ByVal strFieldName As String) As Integer
    
    Dim intIndex As Integer
    
    strFieldName = Trim$(strFieldName)
    
    For intIndex = 1 To Me.Count
        If StrComp(Me.Item(intIndex).Name, strFieldName, vbTextCompare) = 0 Then
            FieldNameIndex = intIndex
            Exit For
        End If
    Next
    
End Function
