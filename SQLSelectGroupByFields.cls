VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectGroupByFields"
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

Private pcolGroupByFields As Collection

Public Function Add( _
    Optional ByVal strFieldName As String, _
    Optional ByVal objTable As SQLSelectTable) As SQLSelectGroupByField
    
    Dim objFieldOrder As SQLSelectGroupByField
    Set objFieldOrder = New SQLSelectGroupByField
    
    With objFieldOrder
        Set .Table = objTable
        .Name = strFieldName
    End With
    
    pcolGroupByFields.Add objFieldOrder
    
    Set Add = objFieldOrder
    
End Function

Public Property Get Item(ByVal vIndex As Variant) As SQLSelectGroupByField
Attribute Item.VB_UserMemId = 0

    Select Case VarType(vIndex)
    Case vbInteger, vbLong
        Set Item = pcolGroupByFields(vIndex)
    Case vbString
        Set Item = pcolGroupByFields(FieldNameIndex(vIndex))
    Case Else
        RaiseError dboErrorNotIntegerOrString
    End Select
    
End Property

Public Function Exists(ByVal strFieldName As String) As Boolean

    Exists = FieldNameIndex(strFieldName) <> 0

End Function

Public Sub Delete(ByRef objGroupByField As SQLSelectGroupByField)
    
    If Not CollectionRemoveItem(pcolGroupByFields, objGroupByField) Then
        RaiseError dboErrorObjectDoesNotExist
    End If
    Set objGroupByField = Nothing
    
End Sub

Public Property Get Count() As Integer

    Count = pcolGroupByFields.Count
    
End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Set Enumerator = pcolGroupByFields.[_NewEnum]

End Property

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

Private Sub Class_Initialize()

    Set pcolGroupByFields = New Collection

End Sub

Private Function FieldNameIndex( _
    ByVal strFieldName As String) As Integer

    Dim intIndex As Integer
    Dim objGroupByField As SQLSelectGroupByField

    strFieldName = Trim$(strFieldName)

    For intIndex = 1 To Me.Count
        Set objGroupByField = pcolGroupByFields(intIndex)
        If StrComp(objGroupByField.Name, strFieldName, vbTextCompare) = 0 Then
            FieldNameIndex = intIndex
            Exit For
        End If
    Next

End Function


