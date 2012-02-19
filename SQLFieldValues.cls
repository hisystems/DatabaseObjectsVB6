VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLFieldValues"
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

Private pcolFields As Collection

Public Function Add( _
    Optional ByVal strName As String, _
    Optional ByVal vEqualToValue As Variant = vbEmpty) As SQLFieldValue
    
    Dim objSQLFieldValue As SQLFieldValue
    Set objSQLFieldValue = New SQLFieldValue
    
    With objSQLFieldValue
        .Name = strName
        .Value = vEqualToValue
    End With
    
    pcolFields.Add objSQLFieldValue
    
    Set Add = objSQLFieldValue
    
End Function

Public Property Get Item(ByVal vIndex As Variant) As SQLFieldValue
Attribute Item.VB_UserMemId = 0

    Dim objSQLFieldValue As SQLFieldValue
    
    Select Case VarType(vIndex)
        Case vbString
            Set objSQLFieldValue = pcolFields(FieldNameIndex(vIndex))
        Case Else
            Set objSQLFieldValue = pcolFields(vIndex)
    End Select
    
    Set Item = objSQLFieldValue
    
End Property

Public Function Exists(ByVal strFieldName As String) As Boolean

    Exists = FieldNameIndex(strFieldName) <> 0

End Function

Public Sub Delete(ByRef objFieldValue As SQLFieldValue)

    If Not CollectionRemoveItem(pcolFields, objFieldValue) Then
        RaiseError dboErrorObjectDoesNotExist
    End If
    
    Set objFieldValue = Nothing
    
End Sub

Public Property Get Count() As Integer

    Count = pcolFields.Count
    
End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Set Enumerator = pcolFields.[_NewEnum]

End Property

Private Sub Class_Initialize()

    Set pcolFields = New Collection

End Sub

Friend Function FieldNameIndex( _
    ByVal strFieldName As String) As Integer
    
    Dim intIndex As Integer
    Dim objSQLFieldValue As SQLFieldValue
    
    strFieldName = Trim$(strFieldName)
    
    For intIndex = 1 To Me.Count
        Set objSQLFieldValue = pcolFields(intIndex)
        If StrComp(objSQLFieldValue.Name, strFieldName, vbTextCompare) = 0 Then
            FieldNameIndex = intIndex
            Exit For
        End If
    Next
    
End Function
