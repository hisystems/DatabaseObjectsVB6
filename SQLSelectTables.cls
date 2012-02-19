VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectTables"
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

Private pobjJoins As SQLSelectTableJoins
Private pcolTables As Collection

Public Function Add( _
    Optional ByVal strTableName As String) As SQLSelectTable

    Dim objTable As SQLSelectTable

    Set objTable = New SQLSelectTable
    objTable.Name = strTableName
    'objTable.Alias = strAlias
    
    pcolTables.Add objTable
    
    Set Add = objTable

End Function

Public Property Get Item(ByVal vIndex As Variant) As SQLSelectTable
Attribute Item.VB_UserMemId = 0

    Select Case VarType(vIndex)
    Case vbInteger, vbLong
        Set Item = pcolTables(vIndex)
    Case vbString
        Set Item = pcolTables(TableNameIndex(vIndex))
    Case Else
        RaiseError dboErrorNotIntegerOrString
    End Select
    
End Property

Public Function Exists(ByVal strFieldName As String) As Boolean

    Exists = TableNameIndex(strFieldName) <> 0

End Function

Public Sub Delete(ByRef objTable As SQLSelectTable)
    
    If Not CollectionRemoveItem(pcolTables, objTable) Then
        RaiseError dboErrorObjectDoesNotExist
    End If
    Set objTable = Nothing

End Sub

Public Property Get Count() As Integer

    Count = pcolTables.Count
    
End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Set Enumerator = pcolTables.[_NewEnum]

End Property

Public Property Get Joins() As SQLSelectTableJoins

    Set Joins = pobjJoins

End Property

Public Property Set Joins(ByVal objValue As SQLSelectTableJoins)

    Set pobjJoins = objValue

End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim intIndex As Integer
    Dim strSQL As String
    Dim bAddTable As Boolean
    
    For intIndex = 1 To Me.Count
        'Include the table if it's not being used in a join
        If pobjJoins Is Nothing Then
            bAddTable = True
        ElseIf Not pobjJoins.Exists(Me.Item(intIndex)) Then
            bAddTable = True
        Else
            bAddTable = False
        End If
        
        If bAddTable Then
            strSQL = strSQL & Me.Item(intIndex).SQL(eConnectionType)
            If intIndex <> Me.Count Then
                strSQL = strSQL & ", "
            End If
        End If
    Next
        
    If Not pobjJoins Is Nothing Then
        If pobjJoins.SQL(eConnectionType) <> vbNullString And strSQL <> vbNullString Then
            strSQL = strSQL & " "
        End If
        strSQL = strSQL & pobjJoins.SQL(eConnectionType)
    End If
    
    SQL = strSQL
    
End Property

Private Sub Class_Initialize()
    
    Set pcolTables = New Collection
    Set pobjJoins = New SQLSelectTableJoins
    
End Sub

Private Function TableNameIndex( _
    ByVal strTableName As String) As Integer

    Dim intIndex As Integer
    Dim objTable As SQLSelectTable

    strTableName = Trim$(strTableName)

    For intIndex = 1 To Me.Count
        Set objTable = pcolTables(intIndex)
        If StrComp(objTable.Name, strTableName, vbTextCompare) = 0 Then
            TableNameIndex = intIndex
            Exit For
        End If
    Next

End Function
