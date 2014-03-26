VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLConditions"
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

Private pcolLogicalOperators As Collection
Private pcolSQLConditions As Collection

Public Function Add( _
    Optional ByVal strFieldName As String, _
    Optional ByVal eCompare As SQLComparisonOperatorEnum = dboComparisonEqualTo, _
    Optional ByVal vValue As Variant, _
    Optional ByVal objTable As SQLSelectTable) As SQLCondition

    EnsurePreviousLogicalOperatorExists
    
    Set Add = New SQLCondition
    
    Set Add.Table = objTable
    Add.FieldName = strFieldName
    Add.Compare = eCompare
    If Not IsMissing(vValue) Then Add.Value = vValue

    pcolSQLConditions.Add Add

End Function

Public Sub AddExpression(ByVal objExpression As ISQLExpression)

    If objExpression Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If
    
    EnsurePreviousLogicalOperatorExists
    pcolSQLConditions.Add objExpression

End Sub

Public Sub AddCompareExpression( _
    ByVal objLeft As ISQLExpression, _
    ByVal eCompare As SQLComparisonOperatorEnum, _
    ByVal objRightExpression As ISQLExpression)
    
    Dim objCompareExpression As SQLComparisonExpression
    Set objCompareExpression = New SQLComparisonExpression
    
    Set objCompareExpression.Left = objLeft
    objCompareExpression.Compare = eCompare
    Set objCompareExpression.Right = objRightExpression
    
    AddExpression objCompareExpression
    
End Sub

Public Sub AddCondition( _
    ByVal objCondition As SQLCondition)
    
    If objCondition Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If
    
    EnsurePreviousLogicalOperatorExists
    pcolSQLConditions.Add objCondition
    
End Sub

Public Sub AddConditions( _
    ByVal objConditions As SQLConditions)
    
    If objConditions.IsEmpty Then
        RaiseError dboErrorInvalidArgument, "SQLConditions does not contain any conditions"
    End If
    
    EnsurePreviousLogicalOperatorExists
    pcolSQLConditions.Add objConditions
    
End Sub

Public Function AddInSelect( _
    Optional ByVal strFieldName As String, _
    Optional ByVal objSelect As SQLSelect, _
    Optional ByVal objTable As SQLSelectTable) As SQLConditionInSelect
    
    EnsurePreviousLogicalOperatorExists
    
    Set AddInSelect = New SQLConditionInSelect
    Set AddInSelect.Table = objTable
    AddInSelect.FieldName = strFieldName
    Set AddInSelect.SelectSet = objSelect

    pcolSQLConditions.Add AddInSelect

End Function

Public Function AddSelect( _
    Optional ByVal objSelect As SQLSelect, _
    Optional ByVal eCompare As SQLComparisonOperatorEnum = dboComparisonEqualTo, _
    Optional ByVal vValue As Variant) As SQLConditionSelect
    
    EnsurePreviousLogicalOperatorExists
    
    Set AddSelect = New SQLConditionSelect
    Set AddSelect.SelectSet = objSelect
    AddSelect.Compare = eCompare
    If Not IsMissing(vValue) Then AddSelect.Value = vValue
    
    pcolSQLConditions.Add AddSelect
     
End Function

Public Function AddFieldCompare( _
    Optional ByVal objTable1 As SQLSelectTable, _
    Optional ByVal strFieldName1 As String, _
    Optional ByVal eCompare As SQLComparisonOperatorEnum, _
    Optional ByVal objTable2 As SQLSelectTable, _
    Optional ByVal strFieldName2 As String) As SQLConditionFieldCompare
    
    EnsurePreviousLogicalOperatorExists
    
    Set AddFieldCompare = New SQLConditionFieldCompare
    
    With AddFieldCompare
        Set .Table1 = objTable1
        .FieldName1 = strFieldName1
        .Compare = eCompare
        Set .Table2 = objTable2
        .FieldName2 = strFieldName2
    End With

    pcolSQLConditions.Add AddFieldCompare

End Function
    
Private Sub EnsurePreviousLogicalOperatorExists()

    'Add the AND operator if an operator hasn't been called after the previous Add or AddConditions call
    If pcolLogicalOperators.Count < pcolSQLConditions.Count Then
        Me.AddLogicalOperator dboLogicalAnd
    End If

End Sub

Public Sub AddLogicalOperator( _
    Optional ByVal eLogicalOperator As SQLLogicalOperatorEnum = dboLogicalAnd)

    If pcolLogicalOperators.Count + 1 > pcolSQLConditions.Count Then
        RaiseError dboErrorGeneral, "First call the Add function - this function has been called without a prior call to Add"
    End If
    
    pcolLogicalOperators.Add eLogicalOperator
    
End Sub

Public Property Get IsEmpty() As Boolean

    IsEmpty = pcolSQLConditions.Count = 0

End Property
    
Private Function LogicalOperator( _
    ByVal intIndex As Integer) As SQLLogicalOperatorEnum
    
    LogicalOperator = pcolLogicalOperators(intIndex)
    
End Function

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim intIndex As Integer
    Dim strSQL As String
    Dim objSQLCondition As SQLCondition
    Dim objSQLConditions As SQLConditions
    Dim objSelectIn As SQLConditionInSelect
    Dim objSelect As SQLConditionSelect
    Dim objFields As SQLConditionFieldCompare
    Dim objExpression As ISQLExpression
    
    With pcolSQLConditions
        For intIndex = 1 To .Count
            If intIndex > 1 Then
                strSQL = strSQL & " " & SQLConvertLogicalOperator(pcolLogicalOperators(intIndex - 1)) & " "
            End If
            If TypeOf .Item(intIndex) Is SQLCondition Then
                Set objSQLCondition = .Item(intIndex)
                strSQL = strSQL & objSQLCondition.SQL(eConnectionType)
            ElseIf TypeOf .Item(intIndex) Is SQLConditions Then
                Set objSQLConditions = .Item(intIndex)
                strSQL = strSQL & "(" & objSQLConditions.SQL(eConnectionType) & ")"
            ElseIf TypeOf .Item(intIndex) Is SQLConditionSelect Then
                Set objSelect = .Item(intIndex)
                strSQL = strSQL & objSelect.SQL(eConnectionType)
            ElseIf TypeOf .Item(intIndex) Is SQLConditionFieldCompare Then
                Set objFields = .Item(intIndex)
                strSQL = strSQL & objFields.SQL(eConnectionType)
            ElseIf TypeOf .Item(intIndex) Is SQLConditionInSelect Then
                Set objSelectIn = .Item(intIndex)
                strSQL = strSQL & objSelectIn.SQL(eConnectionType)
            ElseIf TypeOf .Item(intIndex) Is ISQLExpression Then
                Set objExpression = .Item(intIndex)
                strSQL = strSQL & objExpression.SQL(eConnectionType)
            End If
        Next
    End With

    SQL = strSQL

End Property

Private Sub Class_Initialize()

    Set pcolSQLConditions = New Collection
    Set pcolLogicalOperators = New Collection
    
End Sub

