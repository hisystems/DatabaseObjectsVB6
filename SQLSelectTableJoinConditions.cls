VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectTableJoinConditions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
Private pcolConditions As Collection
Private pobjParent As SQLSelectTableJoin

Friend Property Set Parent(ByVal objValue As SQLSelectTableJoin)

    Set pobjParent = objValue

End Property

Friend Property Get Parent() As SQLSelectTableJoin

    Set Parent = pobjParent

End Property

Public Function Add( _
    Optional ByVal strLeftTableFieldName As String, _
    Optional ByVal eCompare As SQLComparisonOperatorEnum = dboComparisonEqualTo, _
    Optional ByVal strRightTableFieldName As String) As SQLSelectTableJoinCondition

    'Add the AND operator if an operator hasn't been called after the previous Add call
    If pcolLogicalOperators.Count < pcolConditions.Count Then
        Me.AddLogicalOperator dboLogicalAnd
    End If
    
    If eCompare = dboComparisonLike Or eCompare = dboComparisonNotLike Then
        RaiseError dboErrorGeneral, "LIKE operator is not supported for table joins."
    End If

    Set Add = New SQLSelectTableJoinCondition
    Set Add.Parent = Me
    
    Add.LeftTableFieldName = strLeftTableFieldName
    Add.Compare = eCompare
    Add.RightTableFieldName = strRightTableFieldName

    pcolConditions.Add Add

End Function

Public Function AddLogicalOperator( _
    Optional ByVal eLogicalOperator As SQLLogicalOperatorEnum = dboLogicalAnd)

    If pcolLogicalOperators.Count + 1 > pcolConditions.Count Then
        RaiseError dboErrorGeneral, "First call the Add function - this function has been called without a prior call to Add"
    End If
    
    pcolLogicalOperators.Add eLogicalOperator
    
End Function

Public Property Get Item(ByVal intIndex As Integer) As SQLSelectTableJoinCondition
Attribute Item.VB_UserMemId = 0

    Set Item = pcolConditions(intIndex)

End Property

Public Property Get Count() As Integer
    
    Count = pcolConditions.Count

End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Set Enumerator = pcolConditions.[_NewEnum]

End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim intIndex As Integer
    Dim objCondition As SQLSelectTableJoinCondition
    Dim strSQL As String
    
    With pcolConditions
        For intIndex = 1 To .Count
            If intIndex > 1 Then
                strSQL = strSQL & _
                    " " & SQLConvertLogicalOperator(pcolLogicalOperators(intIndex - 1)) & " "
            End If
            
            Set objCondition = .Item(intIndex)
            strSQL = strSQL & objCondition.SQL(eConnectionType)
        Next
    End With
    
    SQL = strSQL

End Property

Private Sub Class_Initialize()

    Set pcolConditions = New Collection
    Set pcolLogicalOperators = New Collection
    
End Sub
