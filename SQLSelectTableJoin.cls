VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectTableJoin"
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

Private pobjLeftTable As SQLSelectTable
Private pobjRightTable As SQLSelectTable
Private pobjConditions As SQLSelectTableJoinConditions
Private pobjParent As SQLSelectTableJoins

Public TheType As SQLJoinEnum

Friend Property Set Parent(ByVal objParent As SQLSelectTableJoins)
    
    Set pobjParent = objParent

End Property

Public Property Set LeftTable(ByVal objValue As SQLSelectTable)

    If objValue Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If

    Set pobjLeftTable = objValue

End Property

Public Property Get LeftTable() As SQLSelectTable
    
    Set LeftTable = pobjLeftTable
    
End Property

Public Property Set RightTable(ByVal objValue As SQLSelectTable)

    If objValue Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If

    Set pobjRightTable = objValue

End Property

Public Property Get RightTable() As SQLSelectTable
    
    Set RightTable = pobjRightTable
    
End Property

Public Property Get Where() As SQLSelectTableJoinConditions

    Set Where = pobjConditions

End Property

Public Property Set Where(ByVal objValue As SQLSelectTableJoinConditions)

    Set pobjConditions = objValue

End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim strSQL As String
    Dim strJoin As String

    Select Case Me.TheType
        Case dboInnerJoin
            strJoin = "INNER JOIN"
        Case dboFullOuterJoin
            strJoin = "FULL OUTER JOIN"
        Case dboLeftOuterJoin
            strJoin = "LEFT OUTER JOIN"
        Case dboRightOuterJoin
            strJoin = "RIGHT OUTER JOIN"
    End Select

    If Index = 1 Then
        strSQL = pobjLeftTable.SQL(eConnectionType)
    Else
        strSQL = pobjParent(Index - 1).SQL(eConnectionType)
    End If

    strSQL = strSQL & " " & strJoin & " " & pobjRightTable.SQL(eConnectionType)
    
    If Not pobjConditions Is Nothing Then
        If pobjConditions.SQL(eConnectionType) <> vbNullString Then
            strSQL = strSQL & " ON " & pobjConditions.SQL(eConnectionType)
        End If
    End If
    
    'Surround the join with parentheses - MS Access won't accept is otherwise
    SQL = "(" & strSQL & ")"

End Property

Private Function Index() As Integer

    Dim intIndex As Integer
    
    For intIndex = 1 To pobjParent.Count
        If pobjParent(intIndex) Is Me Then
            Exit For
        End If
    Next

    Index = intIndex

End Function

Private Sub Class_Initialize()

    Set pobjConditions = New SQLSelectTableJoinConditions
    Set pobjConditions.Parent = Me

End Sub
